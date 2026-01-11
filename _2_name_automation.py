"""
ticket_verification_batch.py
FIXED: Use journey_date for routing (not arrival_date)
If journey_date >= 2026-02-13 → DEPARTURE, write journey_date
"""

import os
import re
from typing import List, Dict, Tuple, Optional
from datetime import date as _date

from rapidfuzz import fuzz
from dateutil import parser as dparser
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv()

# -------------------- CONFIG --------------------
SHEET_ID = os.getenv("SHEET_ID")
TICKET_SHEET = os.getenv("SHEET_NAME", "Sheet13")
MASTER_SHEET = "Master"
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json")

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
sheets_service = build("sheets", "v4", credentials=creds)

MATCH_THRESHOLD = 85

# Date routing: journey_date >= 2026-02-13 means DEPARTURE
DEPARTURE_DATE = _date(2026, 2, 13)


class NameMatcher:
    """Partial matching for short master names"""
    
    @staticmethod
    def normalize_name(name: str) -> str:
        """Remove titles, special chars, normalize spaces"""
        if not name:
            return ""
        
        name = name.upper()
        name = re.sub(r"\b(KR|KUMAR|DEVI|SHRI|SMT|MR|MRS|MS|MISS|SRI)\b", "", name)
        name = re.sub(r"[^A-Z ]", "", name)
        name = re.sub(r"\s+", " ", name).strip()
        return name
    
    @staticmethod
    def match_against_master(passenger_name: str, master_names: List[Dict]) -> Tuple[List[Dict], int]:
        """Use partial_ratio for substring matching"""
        if not passenger_name or not master_names:
            return [], 0
        
        norm_passenger = NameMatcher.normalize_name(passenger_name)
        
        matches = []
        best_score = 0
        
        for master in master_names:
            partial_score = fuzz.partial_ratio(norm_passenger, master["norm"])
            token_score = fuzz.token_sort_ratio(norm_passenger, master["norm"])
            final_score = max(partial_score, token_score)
            
            if final_score >= MATCH_THRESHOLD:
                matches.append({**master, "score": final_score})
                if final_score > best_score:
                    best_score = final_score
        
        matches.sort(key=lambda x: x["score"], reverse=True)
        return matches, best_score


class DateHelper:
    """Date parsing and routing"""
    
    @staticmethod
    def format_mmddyy(date_str: str) -> str:
        """Normalize any date into MM/DD/YY"""
        if not date_str:
            return ""
        try:
            dt = dparser.parse(date_str, dayfirst=False)
            return dt.strftime("%m/%d/%y")
        except Exception:
            return date_str
    
    @staticmethod
    def is_departure_date(date_str: str) -> bool:
        """Check if date >= 2026-02-13 (departure cutoff)"""
        if not date_str:
            return False
        try:
            dt = dparser.parse(date_str, dayfirst=False).date()
            return dt >= DEPARTURE_DATE  # >= not ==
        except Exception:
            return False


class SheetManager:
    """Google Sheets operations"""
    
    @staticmethod
    def read_master_names() -> List[Dict]:
        """Load master guest list with context"""
        print(f"\n📘 Loading master names from '{MASTER_SHEET}' sheet")
        
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{MASTER_SHEET}!A2:D"
        ).execute()
        
        rows = resp.get("values", [])
        master_names = []
        
        for i, row in enumerate(rows, start=2):
            if row and row[0].strip():
                raw_name = row[0].strip()
                place = row[2] if len(row) > 2 else ""
                venue = row[3] if len(row) > 3 else ""
                
                master_names.append({
                    "raw": raw_name,
                    "norm": NameMatcher.normalize_name(raw_name),
                    "row_no": i,
                    "place": place,
                    "venue": venue
                })
        
        print(f"✅ Loaded {len(master_names)} master names")
        return master_names
    
    @staticmethod
    def read_ticket_sheet() -> List[Dict]:
        """Load all tickets"""
        print(f"\n📄 Loading tickets from '{TICKET_SHEET}' sheet")
        
        resp = sheets_service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{TICKET_SHEET}!A2:Q"
        ).execute()
        
        rows = resp.get("values", [])
        tickets = []
        
        for idx, row in enumerate(rows, start=2):
            if len(row) < 8:
                continue
            
            tickets.append({
                "row_no": idx,
                "journey_date": row[0] if len(row) > 0 else "",
                "departure_time": row[1] if len(row) > 1 else "",
                "arrival_date": row[2] if len(row) > 2 else "",
                "arrival_time": row[3] if len(row) > 3 else "",
                "mode": row[4] if len(row) > 4 else "",
                "seat": row[5] if len(row) > 5 else "",
                "details": row[6] if len(row) > 6 else "",
                "name": row[7] if len(row) > 7 else "",
                "train_number": row[8] if len(row) > 8 else "",
                "train_name": row[9] if len(row) > 9 else "",
                "status": row[10] if len(row) > 10 else "",
                "pnr": row[11] if len(row) > 11 else "",
                "source": row[12] if len(row) > 12 else "",
                "suggested": row[13] if len(row) > 13 else "",
                "score": row[14] if len(row) > 14 else "",
                "approved": row[15] if len(row) > 15 else "",
                "commit_status": row[16] if len(row) > 16 else ""
            })
        
        print(f"✅ Loaded {len(tickets)} tickets")
        return tickets
    
    @staticmethod
    def write_match_result(row_no: int, suggested_name: str, score: int):
        """Write match suggestion to columns N-O"""
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{TICKET_SHEET}!N{row_no}:O{row_no}",
            valueInputOption="RAW",
            body={"values": [[suggested_name, score]]}
        ).execute()
    
    @staticmethod
    def write_approved_name(row_no: int, approved_name: str):
        """Write approved name to column P"""
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{TICKET_SHEET}!P{row_no}",
            valueInputOption="RAW",
            body={"values": [[approved_name]]}
        ).execute()
    
    @staticmethod
    def write_status(row_no: int, status: str):
        """Write status to column Q"""
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{TICKET_SHEET}!Q{row_no}",
            valueInputOption="RAW",
            body={"values": [[status]]}
        ).execute()
    
    @staticmethod
    def batch_update_master(updates: List[Dict]):
        """Batch update Master sheet"""
        if not updates:
            return
        
        print(f"\n✍️  Batch updating {len(updates)} records in Master sheet...")
        
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": updates
            }
        ).execute()
        
        print(f"✅ Updated {len(updates)} master records")
    
    @staticmethod
    def batch_update_commit_status(row_updates: List[Tuple[int, str]]):
        """Batch update commit status"""
        if not row_updates:
            return
        
        data = []
        for row_no, status in row_updates:
            data.append({
                "range": f"{TICKET_SHEET}!Q{row_no}",
                "values": [[status]]
            })
        
        sheets_service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={
                "valueInputOption": "RAW",
                "data": data
            }
        ).execute()


class VerificationWorkflow:
    """Batch workflow with partial matching"""
    
    def __init__(self):
        self.sheet_mgr = SheetManager()
        self.matcher = NameMatcher()
        self.date_helper = DateHelper()
    
    def step1_match_and_suggest(self):
        """Step 1: Match with partial_ratio"""
        print("\n" + "="*80)
        print("STEP 1: AUTO-MATCH PASSENGERS")
        print("="*80)
        
        master_names = self.sheet_mgr.read_master_names()
        tickets = self.sheet_mgr.read_ticket_sheet()
        
        print(f"\n🔍 Matching {len(tickets)} passengers...")
        
        matched = 0
        unmatched = 0
        duplicates = 0
        
        for ticket in tickets:
            passenger_name = ticket["name"]
            row_no = ticket["row_no"]

            if ticket.get("commit_status") == "COMMITTED":
                continue
            
            if ticket["mode"] == "ERROR":
                print(f"⏭️  Row {row_no}: Skipping ERROR row")
                continue
            
            if not passenger_name:
                print(f"⏭️  Row {row_no}: No passenger name")
                continue
            
            print(f"\n📍 Row {row_no} | {passenger_name} ({ticket['mode']})")
            
            matches, best_score = self.matcher.match_against_master(passenger_name, master_names)
            
            if len(matches) > 1:
                if len(matches) >= 2 and (matches[0]["score"] - matches[1]["score"]) <= 3:
                    print(f"   ⚠️  DUPLICATE: Found {len(matches)} similar matches:")
                    for m in matches[:3]:
                        print(f"      - {m['raw']} (row {m['row_no']}) [{m['score']:.0f}] - {m['place']} {m['venue']}")
                    
                    self.sheet_mgr.write_match_result(
                        row_no, 
                        f"DUPLICATE: {matches[0]['raw']}", 
                        best_score
                    )
                    duplicates += 1
                else:
                    match = matches[0]
                    self.sheet_mgr.write_match_result(row_no, match["raw"], best_score)
                    print(f"   ✅ Matched: {match['raw']} (row {match['row_no']}) [{best_score:.0f}]")
                    matched += 1
                
            elif len(matches) == 1:
                match = matches[0]
                self.sheet_mgr.write_match_result(row_no, match["raw"], best_score)
                print(f"   ✅ Matched: {match['raw']} (row {match['row_no']}) [{best_score:.0f}]")
                matched += 1
                
            else:
                self.sheet_mgr.write_match_result(row_no, "", best_score)
                print(f"   ❌ No match (best score={best_score:.0f})")
                unmatched += 1
        
        print("\n" + "="*80)
        print(f"✅ STEP 1 COMPLETE")
        print(f"   → Matched: {matched}")
        print(f"   → Duplicates: {duplicates}")
        print(f"   → Unmatched: {unmatched}")
        print(f"\n📌 NEXT: Review column N, correct in column P if needed")
        print(f"   Then run: python test.py 2")
        print("="*80)
    
    def step2_autofill_and_commit(self):
        """Step 2: Auto-fill and commit"""
        print("\n" + "="*80)
        print("STEP 2: AUTO-FILL & COMMIT")
        print("="*80)
        
        master_names = self.sheet_mgr.read_master_names()
        tickets = self.sheet_mgr.read_ticket_sheet()
        
        print(f"\n📝 Processing {len(tickets)} tickets...\n")
        
        # Phase 1: Auto-fill
        print("Phase 1: Auto-filling...")
        autofilled = 0
        
        for ticket in tickets:
            row_no = ticket["row_no"]
            suggested = ticket["suggested"]
            approved = ticket["approved"]
            
            if ticket["mode"] == "ERROR":
                continue
            
            if "DUPLICATE" in suggested:
                print(f"⏭️  Row {row_no}: Duplicate - needs manual input")
                continue
            
            if not approved and suggested:
                self.sheet_mgr.write_approved_name(row_no, suggested)
                print(f"✓ Row {row_no}: '{suggested}'")
                ticket["approved"] = suggested
                autofilled += 1
        
        print(f"\n✅ Auto-filled {autofilled} rows\n")
        
        # Reload
        tickets = self.sheet_mgr.read_ticket_sheet()
        
        # Phase 2: Commit
        print("Phase 2: Committing...")
        
        master_updates = []
        status_updates = []
        committed = 0
        skipped = 0
        errors = 0
        
        for ticket in tickets:
            row_no = ticket["row_no"]
            approved_name = ticket["approved"]
            journey_date = ticket["journey_date"]
            arrival_date = ticket["arrival_date"]
            departure_time = ticket["departure_time"]
            arrival_time = ticket["arrival_time"]
            seat = ticket["seat"]
            train_number = ticket["train_number"]
            train_name = ticket["train_name"]
            mode = ticket["mode"]
            
            if not approved_name or mode == "ERROR":
                skipped += 1
                continue
            
            if ticket["commit_status"] == "COMMITTED":
                skipped += 1
                continue
            
            # Find exact match
            match = None
            for m in master_names:
                if m["raw"].upper() == approved_name.upper():
                    match = m
                    break
            
            if not match:
                status_updates.append((row_no, f"ERROR: '{approved_name}' not found"))
                print(f"❌ Row {row_no}: '{approved_name}' NOT FOUND")
                errors += 1
                continue
            
            master_row = match["row_no"]
            
            # FIXED ROUTING LOGIC
            # Use journey_date to decide ARRIVAL vs DEPARTURE
            # journey_date >= 2026-02-13 → DEPARTURE
            # journey_date < 2026-02-13 → ARRIVAL
            
            is_departure = self.date_helper.is_departure_date(journey_date)
            
            if is_departure:
                # DEPARTURE: Write journey_date
                date_to_write = journey_date
                trip_type = "DEPARTURE"
                
                formatted_date = self.date_helper.format_mmddyy(date_to_write)
                time_value = departure_time if departure_time else arrival_time
                
                # DEPARTURE COLUMNS (AE-AJ)
                master_updates.extend([
                    {"range": f"{MASTER_SHEET}!AE{master_row}", "values": [[formatted_date]]},
                    {"range": f"{MASTER_SHEET}!AF{master_row}", "values": [[mode]]},
                    {"range": f"{MASTER_SHEET}!AG{master_row}", "values": [[seat]]},
                    {"range": f"{MASTER_SHEET}!AH{master_row}", "values": [[train_number]]},
                    {"range": f"{MASTER_SHEET}!AI{master_row}", "values": [[train_name]]},
                    {"range": f"{MASTER_SHEET}!AJ{master_row}", "values": [[time_value]]},
                ])
            else:
                # ARRIVAL: Write journey_date (or arrival_date for flights if available)
                if mode == "FLIGHT":
                    date_to_write = journey_date
                else:
                    date_to_write = arrival_date
                
                trip_type = "ARRIVAL"
                
                formatted_date = self.date_helper.format_mmddyy(date_to_write)
                time_value = arrival_time if arrival_time else departure_time
                
                # ARRIVAL COLUMNS (I-N)
                master_updates.extend([
                    {"range": f"{MASTER_SHEET}!I{master_row}", "values": [[formatted_date]]},
                    {"range": f"{MASTER_SHEET}!J{master_row}", "values": [[mode]]},
                    {"range": f"{MASTER_SHEET}!K{master_row}", "values": [[seat]]},
                    {"range": f"{MASTER_SHEET}!L{master_row}", "values": [[train_number]]},
                    {"range": f"{MASTER_SHEET}!M{master_row}", "values": [[train_name]]},
                    {"range": f"{MASTER_SHEET}!N{master_row}", "values": [[time_value]]},
                ])
            
            status_updates.append((row_no, "COMMITTED"))
            print(f"✅ [{mode} {trip_type}] {approved_name} → Master!{master_row} | {formatted_date} {train_number}")
            committed += 1
        
        if master_updates:
            self.sheet_mgr.batch_update_master(master_updates)
        
        if status_updates:
            self.sheet_mgr.batch_update_commit_status(status_updates)
        
        print("\n" + "="*80)
        print(f"✅ STEP 2 COMPLETE")
        print(f"   → Committed: {committed}")
        print(f"   → Skipped: {skipped}")
        print(f"   → Errors: {errors}")
        print("="*80)
    
    def run(self):
        """Run both steps"""
        self.step1_match_and_suggest()
        input("\n⏸️  Review column N, correct in P if needed. Press Enter...")
        self.step2_autofill_and_commit()


if __name__ == "__main__":
    import sys
    
    workflow = VerificationWorkflow()
    
    if len(sys.argv) > 1:
        step = sys.argv[1]
        if step == "1":
            workflow.step1_match_and_suggest()
        elif step == "2":
            workflow.step2_autofill_and_commit()
        else:
            print("Usage: python test.py [1|2]")
    else:
        workflow.run()
