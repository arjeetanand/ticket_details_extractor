"""
Wedding Ticket Automation - PRODUCTION FINAL
Handles all formats + adds error rows to sheet
"""

import io
import os
import re
import json
import requests
from typing import List, Dict, Optional, Tuple
from datetime import datetime

import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter
import pytesseract
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

load_dotenv()

# ================== CONFIG ==================
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
PROCESSED_FOLDER_ID = os.getenv("PROCESSED_FOLDER_ID")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "Sheet1")
SERVICE_ACCOUNT_FILE = os.getenv("SERVICE_ACCOUNT_FILE", "service_account.json")

RAPIDAPI_KEY = os.getenv("RAPIDAPI_KEY")
RAPIDAPI_HOST = os.getenv("RAPIDAPI_HOST", "irctc-indian-railway-pnr-status.p.rapidapi.com")

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
drive_service = build("drive", "v3", credentials=creds)
sheets_service = build("sheets", "v4", credentials=creds)


class PNRExtractor:
    """Extract PNR and passenger info from tickets"""
    
    @staticmethod
    def file_to_images(file_bytes: bytes, mime_type: str, filename: str) -> List[Image.Image]:
        """Convert PDF/image to PIL images"""
        images = []
        if 'pdf' in mime_type.lower() or filename.lower().endswith('.pdf'):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page in doc:
                pix = page.get_pixmap(dpi=300)
                png_bytes = pix.tobytes("png")
                img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
                images.append(img)
            doc.close()
        else:
            img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            images.append(img)
        return images
    
    @staticmethod
    def ocr_images_full(images: List[Image.Image]) -> str:
        """Full OCR with enhanced quality - IMPROVED VERSION"""
        texts = []
        for idx, img in enumerate(images):
            # Try multiple OCR strategies
            strategies = [
                ("Standard", img, "--psm 6"),
                ("High contrast", PNRExtractor._enhance_image(img, contrast=3.0), "--psm 6"),
                ("Auto", img, "--psm 3"),
            ]
            
            best_text = ""
            for strategy_name, processed_img, psm in strategies:
                try:
                    gray = processed_img.convert("L")
                    gray = gray.resize((gray.width * 3, gray.height * 3), Image.LANCZOS)
                    
                    text = pytesseract.image_to_string(
                        gray, lang="eng",
                        config=f"{psm} -c preserve_interword_spaces=1"
                    )
                    
                    # Use the longest/best result
                    if len(text) > len(best_text):
                        best_text = text
                except:
                    continue
            
            print(f"\n{'='*30}\nOCR OUTPUT — PAGE {idx + 1}\n{'='*30}")
            print(best_text[:500] + "..." if len(best_text) > 500 else best_text)
            print("=" * 30 + "\n")
            
            texts.append(best_text)
        return "\n".join(texts)
    
    @staticmethod
    def _enhance_image(img: Image.Image, contrast: float = 2.5) -> Image.Image:
        """Enhance image quality"""
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(contrast)
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)
        return img
    
    @staticmethod
    def detect_ticket_type(text: str) -> str:
        """Detect TRAIN or FLIGHT"""
        t = text.lower()
        
        # Check for non-ticket content
        non_ticket_words = ['invoice', 'bill', 'receipt', 'buyer', 'seller', 'gstin', 'tea store']
        if any(word in t for word in non_ticket_words):
            # Check if it has ticket keywords
            ticket_words = ['pnr', 'train no', 'passenger', 'booking']
            if not any(word in t for word in ticket_words):
                return "UNKNOWN"
        
        flight_kw = ['indigo', '6e', 'flight', 'terminal', 'check-in', 'boarding pass', 
                     'cleartrip', 'makemytrip', 'booking confirmed']
        train_kw = ['irctc', 'electronic reservation', 'train no', 'coach', 'berth',
                    'quota', 'vananchal', 'kriya yoga', 'vande bharat', 'booking id : tk', 
                    'pnr :', 'passenger status', 'current status']
        
        flight_score = sum(1 for kw in flight_kw if kw in t)
        train_score = sum(1 for kw in train_kw if kw in t)
        
        print(f"🔍 Detection: Flight={flight_score}, Train={train_score}")
        
        if flight_score > train_score and flight_score >= 2:
            return "FLIGHT"
        elif train_score > 0:
            return "TRAIN"
        return "UNKNOWN"
    

    
    @staticmethod
    def extract_pnr_from_text(text: str) -> Optional[str]:
        """Extract PNR (10 digits for train, 6 alphanumeric for flight)"""
        # Explicit PNR label

        def extract_spaced_pnr(text):
            # Example: "6 5 6 2 5 2 6 4 9 6"
            spaced = re.findall(r'(?:\d\s+){9}\d', text)
            for s in spaced:
                pnr = re.sub(r'\s+', '', s)
                if len(pnr) == 10:
                    return pnr
            return None
        pnr = extract_spaced_pnr(text)
        if pnr:
            return pnr

        m = re.search(r"PNR[:=\s]*([A-Z0-9]{6,10})\b", text, re.IGNORECASE)
        if m:
            return m.group(1).upper()
        
        # 10-digit train PNR
        candidates = re.findall(r"\b\d{10}\b", text)
        candidates = [c for c in candidates if not c.startswith(("201", "202", "203", "982", "100"))]
        if candidates:
            return candidates[0]
        
        return None
    
    @staticmethod
    def extract_train_passengers(text: str) -> List[str]:
        """Extract passenger names from train tickets (ALL FORMATS) - IMPROVED"""
        names = []
        
        BLOCKLIST = {
            "CHECK TIMINGS", "PASSENGER DETAILS", "ELECTRONIC RESERVATION",
            "BOOKED FROM", "BOARDING AT", "TRANSACTION ID", "ACRONYMS",
            "NAME", "AGE", "GENDER", "BOOKING STATUS", "CURRENT STATUS",
            "PASSENGER STATUS", "COACH", "SEAT", "BERTH", "CHART NOT PREPARED",
            "CATERING SERVICE"
        }
        
        lines = [re.sub(r"\s+", " ", l.strip()) for l in text.splitlines() if l.strip()]
        
        for line in lines:
            # Skip blocklist
            if any(b in line.upper() for b in BLOCKLIST):
                continue
            if len(line) < 4:  # Changed from 10 to catch short names like "SONAL"
                continue
            
            # Pattern 1: IRCTC format "1. NAME AGE GENDER | CNF"
            m = re.search(
                r"^\d+\.\s+([A-Z!|I][A-Z!\s|I]+?)\s+(\d{1,3})\s+([MFmf|I])\s+[\|\s]*(CNF|WL|RAC|VEG)",
                line, re.IGNORECASE
            )
            if m:
                raw_name = PNRExtractor._clean_ocr_name(m.group(1))
                if raw_name and raw_name not in names:
                    names.append(raw_name)
                    print(f"  → Train passenger (IRCTC): {raw_name}")
                continue
            
            # Pattern 2: ixigo format "1. NAME, AGE, GENDER"
            m2 = re.search(
                r"^\d+\.\s+([A-Z][A-Za-z\s]+?),\s*(\d{1,3})\s*,\s*([MF])\s*$",
                line, re.IGNORECASE
            )
            if m2:
                raw_name = PNRExtractor._clean_ocr_name(m2.group(1))
                if raw_name and raw_name not in names:
                    names.append(raw_name)
                    print(f"  → Train passenger (ixigo): {raw_name}")
                continue
            
            # Pattern 3: App screenshot format "NAME" followed by "Male | AGE yrs"
            if re.match(r'^[A-Z][A-Z\s]{3,40}$', line):  # Changed from 5 to 3 to catch "SONAL"
                # Check if it's a proper name
                words = line.split()
                # Accept single names or multi-word names
                if len(words) >= 1 and all(len(w) >= 2 for w in words):  # Changed from >= 2 to >= 1
                    raw_name = PNRExtractor._clean_ocr_name(line)
                    if raw_name and raw_name not in names:
                        names.append(raw_name)
                        print(f"  → Train passenger (App): {raw_name}")
                        continue
        
        return names
    
    @staticmethod
    def _clean_ocr_name(raw_name: str) -> str:
        """Clean OCR artifacts from name"""
        name = raw_name.strip()
        name = name.replace("!", "I").replace("|", "I")
        name = re.sub(r'\s+', ' ', name)
        
        # Skip if contains numbers or too short
        if re.search(r'\d', name) or len(name) < 3:  # Changed from < 3 to accept short names
            return ""
        
        # Skip common OCR garbage
        BAD_WORDS = ['CONFIRMED', 'AVAILABLE', 'BOOKING', 'STATUS', 'PASSENGER', 'OPTION']
        if any(bad in name.upper() for bad in BAD_WORDS):
            return ""
        
        # Fix concatenated: "ANILSANTHALIA" → "Anil Santhalia"
        if " " not in name and len(name) > 8:
            parts = re.findall(r'[A-Z][a-z]+', name)
            name = " ".join(parts) if len(parts) >= 2 else name.title()
        else:
            name = name.title()
        
        return name
    
    @staticmethod
    def extract_flight_passengers(text: str) -> List[Dict]:
        """Extract flight passengers with validation"""
        passengers = []
        
        # Stricter patterns
        patterns = [
            r"\b(M[rs]s?\.?\s+[A-Z][A-Z\s]+?)\s*\(ADULT\)",  # Ms NAME (ADULT)
            r"\b(M[rs]\.?\s+[A-Z][a-z]+\s+[A-Z][a-z]+)\b",   # Mr Firstname Lastname
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                name = re.sub(r'\s+', ' ', match.strip())
                name = name.replace("Mr.", "Mr").replace("Ms.", "Ms").replace("Mrs.", "Mrs")
                
                # Validation: reject garbage
                if PNRExtractor._is_valid_flight_name(name):
                    # Check if already added
                    if not any(p['name'].upper() == name.upper() for p in passengers):
                        passengers.append({'name': name, 'seat': ''})
                        print(f"  → Flight passenger: {name}")
        
        return passengers
    
    @staticmethod
    def _is_valid_flight_name(name: str) -> bool:
        """Validate flight passenger name"""
        name_upper = name.upper()
        
        # Reject common garbage
        BAD_WORDS = [
            'ALLOWED', 'ITEMS', 'BAGGAGE', 'BOOKING', 'DETAILS', 'PAYMENT',
            'TICKET', 'FLIGHT', 'TERMINAL', 'CHECK', 'INFORMATION', 'IMPORTANT',
            'CONTACT', 'CUSTOMER', 'SUPPORT', 'YATRA', 'DIGI', 'AVOID'
        ]
        
        if any(bad in name_upper for bad in BAD_WORDS):
            return False
        
        # Must have at least one space (first + last name)
        if ' ' not in name:
            return False
        
        # Must be between 5-40 characters
        if len(name) < 5 or len(name) > 40:
            return False
        
        # Must not have numbers
        if re.search(r'\d', name):
            return False
        
        return True


class FlightExtractor:
    """Extract flight details"""
    
    @staticmethod
    def extract_flight_details(text: str) -> Dict:
        """Extract all flight info from ticket"""
        result = {
            'airline': '',
            'flight_number': '',
            'pnr': '',
            'route': '',
            'date': '',
            'departure_time': '',
            'arrival_time': '',
            'passengers': []
        }
        
        # Airline
        airlines = {
            'indigo': 'IndiGo', '6e': 'IndiGo',
            'air india express': 'Air India Express',
            'air india': 'Air India',
            'vistara': 'Vistara',
            'spicejet': 'SpiceJet'
        }
        
        t_lower = text.lower()
        for key, airline in airlines.items():
            if key in t_lower:
                result['airline'] = airline
                break
        
        # Flight number
        patterns = [
            r'\b(IX|I5|AI|6E|UK|SG|QP)\s*[-\s]*(\d{3,4})\b',  # Standard
            r'Flight\s+(?:No\.?|Number)?\s*[:=\s]*([A-Z0-9]{2,3}[-\s]?\d{3,4})',  # "Flight: IX 1234"
        ]
        
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if len(m.groups()) == 2:
                result['flight_number'] = f"{m.group(1)} {m.group(2)}"
            else:
                result['flight_number'] = re.sub(r'\s+', ' ', m.group(1))
            break
        

        # m = re.search(r'\b(6E|AI|UK|SG|QP|IX)\s*(\d{3,4})\b', text, re.IGNORECASE)
        # if m:
        #     result['flight_number'] = f"{m.group(1)} {m.group(2)}"
        
        # PNR
        m = re.search(r'PNR[:=\s]*([A-Z0-9]{6})\b', text, re.IGNORECASE)
        if m:
            result['pnr'] = m.group(1).upper()
        
        # Route
        m = re.search(r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s*[-→–]\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)', text)
        if m:
            result['route'] = f"{m.group(1)} → {m.group(2)}"
        
        # Date
        date_patterns = [
            r'\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+[,\']?(\d{2,4})\b',
            r'\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,?\s+(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})\b',
            r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b'
        ]
        for pattern in date_patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                result['date'] = m.group(0)
                break
        
        # Times
        times = re.findall(r'\b(\d{1,2}:\d{2})\s*(?:hrs|HRS)?\b', text)
        if len(times) >= 2:
            result['departure_time'] = times[0]
            result['arrival_time'] = times[1]
        
        # Passengers
        result['passengers'] = PNRExtractor.extract_flight_passengers(text)
        
        print(f"\n✈️ Extracted: {result['airline']} {result['flight_number']} | {result['route']} | {len(result['passengers'])} pax")
        
        return result


class TrainPNRAPI:
    """Train PNR API integration"""
    
    @staticmethod
    def check_train_pnr(pnr: str) -> Optional[Dict]:
        """Check train PNR via API"""
        url = f"https://{RAPIDAPI_HOST}/getPNRStatus/{pnr}"
        headers = {
            "x-rapidapi-key": RAPIDAPI_KEY,
            "x-rapidapi-host": RAPIDAPI_HOST
        }
        
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                print(f"✓ API fetched PNR {pnr}")
                return data
            else:
                print(f"✗ API error {response.status_code}")
                return None
        except Exception as e:
            print(f"✗ API exception: {e}")
            return None
    
    @staticmethod
    def parse_train_response(response: Dict, passenger_names: List[str] = None) -> Dict:
        """Parse train API response with separate time columns"""
        if not response or not response.get('success'):
            return {
                'pnr': 'UNKNOWN',
                'mode': 'TRAIN',
                'error': 'API failed',
                'passengers': []
            }
        
        data = response.get('data', {})
        if not data:
            return {'pnr': 'UNKNOWN', 'mode': 'TRAIN', 'error': 'No data', 'passengers': []}
        
        try:
            pnr = data.get('pnrNumber', '')
            train_number = data.get('trainNumber', '')
            train_name = data.get('trainName', '')
            
            # Extract date and time from "dateOfJourney": "Feb 13, 2026 4:25:00 PM"
            journey_date_str = data.get('dateOfJourney', '')
            arrival_date_str = data.get('arrivalDate', '')
            
            journey_date = ''
            departure_time = ''
            arrival_date = ''
            arrival_time = ''
            
            # Parse journey date and time
            if journey_date_str:
                try:
                    from dateutil import parser as dparser
                    dt = dparser.parse(journey_date_str)
                    journey_date = dt.strftime("%Y-%m-%d")
                    departure_time = dt.strftime("%H:%M")
                except:
                    journey_date = journey_date_str
                    # Try to extract time manually
                    time_match = re.search(r'(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?', journey_date_str)
                    if time_match:
                        hour = int(time_match.group(1))
                        minute = time_match.group(2)
                        ampm = time_match.group(4)
                        if ampm == 'PM' and hour != 12:
                            hour += 12
                        elif ampm == 'AM' and hour == 12:
                            hour = 0
                        departure_time = f"{hour:02d}:{minute}"
            
            # Parse arrival date and time
            if arrival_date_str:
                try:
                    from dateutil import parser as dparser
                    dt = dparser.parse(arrival_date_str)
                    arrival_date = dt.strftime("%Y-%m-%d")
                    arrival_time = dt.strftime("%H:%M")
                except:
                    arrival_date = arrival_date_str
                    time_match = re.search(r'(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?', arrival_date_str)
                    if time_match:
                        hour = int(time_match.group(1))
                        minute = time_match.group(2)
                        ampm = time_match.group(4)
                        if ampm == 'PM' and hour != 12:
                            hour += 12
                        elif ampm == 'AM' and hour == 12:
                            hour = 0
                        arrival_time = f"{hour:02d}:{minute}"
            
            # Passengers
            passengers = []
            passenger_list = data.get('passengerList', [])
            
            for idx, p in enumerate(passenger_list):
                pax_name = passenger_names[idx] if passenger_names and idx < len(passenger_names) else f"Passenger {idx+1}"
                
                passengers.append({
                    'name': pax_name,
                    'coach': p.get('currentCoachId', p.get('bookingCoachId', '')),
                    'seat': str(p.get('currentBerthNo', p.get('bookingBerthNo', '0'))),
                    'berth': p.get('currentBerthCode', p.get('bookingBerthCode', '')),
                    'status': p.get('currentStatusDetails', '')
                })
            
            print(f"✓ Parsed {len(passengers)} passengers | Dep: {departure_time} | Arr: {arrival_time}")
            
            return {
                'pnr': pnr,
                'mode': 'TRAIN',
                'train_number': train_number,
                'train_name': train_name,
                'journey_date': journey_date,
                'departure_time': departure_time,
                'arrival_date': arrival_date,
                'arrival_time': arrival_time,
                'passengers': passengers,
                'details': f"{train_number} / {train_name}",
                'error': None
            }
        
        except Exception as e:
            print(f"✗ Parse error: {e}")
            import traceback
            traceback.print_exc()
            return {'pnr': 'UNKNOWN', 'mode': 'TRAIN', 'error': str(e), 'passengers': []}


class TicketProcessor:
    """Main ticket processor"""
    
    def __init__(self):
        self.pnr_extractor = PNRExtractor()
        self.train_api = TrainPNRAPI()
        self.flight_extractor = FlightExtractor()
    
    def process_ticket(self, file_bytes: bytes, filename: str, mime_type: str) -> Dict:
        """Process ticket file"""
        print(f"\n{'='*70}\nProcessing: {filename}\n{'='*70}")
        
        try:
            # OCR
            images = self.pnr_extractor.file_to_images(file_bytes, mime_type, filename)
            print(f"✓ Converted to {len(images)} image(s)")
            
            full_text = self.pnr_extractor.ocr_images_full(images)
            
            # Detect type
            ticket_type = self.pnr_extractor.detect_ticket_type(full_text)
            print(f"✓ Detected: {ticket_type}")
            
            if ticket_type == "FLIGHT":
                return self.process_flight(full_text, filename)
            elif ticket_type == "TRAIN":
                return self.process_train(full_text, filename)
            else:
                return {
                    'error': 'Unknown ticket type (not a valid train/flight ticket)',
                    'filename': filename,
                    'mode': 'ERROR'
                }
        except Exception as e:
            return {
                'error': f'Processing failed: {str(e)}',
                'filename': filename,
                'mode': 'ERROR'
            }
    
    def process_flight(self, text: str, filename: str) -> Dict:
        """Process flight ticket"""
        print("\n✈️ Processing FLIGHT...")
        
        flight_data = self.flight_extractor.extract_flight_details(text)
        
        if not flight_data['passengers']:
            return {
                'error': 'No passengers found in flight ticket',
                'filename': filename,
                'mode': 'ERROR'
            }
        
        return {
            'mode': 'FLIGHT',
            'airline': flight_data['airline'],
            'flight_number': flight_data['flight_number'],
            'pnr': flight_data['pnr'],
            'route': flight_data['route'],
            'date': flight_data['date'],
            'departure_time': flight_data['departure_time'],
            'arrival_time': flight_data['arrival_time'],
            'passengers': flight_data['passengers'],
            'details': f"{flight_data['airline']} {flight_data['flight_number']}",
            'filename': filename,
            'error': None
        }
    
    def process_train(self, text: str, filename: str) -> Dict:
        """Process train ticket"""
        print("\n🚂 Processing TRAIN...")
        
        # Get PNR
        pnr = self.pnr_extractor.extract_pnr_from_text(text)
        if not pnr:
            return {
                'error': 'PNR not found in ticket',
                'filename': filename,
                'mode': 'ERROR'
            }
        
        print(f"✓ PNR: {pnr}")
        
        # API call
        api_response = self.train_api.check_train_pnr(pnr)
        if not api_response:
            return {
                'error': f'API failed for PNR {pnr}',
                'filename': filename,
                'pnr': pnr,
                'mode': 'ERROR'
            }
        
        # Extract names
        passenger_names = self.pnr_extractor.extract_train_passengers(text)
        
        # Parse
        result = self.train_api.parse_train_response(api_response, passenger_names)
        result['filename'] = filename
        
        return result


class GoogleSheetsManager:
    """Sheet management with error rows"""
    
    @staticmethod
    def append_ticket_data(ticket_data: Dict) -> bool:
        """Append to sheet (including error rows)"""
        if ticket_data.get('error'):
            # Add error row
            return GoogleSheetsManager._append_error_row(ticket_data)
        
        mode = ticket_data.get('mode')
        
        if mode == 'TRAIN':
            return GoogleSheetsManager._append_train(ticket_data)
        elif mode == 'FLIGHT':
            return GoogleSheetsManager._append_flight(ticket_data)
        
        return False
    
    @staticmethod
    def _append_error_row(ticket_data: Dict) -> bool:
        """Append error row to sheet"""
        error_msg = ticket_data.get('error', 'Unknown error')
        filename = ticket_data.get('filename', 'Unknown file')
        pnr = ticket_data.get('pnr', '')
        
        print(f"⚠️  Adding ERROR ROW: {error_msg}")
        
        # Create blank row with error info in last columns
        row = [
            '',  # A: Journey Date
            '',  # B: Departure Time
            '',  # C: Arrival Date
            '',  # D: Arrival Time
            'ERROR',  # E: Mode
            '',  # F: Seat
            '',  # G: Details
            '',  # H: Name
            '',  # I: Number
            '',  # J: Train/Airline
            f'ERROR: {error_msg}',  # K: Status/Route (error message)
            pnr,  # L: PNR
            filename  # M: Source
        ]
        
        return GoogleSheetsManager._append_rows([row])
    
    @staticmethod
    def _append_train(ticket_data: Dict) -> bool:
        """Append train data with separate time columns"""
        passengers = ticket_data.get('passengers', [])
        if not passengers:
            return False
        
        rows = []
        for pax in passengers:
            seat = f"{pax.get('coach', '')}/{pax.get('seat', '')}/{pax.get('berth', '')}"
            row = [
                ticket_data.get('journey_date', ''),       # A: Journey Date
                ticket_data.get('departure_time', ''),     # B: Departure Time
                ticket_data.get('arrival_date', ''),       # C: Arrival Date
                ticket_data.get('arrival_time', ''),       # D: Arrival Time
                'TRAIN',                                   # E: Mode
                seat,                                      # F: Seat
                ticket_data.get('details', ''),            # G: Train Details
                pax.get('name', ''),                       # H: Name
                ticket_data.get('train_number', ''),       # I: Train#
                ticket_data.get('train_name', ''),         # J: Train Name
                pax.get('status', ''),                     # K: Status
                ticket_data.get('pnr', ''),                # L: PNR
                ticket_data.get('filename', '')            # M: Source
            ]
            rows.append(row)
        
        return GoogleSheetsManager._append_rows(rows)
    
    @staticmethod
    def _append_flight(ticket_data: Dict) -> bool:
        """Append flight data with separate time columns"""
        passengers = ticket_data.get('passengers', [])
        if not passengers:
            return False
        
        rows = []
        for pax in passengers:
            row = [
                ticket_data.get('date', ''),               # A: Date
                ticket_data.get('departure_time', ''),     # B: Departure Time
                '',                                        # C: Arrival Date (same day)
                ticket_data.get('arrival_time', ''),       # D: Arrival Time
                'FLIGHT',                                  # E: Mode
                pax.get('seat', ''),                       # F: Seat
                ticket_data.get('details', ''),            # G: Flight Details
                pax.get('name', ''),                       # H: Name
                ticket_data.get('flight_number', ''),      # I: Flight#
                ticket_data.get('airline', ''),            # J: Airline
                ticket_data.get('route', ''),              # K: Route
                ticket_data.get('pnr', ''),                # L: PNR
                ticket_data.get('filename', '')            # M: Source
            ]
            rows.append(row)
        
        return GoogleSheetsManager._append_rows(rows)
    
    @staticmethod
    def _append_rows(rows: List) -> bool:
        """Append rows"""
        try:
            body = {"values": rows}
            sheets_service.spreadsheets().values().append(
                spreadsheetId=SHEET_ID,
                range=f"{SHEET_NAME}!A:M",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body=body
            ).execute()
            
            print(f"✓ Appended {len(rows)} rows")
            return True
        except Exception as e:
            print(f"✗ Sheet error: {e}")
            return False


class DriveManager:
    """Drive operations"""
    
    @staticmethod
    def list_files(folder_id: str) -> List[Dict]:
        """List files"""
        q = f"'{folder_id}' in parents and trashed = false"
        files = []
        page_token = None
        
        while True:
            resp = drive_service.files().list(
                q=q, spaces='drive',
                fields="nextPageToken, files(id, name, mimeType, createdTime)",
                pageToken=page_token
            ).execute()
            
            files.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        
        return files
    
    @staticmethod
    def download_file(file_id: str) -> bytes:
        """Download file"""
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        
        while not done:
            _, done = downloader.next_chunk()
        
        fh.seek(0)
        return fh.read()
    
    @staticmethod
    def move_file(file_id: str, dest_folder_id: str, source_folder_id: str):
        """Move file"""
        drive_service.files().update(
            fileId=file_id,
            addParents=dest_folder_id,
            removeParents=source_folder_id,
            fields="id, parents"
        ).execute()


class WeddingTicketAutomation:
    """Main automation"""
    
    def __init__(self):
        self.processor = TicketProcessor()
        self.sheets = GoogleSheetsManager()
        self.drive = DriveManager()
    
    def run(self):
        """Execute automation"""
        print("\n" + "="*70)
        print("WEDDING TICKET AUTOMATION - PRODUCTION FINAL")
        print("With error handling + separate time columns")
        print("="*70)
        
        files = self.drive.list_files(DRIVE_FOLDER_ID)
        print(f"\n📁 Found {len(files)} files")
        
        if not files:
            print("No files to process")
            return
        
        success = 0
        errors = 0
        error_files = []
        
        for file_meta in files:
            try:
                file_id = file_meta['id']
                filename = file_meta['name']
                mime_type = file_meta.get('mimeType', '')
                
                file_bytes = self.drive.download_file(file_id)
                ticket_data = self.processor.process_ticket(file_bytes, filename, mime_type)
                
                # Always append (including errors)
                self.sheets.append_ticket_data(ticket_data)
                
                if ticket_data.get('error'):
                    errors += 1
                    error_files.append(f"{filename}: {ticket_data.get('error')}")
                else:
                    success += 1
                    
                    if PROCESSED_FOLDER_ID and len(PROCESSED_FOLDER_ID) > 5:
                        try:
                            self.drive.move_file(file_id, PROCESSED_FOLDER_ID, DRIVE_FOLDER_ID)
                            print("✓ Moved\n")
                        except:
                            print("⚠ Move skipped\n")
                    else:
                        print("⚠ Move skipped\n")
            
            except Exception as e:
                print(f"✗ Critical error: {e}\n")
                import traceback
                traceback.print_exc()
                errors += 1
                error_files.append(f"{filename}: Critical error - {str(e)}")
        
        print("\n" + "="*70)
        print(f"SUMMARY: {success} ✓ SUCCESS | {errors} ✗ ERRORS")
        print("="*70)
        
        if error_files:
            print("\n❌ ERROR FILES:")
            for err in error_files:
                print(f"  • {err}")
        print()


if __name__ == "__main__":
    automation = WeddingTicketAutomation()
    automation.run()
