# 🎫 Wedding Ticket Automation System

A production-ready automation pipeline to process wedding guest travel tickets (Train & Flight), extract passenger details using OCR + APIs, reconcile them against a master guest list, and automatically update Google Sheets with arrival and departure logistics.

This system is built for real-world Indian train and flight tickets, handling noisy OCR, multiple formats, duplicates, and failures gracefully.

---

## ✨ Key Features

### End-to-End Automation

* Reads tickets directly from Google Drive
* Supports PDF and image formats
* Writes structured data into Google Sheets
* Automatically moves processed files

### Intelligent OCR & Extraction

* Multi-pass OCR with image enhancement
* Robust PNR extraction (including spaced digits)
* Handles:

  * IRCTC tickets
  * ixigo screenshots
  * Flight boarding passes
* Supports multiple passengers per ticket

### Train Ticket Handling

* Live PNR verification using RapidAPI
* Extracts:

  * Train number and name
  * Journey date
  * Arrival and departure time (separate columns)
  * Coach / berth / seat
  * Passenger status

### Flight Ticket Handling

* Auto-detects airline and flight number
* Extracts route, date, and times
* Validates passenger names to avoid OCR garbage

### Error-Tolerant by Design

* Invalid or unknown files are still logged
* Errors are written as rows in the sheet
* No silent failures

---

## 📁 Project Structure

```
.
├── ticket_verification_batch.py   # Name matching + master sheet commit
├── wedding_ticket_automation.py   # OCR + extraction + Drive automation
├── service_account.json           # Google service account credentials
├── .env                           # Environment variables
├── README.md
```

---

## 🔄 Overall Workflow

### Phase 1 — Ticket Ingestion & Extraction

1. Read files from Google Drive
2. Convert PDFs/images to images
3. Perform enhanced OCR
4. Detect ticket type (TRAIN / FLIGHT)
5. Extract passenger names, PNR, dates, times, and seat details
6. Append rows to Ticket Sheet
7. Move processed files to another Drive folder

### Phase 2 — Name Matching & Verification

1. Load Master Guest List
2. Perform fuzzy matching using partial ratio and token sort ratio
3. Write suggested names and confidence scores
4. Flag duplicates for manual review

### Phase 3 — Auto-Fill & Commit

1. Auto-fill approved names
2. Route data based on journey date
3. Update Arrival or Departure columns in Master Sheet
4. Mark rows as COMMITTED

---

## 🗓️ Arrival vs Departure Logic

```
journey_date >= 2026-02-13 → DEPARTURE
journey_date <  2026-02-13 → ARRIVAL
```

* Uses journey_date only (not arrival date)
* Ensures correct wedding logistics routing

---

## 📊 Google Sheets Schema

### Ticket Sheet

| Column | Description                   |
| ------ | ----------------------------- |
| A      | Journey Date                  |
| B      | Departure Time                |
| C      | Arrival Date                  |
| D      | Arrival Time                  |
| E      | Mode (TRAIN / FLIGHT / ERROR) |
| F      | Seat                          |
| G      | Details                       |
| H      | Passenger Name                |
| I      | Train / Flight Number         |
| J      | Train Name / Airline          |
| K      | Status / Route                |
| L      | PNR                           |
| M      | Source Filename               |
| N      | Suggested Name                |
| O      | Match Score                   |
| P      | Approved Name                 |
| Q      | Commit Status                 |

### Master Sheet

* Arrival Columns: `I → N`
* Departure Columns: `AE → AJ`

---

## ⚙️ Environment Variables

Create a `.env` file:

```env
SHEET_ID=your_google_sheet_id
SHEET_NAME=Sheet13
MASTER_SHEET=Master
SERVICE_ACCOUNT_FILE=service_account.json

DRIVE_FOLDER_ID=source_drive_folder_id
PROCESSED_FOLDER_ID=processed_drive_folder_id

RAPIDAPI_KEY=your_rapidapi_key
RAPIDAPI_HOST=irctc-indian-railway-pnr-status.p.rapidapi.com
```

---

## ▶️ How to Run

### Step 1 — Process Tickets

```bash
python wedding_ticket_automation.py
```

### Step 2 — Auto-Match Names

```bash
python ticket_verification_batch.py 1
```

### Step 3 — Commit to Master Sheet

```bash
python ticket_verification_batch.py 2
```

---

## 🧪 Edge Cases Handled

* Spaced PNR digits
* OCR artifacts (`|` → `I`, `!` → `I`)
* Single-word names (e.g., SONAL)
* Duplicate fuzzy matches
* Invoice or receipt false positives
* Partial API failures

---

## 🧠 Tech Stack

* Python 3.9+
* Google Drive API
* Google Sheets API
* PyMuPDF
* Tesseract OCR
* RapidFuzz
* RapidAPI (IRCTC PNR)
* Pillow

---

## 🚀 Production Notes

* Idempotent commits (no duplicate writes)
* Batch updates for Google Sheets
* Graceful degradation on OCR or API failure
* All errors logged for manual review

---

## 📌 Future Enhancements

* LLM-based name disambiguation
* Seat conflict detection
* WhatsApp or email notifications
* Admin dashboard
* PNR response caching

