# Wedding Ticket Automation System

End-to-end automation to **ingest wedding travel tickets (Train & Flight)** from Google Drive, **extract passenger and journey details using OCR + APIs**, **auto-match guest names against a Master guest list**, and **commit verified travel details into structured Google Sheets**.

This system is designed for **high-volume, semi-automated verification**, with human approval for edge cases (duplicates / low-confidence matches).

---

## High-Level Architecture

```
Google Drive (Tickets)
        │
        ▼
Apps Script Trigger (onDriveChange)
        │  (HTTP POST via ngrok)
        ▼
FastAPI Server
        │
        ├── WeddingTicketAutomation
        │     ├── OCR (PDF/Image)
        │     ├── Ticket Type Detection
        │     ├── Train PNR API / Flight Parsing
        │     └── Append to Ticket Sheet
        │
        └── VerificationWorkflow
              ├── Step 1: Name Matching & Suggestions
              └── Step 2: Approval-Based Commit to Master Sheet
```

---

## Features

### Ticket Ingestion

* Monitors a **Google Drive folder**
* Supports:

  * **Train tickets (IRCTC, Ixigo, screenshots, PDFs)**
  * **Flight tickets (IndiGo, Air India, etc.)**
* Handles:

  * PDFs & images
  * Multiple passengers per ticket
  * OCR noise & low-quality scans
* Automatically moves processed files to a **Processed folder**

---

### Intelligent Data Extraction

* **OCR with multi-pass enhancement**
* **Ticket type detection** (TRAIN / FLIGHT / UNKNOWN)
* **Train PNR validation** using RapidAPI
* Extracts:

  * Journey date
  * Arrival & departure times (separate columns)
  * Train / flight number
  * Seat / berth details
  * Passenger names

Invalid or unsupported files are logged as **ERROR rows** (never silently dropped).

---

### Name Matching & Verification

* Fuzzy matching against **Master guest list**
* Uses:

  * `rapidfuzz.partial_ratio`
  * `token_sort_ratio`
* Normalizes:

  * Titles (Mr, Mrs, Kumar, Devi, etc.)
  * OCR artifacts
* Flags:

  * **Exact matches**
  * **Duplicate / ambiguous matches**
  * **Unmatched names**

Results are written back to the Ticket Sheet for review.

---

### Approval-Based Commit Logic

* Fully automated **only after human approval**
* Requires:

  * Approved name in column **P**
  * `approve_commit = TRUE` in column **R**
* Routing logic:

  * `journey_date >= 2026-02-13` → **DEPARTURE**
  * Earlier dates → **ARRIVAL**
* Writes data into correct **Master Sheet columns**:

  * Arrival → `I:N`
  * Departure → `AE:AJ`

---

## Google Sheets Structure

### Ticket Sheet (Example: `Sheet13`)

| Column | Purpose                       |
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
| M      | Source File                   |
| N      | Suggested Name                |
| O      | Match Score                   |
| P      | Approved Name                 |
| Q      | Commit Status                 |
| R      | Approve Commit (TRUE/FALSE)   |

---

### Master Sheet

* Column A contains **canonical guest names**
* Arrival & departure details are written into fixed column blocks
* Supports multiple travel entries per guest

---

## FastAPI Endpoints

### `POST /ingest-and-match`

* Full automation:

  * Ingest tickets from Drive
  * Append to sheet
  * Auto-match names

### `POST /step2-commit`

* Executes **approval-based commit**
* Writes finalized data into Master Sheet

---

## Google Apps Script Trigger

```javascript
function onDriveChange(e) {
  UrlFetchApp.fetch(
    "https://<your-ngrok-url>.ngrok-free.dev/ingest-and-match",
    {
      method: "POST",
      muteHttpExceptions: true
    }
  );
}
```

* Triggered on Drive changes
* Acts as the **bridge between Google Drive and FastAPI**
* Uses **ngrok** to expose local FastAPI securely

---

## Environment Variables

```env
# Google
SERVICE_ACCOUNT_FILE=service_account.json
SHEET_ID=xxxxxxxx
SHEET_NAME=Sheet13
DRIVE_FOLDER_ID=xxxxxxxx
PROCESSED_FOLDER_ID=xxxxxxxx

# OCR
TESSERACT_PATH=/usr/bin/tesseract

# Train API
RAPIDAPI_KEY=xxxxxxxx
RAPIDAPI_HOST=irctc-indian-railway-pnr-status.p.rapidapi.com
```

---

## Tech Stack

* **FastAPI** – API orchestration
* **Google Drive API** – File ingestion
* **Google Sheets API** – Data store & workflow control
* **Tesseract OCR + PIL + PyMuPDF**
* **RapidFuzz** – Name matching
* **RapidAPI (IRCTC PNR)**
* **Google Apps Script**
* **ngrok** – Secure public tunnel

---

## Operational Workflow

1. Upload tickets to Google Drive
2. Apps Script triggers FastAPI
3. Tickets are OCR-processed & appended
4. Names are auto-matched
5. Human reviews suggestions
6. Admin sets `approve_commit = TRUE`
7. Step 2 commits data to Master Sheet

---

## Design Philosophy

* **No silent failures**
* **Human-in-the-loop for critical decisions**
* **Idempotent & resumable**
* **Production-safe batch updates**
* **Extensible for buses / hotels later**

---

## Status

**Production-ready**
Actively used for real-world wedding guest travel coordination.
