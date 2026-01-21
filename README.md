# Room Logger

A Flask-based web application for tracking room sign-in and sign-out times. Users can log their entry and exit times using manual entry or RFID cards, and administrators can view, edit, and export logs to CSV or DOCX formats.

## Features

- **Sign In/Out**: 
  - Manual entry: Users can sign in or out with current time or manual time entry
  - RFID card scanning: Automatic sign-in/out using RFID cards (Arduino-compatible)
  - Smart detection: RFID card IDs can be entered in the name field and will be automatically processed
- **RFID Card Management**: 
  - Dedicated page for linking RFID card IDs to user names
  - View all linked RFID cards
  - Add or remove RFID card associations
- **Admin Dashboard**: 
  - View, edit, and delete log entries
  - Logs sorted by newest first (most recent at top)
  - Fixed table size: Always displays at least 10 rows (empty rows if fewer logs), or all rows if more than 10
- **Filtering**: 
  - Filter logs by date, working week (Monday-Friday), or view all dates
  - Auto-selects today's date/week when filtering by date or week
  - Filters automatically apply when selection changes (no need to press a button)
  - Filter state is preserved when editing or deleting entries (stays on current filter unless explicitly changed)
  - Date dropdown persists selected date even after deleting all logs from that date (until filter is changed)
- **Export Options**: 
  - Export to CSV format
  - Export to DOCX format (using FH306 Sign-In Sheet template)
  - Error handling: Shows error toast if attempting to export with no logs for selected date/week
- **Validation**: 
  - Prevents invalid sign-in/out sequences (e.g., can't sign out if not signed in)
  - Validates based on timestamp being submitted, not current state
  - Prevents linking an RFID card that's already linked to another person
- **User Feedback**: 
  - Toast notifications for successful sign-in/out actions (main page)
  - Toast notifications for export success/errors (admin page)
  - Error toasts for validation errors and unregistered RFID cards
  - Auto-dismissing toasts (3 seconds) with centered display
  - Consistent toast styling across all pages (centered, colored borders, icons)
- **Chronological Sorting**: Logs are automatically sorted by timestamp, even if added out of sequence
- **Offline Support**: All dependencies are local - no internet connection required

## Requirements

- Python 3.x
- Flask
- python-docx

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure the following directories exist:
   - `data/` - Stores logs and names
   - `static/` - CSS and JavaScript files
   - `templates/` - HTML templates
   - `docx_templates/` - DOCX template file
   - `exports/` - Export output directory

## Usage

1. Start the application:
```bash
python app.py
```

2. Open your browser and navigate to:
```
http://localhost:8080
```

3. **Sign In/Out Page** (`/`):
   - **Manual Entry**: Select your name from the dropdown or type it, choose to use current time or enter a manual time, then click "Sign In" or "Sign Out"
   - **RFID Card Scanning**: 
     - Scan your RFID card using an Arduino RFID reader (automatically detects and processes)
     - Or type your RFID card ID in the name field - it will be automatically recognized and processed
     - The system automatically determines whether to sign you in or out based on your current state

4. **Admin Page** (`/admin`):
   - View all logs or filter by date/week (defaults to today when filtering)
   - Edit timestamps directly
   - Delete log entries (filter state is preserved after deletion)
   - Export to CSV or DOCX (shows error toast if no logs found for selected filter)
   - Link to RFID Card Management page
   - Toast notifications for export success/errors (matching main page style)

5. **RFID Card Management Page** (`/rfid`):
   - Link RFID card IDs to user names
   - View all linked RFID cards
   - Remove RFID card associations
   - Prevents duplicate card assignments (shows error if card is already linked to another person)
   - Toast notifications for success/error messages

## File Structure

```
Room_Logger/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── data/
│   ├── logs.json         # Log entries (auto-generated)
│   ├── names.json        # User names list
│   └── rfid_cards.json   # RFID card to name mappings (auto-generated)
├── static/
│   ├── bootstrap.min.css
│   ├── bootstrap.bundle.min.js
│   └── style.css
├── templates/
│   ├── index.html        # Sign-in/out page
│   ├── admin.html        # Admin dashboard
│   └── rfid.html         # RFID card management page
├── docx_templates/
│   └── FH306 Sign-In Sheet.docx
└── exports/
    ├── docx_exports/     # Generated DOCX files
    └── room_logs.csv     # Generated CSV files
```

## Export Formats

- **CSV**: Exports log data in CSV format with columns for Name, Date, and up to 4 Time In/Out pairs per row
- **DOCX**: Generates Word documents using the FH306 Sign-In Sheet template, with one document per date

## RFID Card Setup

1. **Hardware Requirements**: 
   - Arduino with RFID reader module
   - RFID cards/tags
   - Arduino configured to type the card ID and press Enter when a card is scanned

2. **Software Setup**:
   - Go to the RFID Card Management page (`/rfid`)
   - Link each user's name to their RFID card ID
   - The RFID card ID is what the Arduino types when scanning

3. **Usage**:
   - **Direct Scanning**: The hidden RFID input field is always ready to receive scans
   - **Manual Entry**: Users can also type their RFID card ID in the name field - it will be automatically recognized
   - **Auto Sign In/Out**: The system automatically determines whether to sign in or out based on current state
   - **Error Handling**: If an unregistered RFID card is scanned, an error toast will appear prompting the user to contact an administrator

## Notes

- Logs are stored in JSON format in `data/logs.json`
- RFID card mappings are stored in `data/rfid_cards.json`
- The application runs on port 8080 by default
- All timestamps are stored in "YYYY-MM-DD HH:MM AM/PM" format
- Missing sign-outs are displayed with empty spots in the output
- Filter selections automatically default to today's date/week for convenience
