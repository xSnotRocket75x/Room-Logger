# Room Logger

A Flask-based web application for tracking room sign-in and sign-out times. Users can log their entry and exit times, and administrators can view, edit, and export logs to CSV or DOCX formats.

## Features

- **Sign In/Out**: Users can sign in or out with current time or manual time entry
- **Admin Dashboard**: View, edit, and delete log entries
- **Filtering**: Filter logs by date, working week (Monday-Friday), or view all dates
- **Export Options**: 
  - Export to CSV format
  - Export to DOCX format (using FH306 Sign-In Sheet template)
- **Validation**: Prevents invalid sign-in/out sequences (e.g., can't sign out if not signed in)
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
   - Select your name from the dropdown or type it
   - Choose to use current time or enter a manual time
   - Click "Sign In" or "Sign Out"

4. **Admin Page** (`/admin`):
   - View all logs or filter by date/week
   - Edit timestamps directly
   - Delete log entries
   - Export to CSV or DOCX

## File Structure

```
Room_Logger/
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
├── data/
│   ├── logs.json         # Log entries (auto-generated)
│   └── names.json        # User names list
├── static/
│   ├── bootstrap.min.css
│   ├── bootstrap.bundle.min.js
│   └── style.css
├── templates/
│   ├── index.html        # Sign-in/out page
│   └── admin.html       # Admin dashboard
├── docx_templates/
│   └── FH306 Sign-In Sheet.docx
└── exports/
    ├── docx_exports/     # Generated DOCX files
    └── room_logs.csv     # Generated CSV files
```

## Export Formats

- **CSV**: Exports log data in CSV format with columns for Name, Date, and up to 4 Time In/Out pairs per row
- **DOCX**: Generates Word documents using the FH306 Sign-In Sheet template, with one document per date

## Notes

- Logs are stored in JSON format in `data/logs.json`
- The application runs on port 8080 by default
- All timestamps are stored in "YYYY-MM-DD HH:MM AM/PM" format
- Missing sign-outs are displayed with empty spots in the output
