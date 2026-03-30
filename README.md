# Internship Sheet Filler (Google Sheets)

Append job-application URLs from a text file into a Google Sheet using the Google Sheets API via a **Google Cloud service account**.

## Prerequisites

- Python 3.10+ recommended
- A Google Sheet you own (or can edit)
- A Google Cloud project with the **Google Sheets API** enabled

## Setup

### 1) Create and activate a virtual environment

From the `adv-python/` folder:

```bash
python -m venv env
```

Activate it:

- **Git Bash**

```bash
source env/Scripts/activate
```

- **PowerShell**

```powershell
.\env\Scripts\Activate.ps1
```

### 2) Install requirements

```bash
python -m pip install -r requirements.txt
```

## Google service account credentials (`service_account.json`)

This project authenticates using a service account key file.

### 1) Create a service account + key

In Google Cloud Console:

- Create a **Service Account**
- Enable the **Google Sheets API** for your project
- Create a **JSON key** for the service account and download it

### 2) Put the JSON key in this folder

Save the downloaded file as:

- `adv-python/service_account.json`

This file is **ignored by git** (see `.gitignore`). Do not commit it.

### 3) Share your spreadsheet with the service account email

Open the JSON and copy `client_email`, then share your Google Sheet with that email as **Editor**.

## Input file format

Your links file must be a UTF-8 text file containing:

- One **http(s)** URL per line
- Blank lines are ignored
- Lines starting with `#` are treated as comments and ignored

Example (`my_links.txt`):

```txt
# applications
https://example.com/jobs/123
https://careers.somewhere.com/apply/abc
```

## Run

You need:

- The **Spreadsheet ID** (from the Sheet URL: between `/d/` and `/edit`)
- A worksheet/tab name (defaults to `Applications`; created if missing)

Example:

```bash
python spreadsheet_filler.py my_links.txt --spreadsheet-id YOUR_SHEET_ID --worksheet Applications
```

Useful options:

- `--credentials PATH`: path to the service account JSON (default `./service_account.json`)
- `--dry-run`: prints the rows it would write, without contacting Google
- `--skip-duplicates`: skips URLs already present in column C ("Application link")
- `--status Applied|Rejected|Accepted|Not started`: default status for imported rows
- `--date-applied YYYY-MM-DD`: default date (defaults to today)

Dry run example:

```bash
python spreadsheet_filler.py my_links.txt --spreadsheet-id YOUR_SHEET_ID --dry-run
```

## Security notes

- Treat `service_account.json` like a password. If you ever committed it, **rotate/revoke the key** in Google Cloud immediately.
- Don’t paste the key into chat logs or commit it to any repository.

