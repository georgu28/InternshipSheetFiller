#!/usr/bin/env python3
"""
Read job-application links from a text file and append rows to a Google Sheet.

Uses the Google Sheets API (via gspread) with a service-account JSON key.
Share your spreadsheet with the service account email (from the JSON) as Editor.

Each non-empty, non-comment line must be exactly one http(s) URL.
Company is inferred from the hostname; other columns use command-line defaults.

Columns written: Company, Role, Application link, Date applied, OA sent,
Interview stage, Status (Not started, applied, rejected, accepted).

python3 spreadsheet_filler.py my_links.txt --spreadsheet-id YOUR_SHEET_ID --worksheet TabName
python3 spreadsheet_filler.py my_links.txt --spreadsheet-id 1ifevI39buZMlk_7w17BKnnlL52O6UyXYFFj35Id5Fmo --worksheet TestSheet
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

STATUS_CANONICAL = {
    "not started": "Not started",
    "applied": "Applied",
    "rejected": "Rejected",
    "accepted": "Accepted",
}
ALLOWED_STATUSES = frozenset(STATUS_CANONICAL.values())
HEADER = [
    "Company",
    "Role",
    "Application link",
    "Date applied",
    "OA sent",
    "Interview stage",
    "Status",
]

def normalize_status(s: str) -> str:
    key = s.strip().lower()
    if key in STATUS_CANONICAL:
        return STATUS_CANONICAL[key]
    raise argparse.ArgumentTypeError(
        f"Invalid status {s!r}; choose from {sorted(ALLOWED_STATUSES)}"
    )


def _is_http_url(s: str) -> bool:
    u = urlparse(s)
    return u.scheme in ("http", "https") and bool(u.netloc)


def _company_guess_from_url(url: str) -> str:
    host = urlparse(url).netloc.lower()
    if not host:
        return ""
    host = host.split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    parts = host.split(".")
    if len(parts) >= 2:
        return parts[0].replace("-", " ").title()
    return host


def load_urls(path: Path) -> list[str]:
    urls: list[str] = []
    for lineno, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        if not _is_http_url(s):
            raise SystemExit(f"{path}:{lineno}: expected an http(s) URL, got: {s!r}")
        urls.append(s)
    return urls


def ensure_headers(ws: object) -> None:
    existing = ws.get_values("A1:G1")
    if not existing or not any(cell.strip() for cell in existing[0]):
        ws.update("A1:G1", [HEADER])


def existing_urls(ws: object) -> set[str]:
    # Column C is "Application link" by design.
    col = ws.col_values(3)
    return {v.strip() for v in col[1:] if v.strip()} if col else set()

def write_rows(ws: object, values: list[list[str]]) -> None:
    """
    Write rows starting at the first empty row in column C.

    This avoids Google Sheets' append heuristics, which can place values
    outside a formatted table when the sheet has pre-styled empty rows.
    """
    col_c = ws.col_values(3)
    start_row = len(col_c) + 1  # header is row 1; col_values omits trailing blanks
    end_row = start_row + len(values) - 1

    if end_row > ws.row_count:
        ws.add_rows(end_row - ws.row_count)

    ws.update(
        range_name=f"A{start_row}:G{end_row}",
        values=values,
        value_input_option="USER_ENTERED",
    )


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Append job-application rows from a .txt file to a Google Sheet."
    )
    p.add_argument(
        "links_file",
        type=Path,
        help="Text file with one http(s) URL per line",
    )
    p.add_argument(
        "--spreadsheet-id",
        required=True,
        help="Google Sheet ID (from the URL between /d/ and /edit)",
    )
    p.add_argument(
        "--credentials",
        type=Path,
        default=Path("service_account.json"),
        help="Path to Google service account JSON key (default: ./service_account.json)",
    )
    p.add_argument(
        "--worksheet",
        default="Applications",
        help="Worksheet title (created if missing; default: Applications)",
    )
    p.add_argument(
        "--date-applied",
        default="",
        help="Default date applied YYYY-MM-DD (default: today's date)",
    )
    p.add_argument(
        "--oa-sent",
        default="",
        help='Default "OA sent" text for this import (default: empty)',
    )
    p.add_argument(
        "--interview-stage",
        default="",
        help='Default interview stage for this import (default: empty)',
    )
    p.add_argument(
        "--status",
        type=normalize_status,
        default="Applied",
        help='Default status for each imported row (default: "Applied")',
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Parse and print rows without contacting Google",
    )
    p.add_argument(
        "--skip-duplicates",
        action="store_true",
        help="Skip rows whose application URL is already in the sheet",
    )
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    default_date = args.date_applied or date.today().isoformat()

    if not args.links_file.is_file():
        print(f"Error: file not found: {args.links_file}", file=sys.stderr)
        return 1

    if not args.dry_run and not args.credentials.is_file():
        print(
            f"Error: credentials file not found: {args.credentials}\n"
            "Create a service account in Google Cloud, enable Sheets API, "
            "download the JSON key, and share your spreadsheet with the "
            "service account email.",
            file=sys.stderr,
        )
        return 1

    urls = load_urls(args.links_file)
    if not urls:
        print("No rows to import.")
        return 0

    values = [
        [
            _company_guess_from_url(u),
            "",
            u,
            default_date,
            args.oa_sent,
            args.interview_stage,
            args.status,
        ]
        for u in urls
    ]

    if args.dry_run:
        print(json.dumps([HEADER] + values, indent=2))
        return 0

    try:
        import gspread  # type: ignore[import-not-found]
    except ModuleNotFoundError as e:
        raise SystemExit(
            "Missing dependency: gspread\n"
            "Install it using the SAME Python you're running, e.g.:\n"
            "  python -m pip install -r requirements.txt\n"
            "Then re-run your command."
        ) from e

    creds_path = str(args.credentials.resolve())
    gc = gspread.service_account(filename=creds_path)
    sh = gc.open_by_key(args.spreadsheet_id)

    try:
        ws = sh.worksheet(args.worksheet)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(
            title=args.worksheet, rows=1000, cols=len(HEADER)
        )

    ensure_headers(ws)

    if args.skip_duplicates:
        seen = existing_urls(ws)
        filtered: list[list[str]] = []
        skipped = 0
        for row in values:
            if row[2] in seen:
                skipped += 1
                continue
            filtered.append(row)
        values = filtered
        if skipped:
            print(f"Skipped {skipped} duplicate URL(s).", file=sys.stderr)

    if not values:
        print("Nothing new to append.")
        return 0

    write_rows(ws, values)
    print(f"Wrote {len(values)} row(s) to {args.worksheet!r}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
