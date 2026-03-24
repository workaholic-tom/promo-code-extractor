# Promo Code Extractor & Validator

A Python tool that reads marketing emails, extracts promotion codes, and validates them on retailer websites.

## Features

- **Read from macOS Mail.app** — Fetches emails directly via AppleScript (no Full Disk Access needed)
- **Read from files** — Processes downloaded `.eml` or `.html` email files
- **Smart extraction** — Pattern-matching with false-positive filtering to find real promo codes
- **Date filtering** — Process only emails from a specific day
- **Website validation** — Opens a browser to test codes on retailer checkout pages (Saks, Herman Miller, PetSmart, DoorDash)
- **Interactive mode** — Handles bot detection and CAPTCHAs by prompting the user

## Requirements

- Python 3.8+
- macOS (for Mail.app integration)

### Install dependencies

```bash
pip install beautifulsoup4 requests playwright
python -m playwright install chromium
```

## Usage

### Read from macOS Mail.app

```bash
# All accounts, INBOX, specific date (required)
python "project 1.py" --mail --date 2026-03-24

# Specific account
python "project 1.py" --mail --date 2026-03-24 --account "user@gmail.com"

# Different mailbox
python "project 1.py" --mail --date 2026-03-24 --mailbox "Sent Messages"
```

### Read from email files

```bash
# Process a directory of .eml/.html files
python "project 1.py" emails/

# Single file
python "project 1.py" email.eml

# Filter by date
python "project 1.py" emails/ --date 2026-03-24
```

### Options

| Flag | Description |
|------|-------------|
| `--mail` | Read from macOS Mail.app instead of files |
| `--date YYYY-MM-DD` | Only process emails from this date (required with `--mail`) |
| `--account NAME` | Mail.app account to read from (default: all) |
| `--mailbox NAME` | Mail.app mailbox to read from (default: INBOX) |
| `--skip-validation` | Extract codes only, skip website validation |

## How it works

1. **Parse** — Reads email content from Mail.app (via AppleScript) or from `.eml`/`.html` files
2. **Extract** — Scans email text for promo codes using patterns like "use code: SAVE20", "promo code: FAM", etc. Filters out common English words and HTML/social media terms
3. **Validate** — Opens a visible Chromium browser, navigates to the retailer's cart page, and attempts to apply the code. Falls back to interactive mode if bot detection or login is required

## Adding new retailers

To support a new website, add entries to these three sections in `project 1.py`:

1. `SENDER_SITE_MAP` — Map the sender's email domain to a site key
2. `SITE_CONFIGS` — Add the cart URL and a product URL for the site
3. `VALIDATORS` — Add a validation function (or reuse `_validate_with_browser`)
