"""
Promotion Code Extractor & Validator

Reads emails from macOS Mail.app or downloaded files (.eml/.html),
extracts promotion codes, and tests them against company websites.

Usage:
    python "project 1.py" --mail --date YYYY-MM-DD     # read from Mail.app
    python "project 1.py" emails/ --date 2026-03-24     # read from files
    python "project 1.py" emails/                       # all files, no date filter
"""

import argparse
import json
import subprocess
import sys
import os
import re
import email
from datetime import datetime, date, timedelta
from email import policy
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


# --- Email Parsing ---

def parse_eml_file(filepath):
    """Parse a .eml file and return the plain text and HTML body."""
    with open(filepath, "rb") as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    text_body = ""
    html_body = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                text_body += part.get_content()
            elif content_type == "text/html":
                html_body += part.get_content()
    else:
        content_type = msg.get_content_type()
        if content_type == "text/plain":
            text_body = msg.get_content()
        elif content_type == "text/html":
            html_body = msg.get_content()

    subject = msg.get("Subject", "")
    sender = msg.get("From", "")

    email_date = None
    date_str = msg.get("Date")
    if date_str:
        try:
            email_date = parsedate_to_datetime(date_str).date()
        except Exception:
            pass

    return subject, sender, text_body, html_body, email_date


def parse_html_file(filepath):
    """Parse a standalone .html email file. Uses file modification date as fallback."""
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        html_body = f.read()
    file_date = date.fromtimestamp(os.path.getmtime(filepath))
    return "", "", "", html_body, file_date


def extract_text_from_html(html):
    """Convert HTML to readable text."""
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text(separator=" ", strip=True)


# --- macOS Mail.app Integration (AppleScript) ---

def _run_applescript(script):
    """Run an AppleScript and return its stdout."""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
    return result.stdout.strip()


def get_mail_accounts():
    """Return list of account names from Mail.app."""
    output = _run_applescript('tell application "Mail" to get name of every account')
    if not output:
        return []
    return [a.strip() for a in output.split(",")]


def fetch_messages_from_mail(target_date=None, account=None, mailbox_name="INBOX"):
    """
    Fetch emails from macOS Mail.app via AppleScript.

    Args:
        target_date: A date object to filter messages (required).
        account: Account name string, or None for all accounts.
        mailbox_name: Mailbox to read from (default: INBOX).

    Returns:
        List of dicts with keys: subject, sender, source, date
    """
    if not target_date:
        raise ValueError("--date is required when reading from Mail.app")

    # Format dates for AppleScript
    # AppleScript needs dates in the system's locale format
    start_date = target_date.strftime("%A, %B %d, %Y") + " at 12:00:00 AM"
    end_date = (target_date + timedelta(days=1)).strftime("%A, %B %d, %Y") + " at 12:00:00 AM"

    accounts = [account] if account else get_mail_accounts()
    all_messages = []

    for acct_name in accounts:
        print(f"Reading {mailbox_name} from account: {acct_name}...")

        # Step 1: Get count and metadata for messages on the target date
        count_script = f'''
tell application "Mail"
    set targetDate to date "{start_date}"
    set endDate to date "{end_date}"
    set msgs to (every message of mailbox "{mailbox_name}" of account "{acct_name}" whose date received ≥ targetDate and date received < endDate)
    return count of msgs
end tell'''
        try:
            count = int(_run_applescript(count_script))
        except (ValueError, RuntimeError) as e:
            print(f"  Could not read from {acct_name}: {e}")
            continue

        if count == 0:
            print(f"  No messages found on {target_date}.")
            continue

        print(f"  Found {count} message(s) on {target_date}. Fetching...")

        # Step 2: Fetch each message's source individually
        # (AppleScript can't return a list of large strings reliably)
        for i in range(1, count + 1):
            msg_script = f'''
tell application "Mail"
    set targetDate to date "{start_date}"
    set endDate to date "{end_date}"
    set msgs to (every message of mailbox "{mailbox_name}" of account "{acct_name}" whose date received ≥ targetDate and date received < endDate)
    set msg to item {i} of msgs
    set subj to subject of msg
    set sndr to sender of msg
    set src to source of msg
    -- Use a delimiter that won't appear in email content
    return subj & "|||DELIM|||" & sndr & "|||DELIM|||" & src
end tell'''
            try:
                raw = _run_applescript(msg_script)
                parts = raw.split("|||DELIM|||", 2)
                if len(parts) == 3:
                    subject, sender, source = parts
                    all_messages.append({
                        "subject": subject.strip(),
                        "sender": sender.strip(),
                        "source": source.strip(),
                        "account": acct_name,
                    })
                    print(f"  [{i}/{count}] {subject.strip()[:70]}")
                else:
                    print(f"  [{i}/{count}] Failed to parse message (unexpected format)")
            except RuntimeError as e:
                print(f"  [{i}/{count}] Error fetching message: {e}")
            except subprocess.TimeoutExpired:
                print(f"  [{i}/{count}] Timed out fetching message, skipping...")

    return all_messages


def parse_mail_source(source):
    """Parse raw email source (from Mail.app) and return subject, sender, text_body, html_body."""
    msg = email.message_from_string(source, policy=policy.default)

    text_body = ""
    html_body = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            try:
                if content_type == "text/plain":
                    text_body += part.get_content()
                elif content_type == "text/html":
                    html_body += part.get_content()
            except Exception:
                continue
    else:
        content_type = msg.get_content_type()
        try:
            if content_type == "text/plain":
                text_body = msg.get_content()
            elif content_type == "text/html":
                html_body = msg.get_content()
        except Exception:
            pass

    subject = msg.get("Subject", "")
    sender = msg.get("From", "")
    return subject, sender, text_body, html_body


# --- Promo Code Extraction ---

# Patterns that require an explicit label before the code (high confidence)
PROMO_PATTERNS = [
    # "promo code: SAVE20", "coupon code: WINTER50", "use code FREESHIP", "enter code ABC123"
    r'(?:promo(?:tion)?[\s\-]*code|coupon[\s\-]*code|discount[\s\-]*code|voucher[\s\-]*code|offer[\s\-]*code|use[\s\-]+code|enter[\s\-]+code|apply[\s\-]+code|your[\s\-]+code(?:\s+is)?)\s*[:\-\s]\s*([A-Z0-9][A-Z0-9\-_]{2,20})',
    # "code: SAVE20" (only when preceded by context words like "with", "the", "your", "this", "our")
    r'(?:with|the|your|this|our|a)\s+code\s*[:\-\s]\s*([A-Z0-9][A-Z0-9\-_]{2,20})',
    # "code SAVE20 at checkout"
    r'code\s+([A-Z0-9][A-Z0-9\-_]{2,20})\s+(?:at|during|for|to)\s+(?:checkout|check\s*out|cart|purchase|order)',
]


def is_likely_promo_code(code):
    """Check if a matched string looks like a real promo code vs a regular word."""
    # Too short or too long
    if len(code) < 3 or len(code) > 20:
        return False
    # Must contain at least one digit, or be a recognizable promo-style word (mix of letters)
    # Pure dictionary-like words are unlikely to be codes
    has_digit = any(c.isdigit() for c in code)
    has_special = any(c in "-_" for c in code)
    # Codes with digits or special chars are almost certainly codes
    if has_digit or has_special:
        return True
    # Pure-letter codes: only accept if they look promo-like (short, brandish)
    # Reject common English words by checking against a word-length heuristic
    # Very common words that appear in marketing emails
    common_words = {
        "THE", "AND", "FOR", "ARE", "BUT", "NOT", "YOU", "ALL", "CAN", "HAD",
        "HER", "WAS", "ONE", "OUR", "OUT", "DAY", "GET", "HAS", "HIM", "HIS",
        "HOW", "ITS", "MAY", "NEW", "NOW", "OLD", "SEE", "WAY", "WHO", "BOY",
        "DID", "LET", "SAY", "SHE", "TOO", "USE", "DAD", "MOM", "OFF", "BIG",
        "JUST", "LIKE", "THAN", "THEM", "BEEN", "CALL", "COME", "EACH", "FIND",
        "FROM", "GIVE", "HAVE", "HELP", "HERE", "HIGH", "HOME", "INTO", "KEEP",
        "LAST", "LONG", "LOOK", "MADE", "MAKE", "MANY", "MORE", "MOST", "MUCH",
        "MUST", "NAME", "NEXT", "ONLY", "OPEN", "OVER", "PART", "PICK", "PLAY",
        "SAME", "SHOW", "SIDE", "SOME", "SUCH", "SURE", "TAKE", "TELL", "THAT",
        "THIS", "TIME", "TURN", "VERY", "WANT", "WELL", "WENT", "WHAT", "WHEN",
        "WILL", "WITH", "WORD", "WORK", "YEAR", "YOUR", "BACK", "BEST", "BOTH",
        "CITY", "CLUB", "CODE", "COPY", "DATA", "DONE", "DOWN", "EASY", "EVER",
        "FACT", "FAST", "FEEL", "FILL", "FORM", "FULL", "GOOD", "HAND", "HEAD",
        "ITEM", "JUST", "KNOW", "LEFT", "LIFE", "LINE", "LIST", "LIVE", "LOVE",
        "MAIL", "MARK", "MOVE", "NEED", "NEWS", "NICE", "NOTE", "PAGE", "PAIR",
        "PLAN", "POST", "PULL", "PUSH", "READ", "REAL", "RENT", "REST", "RISE",
        "ROAD", "RULE", "RUNS", "SAVE", "SEEN", "SELF", "SEND", "SHIP", "SHOP",
        "SIZE", "SOLD", "STAR", "STAY", "STEP", "STOP", "TEST", "THEM", "THEN",
        "THEY", "THUS", "TILL", "TOOL", "TOPS", "TREE", "TRIP", "TRUE", "TYPE",
        "UNIT", "UPON", "USED", "USER", "VIEW", "WAIT", "WALK", "WALL", "WAYS",
        "WEEK", "WIDE", "WISH", "ZERO", "ZONE",
        "ABOUT", "ABOVE", "ADDED", "AFTER", "AGAIN", "ALONG", "APPLY", "BASED",
        "BASIC", "BEGIN", "BEING", "BELOW", "BLACK", "BLOCK", "BOARD", "BONUS",
        "BRAND", "BREAK", "BRING", "BROWN", "BUILD", "BUILT", "BUYER", "CARRY",
        "CATCH", "CAUSE", "CHECK", "CHILD", "CLAIM", "CLASS", "CLEAN", "CLEAR",
        "CLICK", "CLOSE", "COLOR", "COMES", "COULD", "COUNT", "COVER", "DAILY",
        "DEALS", "DOING", "DRAFT", "DRINK", "DRIVE", "EARLY", "EARTH", "EMAIL",
        "EMPTY", "ENDED", "ENJOY", "ENTER", "ERROR", "EVENT", "EVERY", "EXACT",
        "EXTRA", "FANCY", "FEAST", "FIELD", "FINAL", "FIRST", "FIXED", "FLASH",
        "FOCUS", "FORCE", "FOUND", "FRESH", "FRONT", "FULLY", "GIVEN", "GOING",
        "GOODS", "GRAND", "GREAT", "GREEN", "GROUP", "GUIDE", "HAPPY", "HEART",
        "HEAVY", "HOURS", "HOUSE", "HURRY", "IMAGE", "INBOX", "INNER", "INPUT",
        "ITEMS", "JOINT", "KNOWN", "LARGE", "LATER", "LEARN", "LEAVE", "LEVEL",
        "LIGHT", "LIMIT", "LOCAL", "LOOKS", "LOWER", "LUCKY", "LUNCH", "MAGIC",
        "MAJOR", "MATCH", "MEANT", "MEDIA", "MIGHT", "MODEL", "MONEY", "MONTH",
        "NIGHT", "NORTH", "NOTED", "OFFER", "ORDER", "OTHER", "OUTER", "OWNED",
        "PARTY", "PHONE", "PHOTO", "PIECE", "PLACE", "PLANT", "PLATE", "POINT",
        "POWER", "PRESS", "PRICE", "PRIME", "PRINT", "QUICK", "QUITE", "RAISE",
        "RANGE", "RATED", "REACH", "READY", "RIGHT", "ROUND", "ROYAL", "SCORE",
        "SENSE", "SERVE", "SHARE", "SHEET", "SHORT", "SHOWN", "SIGHT", "SINCE",
        "SIZED", "SKILL", "SMALL", "SMART", "SOLID", "SOUND", "SOUTH", "SPACE",
        "SPEAK", "SPEED", "SPEND", "SPLIT", "SPORT", "STAFF", "STAGE", "STAND",
        "START", "STATE", "STEAM", "STILL", "STOCK", "STONE", "STORE", "STORY",
        "STUFF", "STYLE", "SUPER", "TABLE", "TAKEN", "TERMS", "THEIR", "THEME",
        "THERE", "THESE", "THING", "THINK", "THIRD", "THREE", "TODAY", "TOTAL",
        "TOUCH", "TOUGH", "TRACK", "TRADE", "TRAIN", "TREAT", "TREND", "TRIAL",
        "TRULY", "TRUST", "TWICE", "UNDER", "UNTIL", "UPPER", "USING", "USUAL",
        "VALID", "VALUE", "VIDEO", "VISIT", "VOICE", "WATCH", "WATER", "WEEKS",
        "WHERE", "WHICH", "WHILE", "WHITE", "WHOLE", "WOMAN", "WOMEN", "WORLD",
        "WORRY", "WORSE", "WORST", "WORTH", "WOULD", "WRITE", "YOUNG",
        # Social media / tech / HTML terms common in email footers
        "FACEBOOK", "TWITTER", "INSTAGRAM", "LINKEDIN", "PINTEREST", "YOUTUBE",
        "TIKTOK", "THREADS", "WHATSAPP", "SNAPCHAT", "REDDIT", "TUMBLR",
        "FBICON", "ANDROID", "IPHONE", "MOBILE", "TABLET", "DESKTOP",
        "ARIAL", "HELVETICA", "VERDANA", "SERIF", "NORMAL", "BOLD", "ITALIC",
        "SOLID", "BORDER", "PADDING", "MARGIN", "DISPLAY", "INLINE", "HIDDEN",
        "CENTER", "MIDDLE", "BOTTOM", "REPEAT", "SCROLL", "FIXED", "STATIC",
        "RELATIVE", "ABSOLUTE", "CONTENT", "WRAPPER", "HEADER", "FOOTER",
        "CONTAINER", "SECTION", "COLUMN", "MODULE", "LAYOUT", "BUTTON",
        "SUBMIT", "CANCEL", "DELETE", "UPDATE", "INSERT", "SELECT", "OPTION",
        "TOGGLE", "SWITCH", "CHANGE", "REMOVE", "MANAGE", "REVIEW", "SEARCH",
        "FILTER", "BROWSE", "DETAIL", "SIGNUP", "SIGNIN", "LOGOUT",
        "PRIVACY", "POLICY", "COOKIE", "NOTICE", "RIGHTS", "RESERVED",
        "UNSUBSCRIBE", "PREFERENCES", "SETTINGS", "ACCOUNT", "PROFILE",
        "SUBSCRIBE", "NEWSLETTER", "CONTACT", "SUPPORT", "SERVICE",
        # Marketing but not codes
        "SALE", "FREE", "DEAL", "GIFT", "PERCENT", "DISCOUNT", "COUPON",
        "REWARD", "POINTS", "MEMBER", "PROGRAM", "WELCOME", "SPECIAL",
        "LIMITED", "EXCLUSIVE", "PREMIUM", "LUXURY", "SPRING", "SUMMER",
        "WINTER", "AUTUMN", "HOLIDAY", "SEASON", "COLLECTION", "EDITION",
        "CLASSIC", "MODERN", "VINTAGE", "CUSTOM", "DESIGN", "DESIGNED",
        "HANDMADE", "ORIGINAL", "NATURAL", "ORGANIC", "GLOBAL", "SELECT",
        "SELECTED", "POPULAR", "TRENDING", "FEATURED", "RECOMMENDED",
        "ELIGIBLE", "INCLUDED", "REQUIRED", "AVAILABLE", "SHIPPING",
        "DELIVERY", "RETURNS", "REFUND", "EXCHANGE", "WARRANTY",
        "PURCHASE", "PAYMENT", "CHECKOUT", "RECEIPT", "INVOICE",
        "CONFIRM", "CONFIRMED", "COMPLETE", "COMPLETED", "PROCESSING",
        "PENDING", "APPROVED", "DENIED", "EXPIRED", "ACTIVE", "INACTIVE",
        # Common nouns / misc
        "PRODUCTIONS", "INSTALLATION", "PERFORMANCE", "INFRASTRUCTURE",
        "DESCRIPTIONS", "TOKENIZER", "REASONING", "SOFTWARE", "HARDWARE",
        "PLATFORM", "ENCODER", "ANGULAR", "COORDINATE", "CONTINUE",
        "IMPORTANT", "REMINDER", "REMINDERS", "NOTIFICATION", "ANNOUNCE",
        "FISHBOWL", "OUTFITTED", "GENERALITY", "FRESHEST", "SINGLE",
        "SYSTEM", "SYSTEMS", "MODELS", "BANNER", "VISTAS", "PIXELS",
        "NVIDIA", "HELPFUL", "ADULT", "PAYING", "CLIENT", "RISK",
        "AHEAD", "FOOT", "BALL", "NEED", "TRAIN", "THERE", "ONLY",
        "WANT", "AUTO", "CARD", "MOVE", "SENT", "TEAM", "SECURE",
    }
    return code not in common_words


def extract_promo_codes(text):
    """Extract potential promotion codes from email text."""
    codes = set()
    for pattern in PROMO_PATTERNS:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            code = match.strip().upper()
            if is_likely_promo_code(code):
                codes.add(code)
    return list(codes)


# --- Sender-to-site mapping ---

# Maps sender email domain to a site key used for validation routing
SENDER_SITE_MAP = {
    "e.saks.com": "saks",
    "saks.com": "saks",
    "n.hermanmiller.com": "hermanmiller",
    "hermanmiller.com": "hermanmiller",
    "messages.doordash.com": "doordash",
    "doordash.com": "doordash",
    "mail.petsmart.com": "petsmart",
    "petsmart.com": "petsmart",
}


def get_site_from_sender(sender):
    """Extract a site key from the sender email address."""
    match = re.search(r'@([\w.\-]+)', sender)
    if not match:
        return None
    domain = match.group(1).lower()
    # Try exact match, then parent domain
    if domain in SENDER_SITE_MAP:
        return SENDER_SITE_MAP[domain]
    # Try stripping first subdomain (e.g., mail.petsmart.com -> petsmart.com)
    parts = domain.split(".")
    if len(parts) > 2:
        parent = ".".join(parts[1:])
        if parent in SENDER_SITE_MAP:
            return SENDER_SITE_MAP[parent]
    return domain  # fallback: return the domain itself


# --- Promo Code Validation (Playwright browser automation) ---

# Site configurations: URL to visit, and how to find/expand the promo code field
SITE_CONFIGS = {
    "saks": {
        "cart_url": "https://www.saks.com/checkout/bag",
        "product_url": "https://www.saks.com/product/saks-fifth-avenue-collection-cashmere-crewneck-sweater-0400021876498.html",
        "add_to_bag": 'button:has-text("Add to Bag"), button:has-text("ADD TO BAG"), button[data-testid="add-to-bag"]',
    },
    "hermanmiller": {
        "cart_url": "https://store.hermanmiller.com/cart",
        "product_url": "https://store.hermanmiller.com/home-office-accessories/restoring-cream/2522116.html",
        "add_to_bag": 'button:has-text("Add to Cart"), button:has-text("ADD TO CART"), button[data-testid="add-to-cart"]',
    },
    "petsmart": {
        "cart_url": "https://www.petsmart.com/cart/",
        "product_url": "https://www.petsmart.com/dog/treats/dental-treats/greenies-original-regular-dog-dental-treat-5076050.html",
        "add_to_bag": 'button:has-text("Add to Cart"), button:has-text("ADD TO CART")',
    },
}

# Generic selectors to find promo code input and apply button on cart/checkout pages
PROMO_INPUT_SELECTORS = [
    'input[name="promoCode"]',
    'input[name="couponCode"]',
    'input[name="coupon_code"]',
    'input[name="promo"]',
    'input[data-testid*="promo" i]',
    'input[data-testid*="coupon" i]',
    'input[placeholder*="promo" i]',
    'input[placeholder*="coupon" i]',
    'input[placeholder*="discount" i]',
    'input[placeholder*="enter code" i]',
    'input[aria-label*="promo" i]',
    'input[aria-label*="coupon" i]',
    '#promoCode',
    '#couponCode',
    '#promo-code',
    '#coupon-code',
]

# Clickable elements that might expand/reveal a hidden promo code section
PROMO_TOGGLE_SELECTORS = [
    'button:has-text("promo code")',
    'button:has-text("Promo Code")',
    'button:has-text("coupon code")',
    'a:has-text("promo code")',
    'a:has-text("Promo Code")',
    'a:has-text("Have a code")',
    'a:has-text("coupon")',
    'span:has-text("promo code")',
    'div:has-text("promo code") >> visible=true',
    '[data-testid*="promo-toggle"]',
    '[data-testid*="coupon-toggle"]',
    'details:has-text("promo")',
    'summary:has-text("promo")',
]

APPLY_BUTTON_SELECTORS = [
    'button:has-text("Apply")',
    'button:has-text("APPLY")',
    'button[aria-label*="apply" i]',
    'button[data-testid*="apply" i]',
    'button[type="submit"]:near(input[name="promoCode"])',
    'input[type="submit"][value*="Apply" i]',
]


def _launch_browser(pw, headless=False):
    """Launch a Chromium browser. Uses visible mode by default so user can handle CAPTCHAs."""
    browser = pw.chromium.launch(headless=headless)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="en-US",
    )
    return browser, context


def _try_click(page, selectors, timeout=3000):
    """Try clicking the first matching visible element from a list of selectors."""
    for selector in selectors:
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=timeout):
                el.click()
                return True
        except Exception:
            continue
    return False


def _try_add_item(page, site_key):
    """Navigate to a product page and try to add an item to cart."""
    config = SITE_CONFIGS.get(site_key, {})
    product_url = config.get("product_url")
    if not product_url:
        return False

    page.goto(product_url, timeout=30000)
    page.wait_for_timeout(3000)

    add_selectors = [s.strip() for s in config.get("add_to_bag", "").split(",")]
    return _try_click(page, add_selectors, timeout=5000)


def _find_promo_input(page):
    """Try to find and return a promo code input field, expanding toggles if needed."""
    # First, try to find the input directly
    for selector in PROMO_INPUT_SELECTORS:
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=2000):
                return el
        except Exception:
            continue

    # If not found, try clicking toggles to reveal it
    _try_click(page, PROMO_TOGGLE_SELECTORS, timeout=2000)
    page.wait_for_timeout(1000)

    # Try again after expanding
    for selector in PROMO_INPUT_SELECTORS:
        try:
            el = page.locator(selector).first
            if el.is_visible(timeout=2000):
                return el
        except Exception:
            continue

    return None


def _apply_code_and_read_result(page, code):
    """Fill in a promo code, submit it, and return (valid, message)."""
    promo_input = _find_promo_input(page)
    if not promo_input:
        return None, "Could not find promo code input field on the page."

    promo_input.click()
    promo_input.fill(code)
    page.wait_for_timeout(500)

    # Try clicking Apply button
    if not _try_click(page, APPLY_BUTTON_SELECTORS, timeout=3000):
        promo_input.press("Enter")

    page.wait_for_timeout(4000)

    # Read the page and analyze
    text = page.content().lower()

    # Positive signals
    if any(s in text for s in ["successfully applied", "code applied", "coupon applied",
                                "discount applied", "promo applied", "has been applied",
                                "promotion applied", "you saved"]):
        return True, "Promo code was successfully applied!"

    # Negative signals — expired
    if any(s in text for s in ["expired", "no longer valid", "no longer active",
                                "code has expired", "promotion has ended",
                                "offer has ended", "offer expired"]):
        return False, "Promo code has expired."

    # Negative signals — invalid
    if any(s in text for s in ["invalid code", "invalid coupon", "invalid promo",
                                "not a valid", "is not valid", "code not found",
                                "doesn't exist", "does not exist", "not recognized",
                                "unable to apply", "cannot be applied", "can't be applied",
                                "didn't work", "did not work", "incorrect code"]):
        return False, "Promo code is invalid or not recognized."

    # Conditional
    if any(s in text for s in ["minimum purchase", "minimum order", "does not meet",
                                "not eligible", "not applicable", "requirements not met"]):
        return None, "Code may be valid but requires minimum purchase or specific items in cart."

    return None, "Code was submitted but could not determine validity from page response."


def _validate_with_browser(code, site_key):
    """
    Open a visible browser for the user to interact with, then attempt to
    apply the promo code. Works in two modes:

    1. Auto mode: Tries to navigate to cart/checkout and apply the code automatically.
    2. Interactive fallback: If auto mode can't find the promo field (e.g., empty cart,
       bot detection), it prompts the user to manually add an item and get to the
       promo code screen, then the script takes over to enter and submit the code.
    """
    config = SITE_CONFIGS.get(site_key, {})
    cart_url = config.get("cart_url")
    if not cart_url:
        return None, f"No cart URL configured for site '{site_key}'."

    try:
        with sync_playwright() as pw:
            browser, context = _launch_browser(pw, headless=False)
            page = context.new_page()

            # Step 1: Try automatic approach first
            print(f"      Opening {site_key} cart page...")
            page.goto(cart_url, timeout=30000)
            page.wait_for_timeout(4000)

            # Check if we hit bot detection
            page_text = page.content().lower()
            if any(s in page_text for s in ["unusual activity", "captcha", "verify you are human",
                                             "access denied", "blocked"]):
                print(f"      Bot detection triggered. Switching to interactive mode...")
                input(f"\n      >>> Please resolve the bot check in the browser window, "
                      f"then press ENTER here to continue...")
                page.wait_for_timeout(2000)

            # Check if cart is empty
            page_text = page.content().lower()
            if any(s in page_text for s in ["cart is empty", "bag is empty", "no items",
                                             "nothing in your cart", "nothing in your bag"]):
                print(f"      Cart is empty. Please add an item to test the promo code.")
                input(f"\n      >>> In the browser: add any item to your cart, then go to "
                      f"cart/checkout.\n"
                      f"      >>> Once you can see a promo code field, press ENTER here...")
                page.wait_for_timeout(2000)

            # Step 2: Try to find and fill promo code
            valid, message = _apply_code_and_read_result(page, code)

            # Step 3: If still can't find the field, ask user one more time
            if valid is None and "could not find" in message.lower():
                print(f"      Still can't find the promo code field automatically.")
                user_input = input(
                    f"\n      >>> If you can see a promo code field in the browser, type 'retry'.\n"
                    f"      >>> Or type 'manual' to enter the code yourself and report the result.\n"
                    f"      >>> Or press ENTER to skip: "
                ).strip().lower()

                if user_input == "retry":
                    page.wait_for_timeout(1000)
                    valid, message = _apply_code_and_read_result(page, code)
                elif user_input == "manual":
                    result = input(
                        f"      >>> Enter the code '{code}' on the website manually.\n"
                        f"      >>> Then type the result (valid/invalid/expired/unknown): "
                    ).strip().lower()
                    if result == "valid":
                        valid, message = True, "Manually confirmed as valid by user."
                    elif result in ("invalid", "expired"):
                        valid, message = False, f"Manually confirmed as {result} by user."
                    else:
                        valid, message = None, "User could not determine validity."

            # Take a screenshot for reference
            screenshot_path = f"validation_{site_key}_{code}.png"
            page.screenshot(path=screenshot_path)
            message += f" (screenshot: {screenshot_path})"

            browser.close()
            return valid, message
    except Exception as e:
        return None, f"Browser automation failed: {e}"


def validate_saks(code):
    """Validate a promo code on saks.com."""
    return _validate_with_browser(code, "saks")


def validate_hermanmiller(code):
    """Validate a promo code on hermanmiller.com."""
    return _validate_with_browser(code, "hermanmiller")


def validate_petsmart(code):
    """Validate a promo code on petsmart.com."""
    return _validate_with_browser(code, "petsmart")


def validate_doordash(code):
    """DoorDash promo codes are account-specific and require authentication."""
    return None, "DoorDash codes are account-specific and require login. Cannot validate without authentication."


# Dispatch table: site key -> validation function
VALIDATORS = {
    "saks": validate_saks,
    "hermanmiller": validate_hermanmiller,
    "petsmart": validate_petsmart,
    "doordash": validate_doordash,
}


def validate_promo_code(code, site=None):
    """
    Validate a promo code against the appropriate website.

    Args:
        code: The promo code string.
        site: Site key (e.g., "saks", "petsmart"). If None, returns unknown.

    Returns:
        (valid, message) where valid is True/False/None and message is a detail string.
    """
    if site and site in VALIDATORS:
        print(f"    Validating '{code}' on {site}...")
        return VALIDATORS[site](code)
    elif site:
        return None, f"No validator configured for site '{site}'. Add one to VALIDATORS dict."
    else:
        return None, "Could not determine source website for this code."


# --- Main ---

def process_email_file(filepath, filter_date=None):
    """Process a single email file. Returns (codes, sender, skipped)."""
    filepath = Path(filepath)
    ext = filepath.suffix.lower()

    if ext == ".eml":
        subject, sender, text_body, html_body, email_date = parse_eml_file(filepath)
    elif ext in (".html", ".htm"):
        subject, sender, text_body, html_body, email_date = parse_html_file(filepath)
    else:
        print(f"  Skipping unsupported file: {filepath}")
        return [], "", True

    # Filter by date if specified
    if filter_date and email_date != filter_date:
        return [], sender, True

    # Combine text from all sources for code extraction
    full_text = " ".join(filter(None, [
        subject,
        text_body,
        extract_text_from_html(html_body),
    ]))

    codes = extract_promo_codes(full_text)
    return codes, sender, False


def main():
    parser = argparse.ArgumentParser(description="Extract and validate promo codes from emails.")
    parser.add_argument("target", nargs="?", default=None,
                        help="Email file (.eml/.html) or directory of email files")
    parser.add_argument("--mail", action="store_true",
                        help="Read emails directly from macOS Mail.app (requires --date)")
    parser.add_argument("--account", default=None,
                        help="Mail.app account name to read from (default: all accounts)")
    parser.add_argument("--mailbox", default="INBOX",
                        help="Mail.app mailbox to read from (default: INBOX)")
    parser.add_argument("--date", help="Only process emails received on this date (YYYY-MM-DD)", default=None)
    parser.add_argument("--skip-validation", action="store_true",
                        help="Only extract codes, skip website validation")
    args = parser.parse_args()

    if not args.mail and not args.target:
        parser.print_help()
        print("\nError: Provide a file/directory path, or use --mail to read from Mail.app.")
        sys.exit(1)

    filter_date = None
    if args.date:
        try:
            filter_date = datetime.strptime(args.date, "%Y-%m-%d").date()
        except ValueError:
            print(f"Error: Invalid date format '{args.date}'. Use YYYY-MM-DD.")
            sys.exit(1)

    if args.mail and not filter_date:
        print("Error: --date is required when using --mail (to avoid fetching thousands of emails).")
        sys.exit(1)

    # code -> {"sources": [filenames], "sender": first sender seen, "site": site key}
    all_codes = {}

    if args.mail:
        # --- Read from macOS Mail.app ---
        print(f"Reading from Mail.app — date: {filter_date}, "
              f"account: {args.account or 'all'}, mailbox: {args.mailbox}\n")

        messages = fetch_messages_from_mail(
            target_date=filter_date,
            account=args.account,
            mailbox_name=args.mailbox,
        )

        print(f"\nFetched {len(messages)} message(s). Extracting promo codes...\n")

        for msg_data in messages:
            subject = msg_data["subject"]
            sender = msg_data["sender"]
            source = msg_data["source"]

            # Parse the raw email source
            parsed_subject, parsed_sender, text_body, html_body = parse_mail_source(source)
            # Use AppleScript subject/sender as fallback (more reliable for encoding)
            display_sender = sender or parsed_sender
            display_subject = subject or parsed_subject

            full_text = " ".join(filter(None, [
                display_subject,
                text_body,
                extract_text_from_html(html_body),
            ]))

            codes = extract_promo_codes(full_text)
            short_subject = display_subject[:70]
            if codes:
                print(f"  Found codes: {', '.join(codes)}  <- {short_subject}")
                site = get_site_from_sender(display_sender)
                for code in codes:
                    source_label = f"{short_subject} ({msg_data['account']})"
                    if code not in all_codes:
                        all_codes[code] = {"sources": [], "sender": display_sender, "site": site}
                    all_codes[code]["sources"].append(source_label)

    else:
        # --- Read from files ---
        if filter_date:
            print(f"Filtering emails for date: {filter_date}\n")

        target = args.target
        if os.path.isdir(target):
            files = sorted(Path(target).glob("*"))
        else:
            files = [Path(target)]

        processed = 0
        skipped = 0

        for filepath in files:
            if filepath.is_file() and filepath.suffix.lower() in (".eml", ".html", ".htm"):
                codes, sender, was_skipped = process_email_file(filepath, filter_date)
                if was_skipped:
                    skipped += 1
                    continue
                processed += 1
                print(f"Processing: {filepath.name}")
                if codes:
                    print(f"  Found codes: {', '.join(codes)}")
                    site = get_site_from_sender(sender)
                    for code in codes:
                        if code not in all_codes:
                            all_codes[code] = {"sources": [], "sender": sender, "site": site}
                        all_codes[code]["sources"].append(filepath.name)
                else:
                    print("  No promo codes found.")

        if filter_date:
            print(f"\n{processed} email(s) matched date {filter_date}, {skipped} skipped.")

    if not all_codes:
        print("\nNo promotion codes found in any emails.")
        return

    # Deduplicate and validate
    print(f"\n{'='*50}")
    print(f"Found {len(all_codes)} unique promo code(s):\n")

    for code, info in sorted(all_codes.items()):
        site = info["site"]
        sources = info["sources"]
        print(f"  Code: {code}")
        print(f"    Site:    {site or 'unknown'}")
        print(f"    Source:  {', '.join(sources)}")

        if args.skip_validation:
            print(f"    Status:  SKIPPED (--skip-validation)")
        else:
            valid, message = validate_promo_code(code, site)
            if valid is None:
                status = "UNKNOWN"
            elif valid:
                status = "VALID"
            else:
                status = "EXPIRED/INVALID"

            print(f"    Status:  {status}")
            print(f"    Detail:  {message}")
        print()


if __name__ == "__main__":
    main()
