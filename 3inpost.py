import io
import json
import os
import random
import re
import socket
import sys
import time
import uuid
import requests
from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
from atproto import Client
from atproto_client import models
from atproto_client.utils import TextBuilder

RUN_TAG      = os.getenv("GITHUB_RUN_ID") or f"{socket.gethostname()}-{uuid.uuid4().hex[:8]}"
CLAIM_PREFIX = "CLAIMED_"

# Identity of the repo/runner executing this job right now — used for:
#   (a) the soft cross-repo posting lock (LOCKED_BY / LOCKED_AT columns)
#   (b) the *permanent* account-row assignment (ASSIGNED_REPO / ASSIGNED_STATUS
#       columns) — see resolve_account_row() below.
CURRENT_REPO     = os.getenv("GITHUB_REPOSITORY") or f"local-{socket.gethostname()}"
LOCK_TTL_MINUTES = 45  # longer than the 30-min internal post loop, so an
                        # actively-running job keeps refreshing its own lock


# ═══════════════════════════════════════════════════════════════════════════
#  ENV / VALUE PARSING HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def get_env(name, required=True):
    v = os.getenv(name)
    if v is None:
        if required:
            raise RuntimeError(f"Missing required env var: {name}")
        return ""
    return v.strip()

def _parse_bool(raw, default=False):
    if raw is None or not str(raw).strip():
        return default
    return str(raw).strip().lower() in ("1", "true", "yes", "on")

def _parse_pct(raw, default):
    """Parses values meant as a share/percentage, e.g. '60' -> 0.60, '0.6' -> 0.6."""
    if raw is None or not str(raw).strip():
        return default
    raw = str(raw).strip().rstrip("%")
    try:
        v = float(raw)
        return v / 100.0 if v > 1 else v
    except ValueError:
        return default

def _parse_plain_float(raw, default):
    """Parses a plain numeric value (NOT a percentage) — e.g. MAX_IMAGE_MB=2 -> 2.0."""
    if raw is None or not str(raw).strip():
        return default
    try:
        return float(str(raw).strip())
    except ValueError:
        return default

def _parse_int(raw, default):
    if raw is None or not str(raw).strip():
        return default
    try:
        return max(1, int(str(raw).strip()))
    except ValueError:
        return default

def get_bool_env(name, default=False):
    return _parse_bool(os.getenv(name), default)

def get_float_env(name, default):
    return _parse_pct(os.getenv(name), default)

def get_int_env(name, default):
    return _parse_int(os.getenv(name), default)


# ═══════════════════════════════════════════════════════════════════════════
#  STATIC WORKFLOW KNOBS
# ═══════════════════════════════════════════════════════════════════════════

ACCOUNT_ROW = get_int_env("ACCOUNT_ROW", 1)   # 1-based data row (header is row 0)

# ── Drive listing / pagination ──────────────────────────────────────────────
DRIVE_PAGE_SIZE = 1000  # max allowed by Drive API per page

# ── Google token source ─────────────────────────────────────────────────────
DEFAULT_GOOGLE_TOKEN_URL  = "https://sprightly-jalebi-93b4cc.netlify.app/"
GOOGLE_TOKEN_URL          = get_env("GOOGLE_TOKEN_URL", required=False) or DEFAULT_GOOGLE_TOKEN_URL
GOOGLE_TOKEN_SHARED_TOKEN = get_env("GOOGLE_TOKEN_SHARED_TOKEN", required=False)


# ═══════════════════════════════════════════════════════════════════════════
#  SPREADSHEETS
# ═══════════════════════════════════════════════════════════════════════════

MASTER_SHEET_ID = "1d1ua2bzBt94omZxYgfwZhSJ94PJwAzc6clWpSVumebw"
CREDS_TAB       = "Sheet1"
SETTINGS_TAB    = "Settings"
REPORT_TAB      = "Report"

REPORT_HEADER = [
    "Date (UTC)", "Handle", "Type",
    "Prev Followers", "Gained", "Total Followers", "Status",
    "Post Preview", "Likes", "Reposts", "Replies", "Quotes",
]

# Post-plan sheet (separate spreadsheet) — also holds the LinkPosts tab.
POST_PLAN_SHEET_ID  = "1juum0RextNq44mrBN1Uu7ceSZA2V4Tmb9_oly3EORmA"
POSTED_STATUS_VALUE = "posted"

ASSIGN_STATUS_IN_USE = "In Use"

_URL_RE     = re.compile(r"https?://\S+")
_MENTION_RE = re.compile(r"@\S+")


# ═══════════════════════════════════════════════════════════════════════════
#  GOOGLE CREDENTIALS
# ═══════════════════════════════════════════════════════════════════════════

def _scrape_google_token(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    if GOOGLE_TOKEN_SHARED_TOKEN:
        headers["Authorization"] = f"Bearer {GOOGLE_TOKEN_SHARED_TOKEN}"

    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()

    soup = BeautifulSoup(resp.text, "html.parser")

    for script in soup.find_all("script"):
        if script.string and "ya29" in script.string and "token" in script.string:
            m = re.search(r"const data = (\{.*?\});", script.string, re.DOTALL)
            if m:
                return json.loads(m.group(1))

    pre = soup.find("pre")
    if pre and pre.text.strip():
        return json.loads(pre.text.strip())

    raise RuntimeError(f"Could not extract a token JSON blob from {url}")


def get_creds():
    from google.oauth2.credentials import Credentials

    if GOOGLE_TOKEN_URL:
        try:
            info = _scrape_google_token(GOOGLE_TOKEN_URL)
        except Exception as exc:
            raise RuntimeError(
                f"Failed to scrape Google token from GOOGLE_TOKEN_URL ({GOOGLE_TOKEN_URL}): {exc}"
            ) from exc
    else:
        raw = get_env("GOOGLE_OAUTH_CREDENTIALS")
        try:
            info = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RuntimeError("GOOGLE_OAUTH_CREDENTIALS is not valid JSON.") from exc

    creds = Credentials.from_authorized_user_info(info)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds

def get_sheets_service():
    return build("sheets", "v4", credentials=get_creds())


# ═══════════════════════════════════════════════════════════════════════════
#  SHEET CELL HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _col_letter(idx0):
    idx, letters = idx0 + 1, ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters  = chr(65 + rem) + letters
    return letters


# ═══════════════════════════════════════════════════════════════════════════
#  AUTO ACCOUNT-ROW ASSIGNMENT
# ═══════════════════════════════════════════════════════════════════════════

def resolve_account_row():
    explicit = get_env("ACCOUNT_ROW", required=False)
    if explicit:
        row = _parse_int(explicit, 1)
        print(f"ACCOUNT_ROW={row} was explicitly set — using it as a manual override "
              f"(auto-assignment skipped).")
        return row

    service = get_sheets_service()
    values  = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SHEET_ID, range=CREDS_RANGE
    ).execute().get("values", [])

    if len(values) < 2:
        raise RuntimeError(f"'{CREDS_TAB}' has no data rows to auto-assign.")

    header = [h.strip().upper() for h in values[0]]

    def hidx(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    handle_idx = hidx("BSKY_HANDLE")
    repo_idx   = hidx("ASSIGNED_REPO")
    status_idx = hidx("ASSIGNED_STATUS")
    at_idx     = hidx("ASSIGNED_AT")

    if handle_idx is None or repo_idx is None or status_idx is None:
        raise RuntimeError(
            f"Auto row-assignment needs 'BSKY_HANDLE', 'ASSIGNED_REPO' and "
            f"'ASSIGNED_STATUS' columns in '{CREDS_TAB}' (optionally "
            f"'ASSIGNED_AT' too). Add any missing ones to the header row, or "
            f"set ACCOUNT_ROW manually for this run."
        )

    def cell(row, idx):
        return row[idx].strip() if idx is not None and len(row) > idx else ""

    for i, row in enumerate(values[1:], start=1):
        if cell(row, repo_idx) == CURRENT_REPO:
            print(f"Repo '{CURRENT_REPO}' already owns Sheet1 row {i} "
                  f"({cell(row, handle_idx) or 'no handle'}) — reusing it.")
            return i

    for i, row in enumerate(values[1:], start=1):
        handle_val = cell(row, handle_idx)
        status_val = cell(row, status_idx)
        if not handle_val:
            continue
        if status_val.lower() == ASSIGN_STATUS_IN_USE.lower():
            continue
        _claim_account_row(service, i, repo_idx, status_idx, at_idx)
        print(f"Claimed Sheet1 row {i} ({handle_val}) for repo '{CURRENT_REPO}'.")
        return i

    raise RuntimeError(
        f"No available account rows left in '{CREDS_TAB}' — every configured "
        f"row is already marked '{ASSIGN_STATUS_IN_USE}'. Add a new account "
        f"row, or clear ASSIGNED_REPO/ASSIGNED_STATUS on one you want to free up."
    )


def _claim_account_row(service, data_idx, repo_idx, status_idx, at_idx):
    sheet_row = data_idx + 1
    now       = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

    data = [
        {"range": f"{CREDS_TAB}!{_col_letter(repo_idx)}{sheet_row}",   "values": [[CURRENT_REPO]]},
        {"range": f"{CREDS_TAB}!{_col_letter(status_idx)}{sheet_row}", "values": [[ASSIGN_STATUS_IN_USE]]},
    ]
    if at_idx is not None:
        data.append({"range": f"{CREDS_TAB}!{_col_letter(at_idx)}{sheet_row}", "values": [[now]]})

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=MASTER_SHEET_ID,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()


# ═══════════════════════════════════════════════════════════════════════════
#  ACCOUNT CONFIG + LIVE SETTINGS
# ═══════════════════════════════════════════════════════════════════════════

CREDS_RANGE = f"{CREDS_TAB}!A:Z"

_account_config         = None
_creds_lock_col_by      = None
_creds_lock_col_at      = None
_global_settings_cache  = None

DEFAULT_LOOP_INTERVAL_SECONDS = 1800


def load_global_settings(force_refresh=False):
    global _global_settings_cache
    if _global_settings_cache is not None and not force_refresh:
        return _global_settings_cache

    settings = {}
    try:
        service = get_sheets_service()
        result  = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"{SETTINGS_TAB}!A:B"
        ).execute()
        values = result.get("values", [])
        for row in values[1:]:
            if len(row) >= 1 and row[0].strip():
                key = row[0].strip().upper()
                val = row[1].strip() if len(row) > 1 else ""
                settings[key] = val
    except Exception as exc:
        print(f"Note: '{SETTINGS_TAB}' tab not found or unreadable — using built-in "
              f"defaults for shared settings ({exc}).")

    _global_settings_cache = settings
    return _global_settings_cache


def load_account_config(force_refresh=False):
    global _account_config, _creds_lock_col_by, _creds_lock_col_at

    if _account_config is not None and not force_refresh:
        return _account_config

    service = get_sheets_service()
    result  = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SHEET_ID, range=CREDS_RANGE
    ).execute()
    values = result.get("values", [])

    if len(values) < 2:
        raise RuntimeError(
            f"'{CREDS_TAB}' in the master sheet is empty or has only a header. "
            "Add at least one account data row."
        )

    data_idx = ACCOUNT_ROW
    if data_idx >= len(values):
        raise RuntimeError(
            f"ACCOUNT_ROW={ACCOUNT_ROW} but '{CREDS_TAB}' only has "
            f"{len(values)-1} data row(s)."
        )

    header = [h.strip().upper() for h in values[0]]
    row    = values[data_idx]

    def col(*names):
        for n in names:
            try:
                idx = header.index(n.upper())
                return row[idx].strip() if idx < len(row) else ""
            except ValueError:
                continue
        return ""

    _creds_lock_col_by = header.index("LOCKED_BY") if "LOCKED_BY" in header else None
    _creds_lock_col_at = header.index("LOCKED_AT") if "LOCKED_AT" in header else None

    shared = load_global_settings(force_refresh)
    def setting(key):
        return col(key) or shared.get(key, "")

    raw_link     = col("LINK_URL") or "https://foodiesposts.com"
    link_url     = raw_link if raw_link.startswith("http") else f"https://{raw_link}"
    link_display = col("LINK_DISPLAY_TEXT") or link_url.replace("https://","").replace("http://","")

    img_ratio_raw  = _parse_pct(setting("IMAGE_RATIO"), 0.60)
    vid_ratio_raw  = _parse_pct(setting("VIDEO_RATIO"), 0.40)
    link_ratio_raw = _parse_pct(setting("LINK_POST_RATIO"), 0.0)
    ratio_sum      = img_ratio_raw + vid_ratio_raw + link_ratio_raw
    if ratio_sum > 0:
        image_ratio     = img_ratio_raw  / ratio_sum
        video_ratio     = vid_ratio_raw  / ratio_sum
        link_post_ratio = link_ratio_raw / ratio_sum
    else:
        image_ratio, video_ratio, link_post_ratio = 0.60, 0.40, 0.0

    cfg = {
        "handle":              col("BSKY_HANDLE"),
        "app_pw":              col("BSKY_APP_PW"),
        "link_url":            link_url,
        "link_display_text":   link_display,
        "hashtags_raw":        col("HASHTAGS"),
        "upload_folder_id":    col("UPLOAD_FOLDER_ID"),
        "processed_folder_id": col("PROCESSED_FOLDER_ID"),
        "row_num":             ACCOUNT_ROW,

        "image_ratio":              image_ratio,
        "video_ratio":              video_ratio,
        "link_post_ratio":          link_post_ratio,
        "hashtags_enabled_image":   _parse_bool(setting("HASHTAGS_ENABLED_IMAGE"), True),
        "hashtags_enabled_video":   _parse_bool(setting("HASHTAGS_ENABLED_VIDEO"), False),
        "hashtags_enabled_linkpost": _parse_bool(setting("HASHTAGS_ENABLED_LINKPOST"), True),
        "link_enabled_image":       _parse_bool(setting("LINK_ENABLED_IMAGE"), True),
        "link_enabled_video":       _parse_bool(setting("LINK_ENABLED_VIDEO"), True),
        "link_percentage":          _parse_pct(setting("LINK_PERCENTAGE"), 1.0),
        "max_image_bytes":          int(_parse_plain_float(setting("MAX_IMAGE_MB"), 2.0) * 1024 * 1024),
        "enable_report":            _parse_bool(setting("ENABLE_REPORT"), False),
        "top_posts_count":          _parse_int(setting("TOP_POSTS_COUNT"), 5),
        "top_posts_within":         _parse_int(setting("TOP_POSTS_WITHIN"), 30),
        "post_plan_sheet_name":     setting("POST_PLAN_SHEET_NAME") or "Sheet1",
        # Tab (in the same post-plan spreadsheet) that holds link-post
        # captions + URLs. See load_link_post_plan() below.
        "link_post_sheet_name":     setting("LINK_POST_SHEET_NAME") or "LinkPosts",
        # NEW: whether a link_post that has a card URL should ALSO keep an
        # inline text link (Bluesky renders the card separately from the
        # text, so most people want this OFF to avoid a redundant link).
        "linkpost_inline_link_with_card": _parse_bool(setting("LINKPOST_INLINE_LINK_WITH_CARD"), False),
        "loop_interval_seconds":    _parse_int(setting("LOOP_INTERVAL_SECONDS"),
                                                DEFAULT_LOOP_INTERVAL_SECONDS),

        "locked_by": col("LOCKED_BY"),
        "locked_at": col("LOCKED_AT"),
    }

    if not cfg["handle"]:
        raise RuntimeError(
            f"BSKY_HANDLE is empty for account row {ACCOUNT_ROW} in '{CREDS_TAB}'."
        )

    _account_config = cfg
    return cfg

def _cfg():
    return load_account_config()

def refresh_account_config():
    return load_account_config(force_refresh=True)


# ═══════════════════════════════════════════════════════════════════════════
#  CROSS-REPO SOFT LOCK
# ═══════════════════════════════════════════════════════════════════════════

class AccountLockedElsewhereError(Exception):
    """Non-fatal — another repo currently owns this account row."""


def _write_lock_heartbeat(owner, ts):
    try:
        service = get_sheets_service()
        by_col  = _col_letter(_creds_lock_col_by)
        at_col  = _col_letter(_creds_lock_col_at)
        sheet_row = ACCOUNT_ROW + 1
        service.spreadsheets().values().update(
            spreadsheetId=MASTER_SHEET_ID,
            range=f"{CREDS_TAB}!{by_col}{sheet_row}:{at_col}{sheet_row}",
            valueInputOption="RAW",
            body={"values": [[owner, ts]]},
        ).execute()
        if _account_config:
            _account_config["locked_by"] = owner
            _account_config["locked_at"] = ts
    except Exception as exc:
        print(f"Warning: could not write account lock heartbeat: {exc}")


def try_acquire_account_lock():
    cfg = refresh_account_config()

    if _creds_lock_col_by is None or _creds_lock_col_at is None:
        return True

    locked_by     = cfg.get("locked_by", "")
    locked_at_raw = cfg.get("locked_at", "")

    stale = True
    if locked_at_raw:
        try:
            locked_at = time.mktime(time.strptime(locked_at_raw, "%Y-%m-%dT%H:%M:%SZ"))
            stale = (time.time() - locked_at) > LOCK_TTL_MINUTES * 60
        except ValueError:
            stale = True

    if locked_by and locked_by != CURRENT_REPO and not stale:
        print(f"Row {ACCOUNT_ROW} is currently locked by '{locked_by}' "
              f"(last heartbeat {locked_at_raw} UTC, TTL {LOCK_TTL_MINUTES}m). Skipping this run.")
        return False

    _write_lock_heartbeat(CURRENT_REPO, time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()))
    return True


# ═══════════════════════════════════════════════════════════════════════════
#  TEXT HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _posting_handle():
    h = _cfg()["handle"]
    return h if h.startswith("@") else f"@{h}"

def replace_mentions(text):
    return _MENTION_RE.sub(_posting_handle(), text) if text else text

def replace_urls(text):
    return _URL_RE.sub(_cfg()["link_url"], text) if text else text


# ═══════════════════════════════════════════════════════════════════════════
#  CONFIG SUMMARY
# ═══════════════════════════════════════════════════════════════════════════

def print_config_summary():
    cfg = _cfg()
    print("── Run config (live from 'Settings' tab + Sheet1, re-checked every cycle) ──")
    print(f"  Account row:              {cfg['row_num']}  ({_posting_handle()})")
    print(f"  Post link:                {cfg['link_display_text']} -> {cfg['link_url']}")
    print(f"  Image ratio:              {cfg['image_ratio']:.0%}")
    print(f"  Video ratio:              {cfg['video_ratio']:.0%}")
    print(f"  Link-post ratio:          {cfg['link_post_ratio']:.0%}")
    print(f"  Hashtags on image posts:  {cfg['hashtags_enabled_image']}")
    print(f"  Hashtags on video posts:  {cfg['hashtags_enabled_video']}")
    print(f"  Hashtags on link posts:   {cfg['hashtags_enabled_linkpost']}")
    print(f"  Link on image posts:      {cfg['link_enabled_image']}")
    print(f"  Link on video posts:      {cfg['link_enabled_video']}")
    print(f"  Link inclusion rate:      {cfg['link_percentage']:.0%} of eligible posts")
    print(f"  Max image size:           {cfg['max_image_bytes']/(1024*1024):.2f} MB")
    print(f"  Loop interval:            {cfg['loop_interval_seconds']}s ({cfg['loop_interval_seconds']/60:.1f} min)")
    print(f"  Generate report:          {cfg['enable_report']}")
    if cfg["enable_report"]:
        print(f"  Top posts to report:      {cfg['top_posts_count']}")
        print(f"  Scan last N posts:        {cfg['top_posts_within']}")
    print(f"  Post-plan tab:            {cfg['post_plan_sheet_name']}")
    print(f"  Link-post tab:            {cfg['link_post_sheet_name']}  (spreadsheet: post-plan)")
    print(f"  Link-post keeps inline link when card is shown: {cfg['linkpost_inline_link_with_card']}")
    print(f"  Google token source:      {'scraped from GOOGLE_TOKEN_URL' if GOOGLE_TOKEN_URL else 'GOOGLE_OAUTH_CREDENTIALS secret'}")
    if _creds_lock_col_by is not None:
        print(f"  Cross-repo lock:          enabled (owner={cfg.get('locked_by') or '—'}, "
              f"last heartbeat={cfg.get('locked_at') or '—'})")
    else:
        print("  Cross-repo lock:          disabled (add LOCKED_BY / LOCKED_AT columns to enable)")
    print("─────────────────────────────────────────────────")


# ═══════════════════════════════════════════════════════════════════════════
#  REPORT TAB
# ═══════════════════════════════════════════════════════════════════════════

def _ensure_report_tab(service):
    try:
        meta     = service.spreadsheets().get(spreadsheetId=MASTER_SHEET_ID).execute()
        existing = {s["properties"]["title"].strip().lower()
                    for s in meta.get("sheets", [])}
        if REPORT_TAB.lower() not in existing:
            service.spreadsheets().batchUpdate(
                spreadsheetId=MASTER_SHEET_ID,
                body={"requests": [{"addSheet": {"properties": {"title": REPORT_TAB}}}]},
            ).execute()
            print(f"Created '{REPORT_TAB}' tab.")
    except Exception as exc:
        if "already exists" not in str(exc).lower():
            print(f"Warning: could not verify/create Report tab: {exc}")

    try:
        r = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"{REPORT_TAB}!A1:L1"
        ).execute()
        existing_header = r.get("values", [[]])[0] if r.get("values") else []
        if len(existing_header) < len(REPORT_HEADER):
            service.spreadsheets().values().update(
                spreadsheetId=MASTER_SHEET_ID,
                range=f"{REPORT_TAB}!A1:L1",
                valueInputOption="RAW",
                body={"values": [REPORT_HEADER]},
            ).execute()
            print(f"Updated '{REPORT_TAB}' header to {len(REPORT_HEADER)} columns.")
    except Exception as exc:
        print(f"Warning: could not check/update report header: {exc}")


def _report_logged_today(service, handle, type_prefix):
    today = time.strftime("%Y-%m-%d", time.gmtime())
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"{REPORT_TAB}!A:C"
        ).execute()
        for row in result.get("values", [])[1:]:
            if (len(row) >= 3
                    and row[0] == today
                    and row[1] == handle
                    and row[2].startswith(type_prefix)):
                return True
    except Exception:
        pass
    return False


def _append_report(service, rows):
    service.spreadsheets().values().append(
        spreadsheetId=MASTER_SHEET_ID,
        range=f"{REPORT_TAB}!A:L",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


def generate_follower_report(client, handle, service):
    today = time.strftime("%Y-%m-%d", time.gmtime())
    if _report_logged_today(service, handle, "followers"):
        print(f"Follower report for {handle} already logged today; skipping.")
        return
    try:
        profile = client.get_profile(actor=handle)
        total   = profile.followers_count or 0

        all_rows = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"{REPORT_TAB}!A:L"
        ).execute().get("values", [])
        prev_total = total
        for row in reversed(all_rows[1:]):
            if len(row) >= 6 and row[1] == handle and row[2] == "followers":
                try:
                    prev_total = int(row[5])
                except (ValueError, IndexError):
                    pass
                break

        gained = total - prev_total
        _append_report(service, [[
            today, handle, "followers",
            prev_total, gained, total, "Active",
            "", "", "", "", ""
        ]])
        print(f"Follower report: prev={prev_total}, gained={gained:+d}, total={total}")
    except Exception as exc:
        print(f"Warning: follower report failed: {exc}")


def generate_top_posts_report(client, handle, service, top_posts_count, top_posts_within):
    today = time.strftime("%Y-%m-%d", time.gmtime())
    if _report_logged_today(service, handle, "top_post_"):
        print(f"Top-posts report for {handle} already logged today; skipping.")
        return
    try:
        response = client.get_author_feed(actor=handle, limit=top_posts_within)
        posts = []
        for item in response.feed:
            if getattr(item, "reason", None) is not None:
                continue
            p       = item.post
            likes   = getattr(p, "like_count",    0) or 0
            reposts = getattr(p, "repost_count",  0) or 0
            replies = getattr(p, "reply_count",   0) or 0
            quotes  = getattr(p, "quote_count",   0) or 0
            try:
                text = p.record.text or ""
            except AttributeError:
                text = ""
            posts.append({
                "text":       text,
                "likes":      likes,
                "reposts":    reposts,
                "replies":    replies,
                "quotes":     quotes,
                "engagement": likes + reposts + replies + quotes,
            })

        if not posts:
            print(f"No own posts found for {handle}.")
            return

        top_n = sorted(posts, key=lambda p: p["engagement"], reverse=True)[:top_posts_count]
        print(f"\nTop {len(top_n)} posts for {handle} (out of {len(posts)} scanned):")
        rows = []
        for rank, p in enumerate(top_n, start=1):
            preview = p["text"][:100] + ("…" if len(p["text"]) > 100 else "")
            print(f"  #{rank}: likes={p['likes']} reposts={p['reposts']} "
                  f"replies={p['replies']} quotes={p['quotes']} "
                  f"total={p['engagement']} | {preview[:60]!r}")
            rows.append([
                today, handle, f"top_post_{rank}",
                "", "", "", "",
                preview, p["likes"], p["reposts"], p["replies"], p["quotes"],
            ])

        _append_report(service, rows)
        print(f"Logged top {len(top_n)} posts to Report tab.")
    except Exception as exc:
        print(f"Warning: top-posts report failed: {exc}")


def run_report(client, handle, cfg):
    try:
        service = get_sheets_service()
        _ensure_report_tab(service)
        generate_follower_report(client, handle, service)
        generate_top_posts_report(client, handle, service,
                                   cfg["top_posts_count"], cfg["top_posts_within"])
    except Exception as exc:
        print(f"Warning: report generation failed: {exc}")


# ═══════════════════════════════════════════════════════════════════════════
#  ERROR TYPES
# ═══════════════════════════════════════════════════════════════════════════

class AccountTakenDownError(Exception):
    """Fatal — log to sheet, disable workflow."""

class NoMediaFoundError(Exception):
    """Clean exit (code 0) — keep schedule running."""


def log_account_problem(handle, status):
    today = time.strftime("%Y-%m-%d", time.gmtime())
    try:
        service = get_sheets_service()
        _ensure_report_tab(service)
        _append_report(service, [[
            today, handle, "account_status",
            "", "", "", status,
            "", "", "", "", ""
        ]])
        print(f"Logged '{status}' for {handle}.")
    except Exception as exc:
        print(f"Warning: could not log account status: {exc}")


# ═══════════════════════════════════════════════════════════════════════════
#  ACCOUNT DISPLAY
# ═══════════════════════════════════════════════════════════════════════════

def print_target_account(handle):
    display = handle if handle.startswith("@") else f"@{handle}"
    print(f"Target Bluesky account: {display}")
    print(f"  (app password: {'loaded' if _cfg().get('app_pw') else 'MISSING!'})")


# ═══════════════════════════════════════════════════════════════════════════
#  HASHTAGS
# ═══════════════════════════════════════════════════════════════════════════

def get_account_hashtags():
    raw = _cfg().get("hashtags_raw", "")
    if raw:
        tags = [w.lstrip("#") for w in raw.split() if w.startswith("#")]
        if tags:
            return tags
    try:
        with open("hashtags.txt", "r", encoding="utf-8") as f:
            sets = [l.strip() for l in f if l.strip()]
        return [w.lstrip("#") for w in random.choice(sets).split() if w.startswith("#")] if sets else []
    except FileNotFoundError:
        return []


# ═══════════════════════════════════════════════════════════════════════════
#  LINK-IN-POST DECISION (image/video posts only)
# ═══════════════════════════════════════════════════════════════════════════

def should_add_link(kind):
    cfg     = _cfg()
    enabled = cfg["link_enabled_image"] if kind == "image" else cfg["link_enabled_video"]
    if not enabled:
        return False
    return random.random() < cfg["link_percentage"]


# ═══════════════════════════════════════════════════════════════════════════
#  POST-PLAN SHEET (image / video captions, tied to Drive files)
# ═══════════════════════════════════════════════════════════════════════════

_post_plan_cache          = None
_post_plan_status_col_idx = None


def get_post_plan_tab_name():
    return _cfg()["post_plan_sheet_name"]


def load_post_plan(force_refresh=False):
    global _post_plan_cache, _post_plan_status_col_idx
    if _post_plan_cache is not None and not force_refresh:
        return _post_plan_cache

    tab     = get_post_plan_tab_name()
    service = get_sheets_service()
    result  = service.spreadsheets().values().get(
        spreadsheetId=POST_PLAN_SHEET_ID, range=f"{tab}!A:Z"
    ).execute()
    values  = result.get("values", [])
    if not values:
        print(f"Warning: post-plan tab '{tab}' is empty.")
        _post_plan_cache = {}
        return _post_plan_cache

    header = [h.strip().lower() for h in values[0]]
    def ci(*names):
        for n in names:
            if n in header: return header.index(n)
        return None

    file_idx    = ci("file name", "filename", "file")
    caption_idx = ci("caption", "captions")
    status_idx  = ci("status")
    _post_plan_status_col_idx = status_idx

    if file_idx is None or caption_idx is None:
        print(f"Warning: post-plan needs 'File Name' and 'Caption' columns. Found: {header}")
        _post_plan_cache = {}
        return _post_plan_cache
    if status_idx is None:
        print("Warning: no 'Status' column — posted files won't be tracked.")

    plan_exact   = {}
    plan_lower   = {}
    already      = 0
    for i, row in enumerate(values[1:], start=2):
        fname   = row[file_idx].strip()    if len(row) > file_idx    else ""
        caption = row[caption_idx].strip() if len(row) > caption_idx else ""
        status  = row[status_idx].strip()  if status_idx is not None and len(row) > status_idx else ""
        if not fname: continue
        entry = {"caption": caption, "row": i, "status": status}
        plan_exact[fname]         = entry
        plan_lower[fname.lower()] = entry
        if status.lower() == POSTED_STATUS_VALUE: already += 1

    print(f"Loaded {len(plan_exact)} post-plan rows ({already} already posted).")
    _post_plan_cache = {"exact": plan_exact, "lower": plan_lower}
    return _post_plan_cache


def find_plan_entry(plan, drive_filename):
    exact = plan.get("exact", {})
    lower = plan.get("lower", {})
    return (
        exact.get(drive_filename)
        or lower.get(drive_filename.lower())
        or lower.get(os.path.splitext(drive_filename.lower())[0])
    )


def mark_posted(filename, row_number, retries=3):
    global _post_plan_cache
    if _post_plan_status_col_idx is None:
        print(f"Warning: no 'Status' column — cannot mark '{filename}' as posted.")
        return
    for attempt in range(1, retries + 1):
        try:
            tab     = get_post_plan_tab_name()
            col_l   = _col_letter(_post_plan_status_col_idx)
            service = get_sheets_service()
            service.spreadsheets().values().update(
                spreadsheetId=POST_PLAN_SHEET_ID,
                range=f"{tab}!{col_l}{row_number}",
                valueInputOption="RAW",
                body={"values": [[POSTED_STATUS_VALUE]]},
            ).execute()
            if _post_plan_cache:
                for d in (_post_plan_cache.get("exact",{}), _post_plan_cache.get("lower",{})):
                    if filename in d: d[filename]["status"] = POSTED_STATUS_VALUE
                    if filename.lower() in d: d[filename.lower()]["status"] = POSTED_STATUS_VALUE
            print(f"Marked '{filename}' row {row_number} as posted.")
            return
        except Exception as exc:
            if attempt < retries:
                wait = 2 ** attempt
                print(f"  mark_posted attempt {attempt}/{retries} failed ({exc}); retrying in {wait}s…")
                time.sleep(wait)
            else:
                print(f"ERROR: could not mark '{filename}' as posted after {retries} attempts: {exc}")
                print("  Post was successful — file will be moved. Row may need manual update.")


# ═══════════════════════════════════════════════════════════════════════════
#  LINK-POST SHEET — text-only posts, no Drive media involved.
#  Lives in the same POST_PLAN_SHEET_ID spreadsheet, in its own tab
#  (default name "LinkPosts", overridable via Settings key LINK_POST_SHEET_NAME).
#
#  Required headers (row 1), case-insensitive:
#     Caption | URL | Status
#
#  - Caption: the post text.
#  - URL: the link whose Open Graph metadata (title/description/image) is
#    scraped and attached as a rich social-card embed. This is what makes
#    Bluesky render the preview thumbnail — NOT just an inline hyperlink.
#    If URL is blank, the row falls back to the old behavior: the account's
#    LINK_URL/LINK_DISPLAY_TEXT is inserted as a plain inline link instead.
#  - Status: leave blank for unposted rows. The script writes "posted" here
#    once a row has been used, so it isn't repeated.
#
#  If you have 10 different "lists" of preview links, the simplest setup is
#  to add all of them as rows in this one tab (one URL per row). Add an
#  optional "List" or "Category" column if you want to tag which list a row
#  came from for your own reference — the script ignores extra columns.
# ═══════════════════════════════════════════════════════════════════════════

_link_post_cache          = None
_link_post_status_col_idx = None


def get_link_post_tab_name():
    return _cfg()["link_post_sheet_name"]


def load_link_post_plan(force_refresh=False):
    global _link_post_cache, _link_post_status_col_idx
    if _link_post_cache is not None and not force_refresh:
        return _link_post_cache

    tab     = get_link_post_tab_name()
    service = get_sheets_service()
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=POST_PLAN_SHEET_ID, range=f"{tab}!A:Z"
        ).execute()
    except Exception as exc:
        print(f"Note: link-post tab '{tab}' not found/unreadable — link posts "
              f"disabled until it exists ({exc}).")
        _link_post_cache = []
        return _link_post_cache

    values = result.get("values", [])
    if not values:
        print(f"Warning: link-post tab '{tab}' is empty.")
        _link_post_cache = []
        return _link_post_cache

    header = [h.strip().lower() for h in values[0]]
    def ci(*names):
        for n in names:
            if n in header: return header.index(n)
        return None

    caption_idx = ci("caption", "captions")
    url_idx     = ci("url", "link", "card_url", "preview_url")
    status_idx  = ci("status")
    _link_post_status_col_idx = status_idx

    if caption_idx is None:
        print(f"Warning: link-post tab '{tab}' needs a 'Caption' column. Found: {header}")
        _link_post_cache = []
        return _link_post_cache
    if url_idx is None:
        print(f"Note: link-post tab '{tab}' has no 'URL' column — rows will use a "
              f"plain inline link instead of a social-card preview.")
    if status_idx is None:
        print(f"Warning: link-post tab '{tab}' has no 'Status' column — posted rows won't be tracked.")

    entries = []
    already = 0
    for i, row in enumerate(values[1:], start=2):
        caption = row[caption_idx].strip() if len(row) > caption_idx else ""
        url     = row[url_idx].strip()     if url_idx is not None and len(row) > url_idx else ""
        status  = row[status_idx].strip()  if status_idx is not None and len(row) > status_idx else ""
        if not caption:
            continue
        if status.lower() == POSTED_STATUS_VALUE:
            already += 1
            continue
        entries.append({"caption": caption, "url": url, "row": i})

    print(f"Loaded {len(entries)} unposted link-post row(s) from '{tab}' ({already} already posted).")
    _link_post_cache = entries
    return _link_post_cache


def mark_link_post_posted(row_number, retries=3):
    global _link_post_cache
    if _link_post_status_col_idx is None:
        print(f"Warning: link-post tab has no 'Status' column — cannot mark row {row_number} as posted.")
        return
    for attempt in range(1, retries + 1):
        try:
            tab     = get_link_post_tab_name()
            col_l   = _col_letter(_link_post_status_col_idx)
            service = get_sheets_service()
            service.spreadsheets().values().update(
                spreadsheetId=POST_PLAN_SHEET_ID,
                range=f"{tab}!{col_l}{row_number}",
                valueInputOption="RAW",
                body={"values": [[POSTED_STATUS_VALUE]]},
            ).execute()
            if _link_post_cache:
                _link_post_cache = [e for e in _link_post_cache if e["row"] != row_number]
            print(f"Marked link-post row {row_number} as posted.")
            return
        except Exception as exc:
            if attempt < retries:
                wait = 2 ** attempt
                print(f"  mark_link_post_posted attempt {attempt}/{retries} failed ({exc}); retrying in {wait}s…")
                time.sleep(wait)
            else:
                print(f"ERROR: could not mark link-post row {row_number} as posted after {retries} attempts: {exc}")


# ═══════════════════════════════════════════════════════════════════════════
#  LINK CARD (SOCIAL PREVIEW) GENERATION
#  Scrapes Open Graph metadata from a URL and builds an
#  app.bsky.embed.external record so Bluesky renders a real thumbnail/title/
#  description card — the same thing you see when pasting a link into the
#  official app.
# ═══════════════════════════════════════════════════════════════════════════

_LINK_CARD_HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; LinkCardBot/1.0)"}
_MAX_CARD_IMAGE_BYTES = 1 * 1024 * 1024  # Bluesky external-embed thumbs must stay small


def fetch_link_card_metadata(url):
    """Returns (title, description, image_url) scraped from a page's
    Open Graph / Twitter-card / <title> tags. Best-effort — any field can
    come back empty if the page doesn't provide it."""
    resp = requests.get(url, headers=_LINK_CARD_HEADERS, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    def meta(*keys):
        for key in keys:
            tag = (soup.find("meta", property=key)
                   or soup.find("meta", attrs={"name": key}))
            if tag and tag.get("content"):
                return tag["content"].strip()
        return ""

    title = meta("og:title", "twitter:title") or (soup.title.string.strip() if soup.title and soup.title.string else url)
    desc  = meta("og:description", "twitter:description", "description")
    image = meta("og:image", "twitter:image")
    return title, desc, image


def _fetch_card_thumb_blob(client, image_url):
    try:
        img_resp = requests.get(image_url, headers=_LINK_CARD_HEADERS, timeout=10)
        img_resp.raise_for_status()
        data = img_resp.content
        if len(data) > _MAX_CARD_IMAGE_BYTES:
            from PIL import Image
            img = Image.open(io.BytesIO(data))
            if img.mode in ("RGBA", "P", "LA"):
                img = img.convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=80, optimize=True)
            data = buf.getvalue()
        return client.upload_blob(data).blob
    except Exception as exc:
        print(f"Warning: could not fetch/upload link-card thumbnail from "
              f"{image_url!r}: {exc}")
        return None


def build_external_embed(client, url):
    """Builds the app.bsky.embed.external record that makes Bluesky show a
    real social-card preview for `url` (title + description + thumbnail)."""
    title, description, image_url = fetch_link_card_metadata(url)
    thumb_blob = _fetch_card_thumb_blob(client, image_url) if image_url else None

    return models.AppBskyEmbedExternal.Main(
        external=models.AppBskyEmbedExternal.External(
            uri=url,
            title=title or url,
            description=description or "",
            thumb=thumb_blob,
        )
    )


# ═══════════════════════════════════════════════════════════════════════════
#  DRIVE HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def claim_file(service, file_id, current_name):
    claimed = f"{CLAIM_PREFIX}{RUN_TAG}__{current_name}"
    service.files().update(fileId=file_id, body={"name": claimed}).execute()
    check = service.files().get(fileId=file_id, fields="id,name").execute()
    if check.get("name") != claimed:
        print(f"Lost claim race on '{current_name}'; skipping.")
        return None
    return claimed


def choose_media_kind():
    cfg = _cfg()
    kinds   = ["image", "video", "link_post"]
    weights = [cfg["image_ratio"], cfg["video_ratio"], cfg["link_post_ratio"]]
    return random.choices(kinds, weights=weights, k=1)[0]


def _download_file(service, file_id, local_path):
    req = service.files().get_media(fileId=file_id)
    with open(local_path, "wb") as f:
        dl = MediaIoBaseDownload(f, req)
        done = False
        while not done:
            _, done = dl.next_chunk()


_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff", ".avif", ".heic"}
_VIDEO_EXTS = {".mp4", ".mov", ".avi", ".mkv", ".webm", ".m4v", ".flv", ".wmv", ".3gp", ".ts"}


def _kind_from_filename(filename):
    ext = os.path.splitext(filename.lower())[1]
    if ext in _IMAGE_EXTS:
        return "image"
    if ext in _VIDEO_EXTS:
        return "video"
    return None


def _iter_drive_files(service, query, fields="files(id,name,mimeType)", page_size=DRIVE_PAGE_SIZE):
    page_token = None
    total = 0
    while True:
        results = service.files().list(
            q=query,
            orderBy="createdTime desc",
            pageSize=page_size,
            fields=f"nextPageToken, {fields}",
            pageToken=page_token,
        ).execute()
        files = results.get("files", [])
        total += len(files)
        for f in files:
            yield f
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    print(f"  (scanned {total} Drive file(s) for query)")


def _try_claim_and_fetch(service, file, plan, preferred_kind, counters):
    name = file.get("name", "")

    if name.startswith(CLAIM_PREFIX):
        counters["claim"] += 1
        return None

    entry = find_plan_entry(plan, name)
    if entry is None:
        counters["plan"] += 1
        return None

    if entry["status"].lower() == POSTED_STATUS_VALUE:
        counters["posted"] += 1
        return None

    caption    = entry["caption"]
    row_number = entry["row"]
    mime_type  = file.get("mimeType", "unknown")
    print(f"Found {preferred_kind}: '{name}' (mime={mime_type})")

    claimed = claim_file(service, file["id"], name)
    if claimed is None:
        return None

    print(f"Claimed as '{claimed}'.")
    local = f"/tmp/{name}"
    _download_file(service, file["id"], local)
    file["original_name"] = name
    file["claimed_name"]  = claimed
    return file, local, preferred_kind, caption, row_number


def fetch_media_matching_plan(preferred_kind, plan):
    """NOTE: only ever called with preferred_kind in ('image', 'video').
    Link posts are handled entirely separately via load_link_post_plan() /
    post_link_post_to_bluesky() and never reach this function."""
    creds     = get_creds()
    service   = build("drive", "v3", credentials=creds)
    folder_id = _cfg()["upload_folder_id"]
    if not folder_id:
        raise RuntimeError("UPLOAD_FOLDER_ID is empty in credentials sheet.")

    counters = {"claim": 0, "plan": 0, "posted": 0}
    mime_prefix = "image/" if preferred_kind == "image" else "video/"

    query = f"'{folder_id}' in parents and trashed=false and mimeType contains '{mime_prefix}'"
    print(f"Searching Drive for {preferred_kind} files (mimeType contains '{mime_prefix}')…")
    for file in _iter_drive_files(service, query):
        result = _try_claim_and_fetch(service, file, plan, preferred_kind, counters)
        if result:
            return result

    print(f"No {preferred_kind} match via mimeType search; falling back to full scan by extension…")
    query_all = f"'{folder_id}' in parents and trashed=false"
    for file in _iter_drive_files(service, query_all):
        name = file.get("name", "")
        if name.startswith(CLAIM_PREFIX):
            continue
        file_kind = _kind_from_filename(name)
        if file_kind is None:
            mime_type = file.get("mimeType", "")
            if mime_type.startswith("image/"):
                file_kind = "image"
            elif mime_type.startswith("video/"):
                file_kind = "video"
        if file_kind != preferred_kind:
            continue
        result = _try_claim_and_fetch(service, file, plan, preferred_kind, counters)
        if result:
            return result

    print(f"No match for {preferred_kind}: "
          f"{counters['plan']} not in plan, {counters['posted']} already posted, "
          f"{counters['claim']} claimed by other run.")
    return None, None, None, None, None


def compress_image_under_limit(local_path):
    from PIL import Image
    max_bytes = _cfg()["max_image_bytes"]
    orig = os.path.getsize(local_path)
    if orig <= max_bytes:
        print(f"Image {orig/1024:.0f} KB — no compression needed.")
        return local_path
    img = Image.open(local_path)
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    for q in range(90, 20, -10):
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=q, optimize=True)
        if buf.tell() <= max_bytes:
            with open(local_path, "wb") as f: f.write(buf.getvalue())
            print(f"Compressed {orig/1024:.0f} KB → {buf.tell()/1024:.0f} KB (q={q}).")
            return local_path
    w, h = img.size
    scale = 0.9
    while scale > 0.3:
        r = img.resize((max(1,int(w*scale)), max(1,int(h*scale))), Image.LANCZOS)
        buf = io.BytesIO()
        r.save(buf, format="JPEG", quality=70, optimize=True)
        if buf.tell() <= max_bytes:
            with open(local_path, "wb") as f: f.write(buf.getvalue())
            print(f"Resized+compressed → {buf.tell()/1024:.0f} KB.")
            return local_path
        scale -= 0.1
    with open(local_path, "wb") as f: f.write(buf.getvalue())
    print(f"Warning: best-effort compression = {buf.tell()/1024:.0f} KB.")
    return local_path


def move_file(file_id, restore_name=None):
    creds   = get_creds()
    service = build("drive", "v3", credentials=creds)
    cfg     = _cfg()
    body    = {"name": restore_name} if restore_name else {}
    service.files().update(
        fileId=file_id,
        addParents=cfg["processed_folder_id"],
        removeParents=cfg["upload_folder_id"],
        body=body,
    ).execute()
    print("Moved to processed folder.")


def release_claim(file_id, original_name):
    try:
        service = build("drive", "v3", credentials=get_creds())
        service.files().update(fileId=file_id, body={"name": original_name}).execute()
        print(f"Released claim on '{original_name}'.")
    except Exception as exc:
        print(f"Warning: could not release claim: {exc}")


# ═══════════════════════════════════════════════════════════════════════════
#  POST BUILDING
# ═══════════════════════════════════════════════════════════════════════════

MAX_POST_GRAPHEMES = 300

def build_post_from_caption(caption, tags, add_link):
    cfg  = _cfg()
    text = replace_mentions(caption) if caption else ""

    def _assemble(caption_text):
        tb = TextBuilder()
        if add_link:
            m = _URL_RE.search(caption_text)
            if m:
                before = caption_text[:m.start()].rstrip()
                after  = _URL_RE.sub("", caption_text[m.end():]).strip()
                if before:
                    tb.text(before + " ")
                tb.link(cfg["link_display_text"], cfg["link_url"])
                if after:
                    tb.text(" " + after)
            else:
                if caption_text:
                    tb.text(caption_text)
                    tb.text("\n\n")
                tb.link(cfg["link_display_text"], cfg["link_url"])
        else:
            text_no_url = _URL_RE.sub("", caption_text).strip()
            if text_no_url:
                tb.text(text_no_url)

        if tags:
            tb.text("\n\n")
            for i, tag in enumerate(tags):
                tb.tag(f"#{tag}", tag)
                if i < len(tags) - 1:
                    tb.text(" ")
        return tb

    tb    = _assemble(text)
    plain = tb.build_text()

    if len(plain) > MAX_POST_GRAPHEMES:
        lo, hi, best_text = 0, len(text), ""
        while lo <= hi:
            mid   = (lo + hi) // 2
            trial = text[:mid].rstrip()
            if mid < len(text):
                trial += "…"
            if len(_assemble(trial).build_text()) <= MAX_POST_GRAPHEMES:
                best_text = trial
                lo = mid + 1
            else:
                hi = mid - 1
        print(f"Caption too long for post limit ({len(plain)} > {MAX_POST_GRAPHEMES}); "
              f"trimmed caption to fit.")
        tb = _assemble(best_text)

    return tb


def build_link_post_text(caption, tags, card_url):
    """Builds the text for a link_post. If a card_url is present, the raw
    URL is stripped out of the visible text (the card itself carries the
    link) unless LINKPOST_INLINE_LINK_WITH_CARD is enabled in Settings."""
    cfg = _cfg()
    keep_inline_link = (not card_url) or cfg["linkpost_inline_link_with_card"]
    return build_post_from_caption(caption, tags, add_link=keep_inline_link)


def post_to_bluesky(client, media_name, local_path, kind, caption, tags, add_link):
    tb = build_post_from_caption(caption, tags, add_link)
    if kind == "video":
        with open(local_path, "rb") as f:
            client.send_video(text=tb, video=f.read(), video_alt=media_name)
    else:
        with open(local_path, "rb") as f:
            client.send_image(text=tb, image=f.read(), image_alt=media_name)

    preview = replace_mentions(caption or "")
    if add_link:
        m = _URL_RE.search(preview)
        if m:
            preview = (preview[:m.start()].rstrip()
                       + f" [{_cfg()['link_display_text']}]"
                       + _URL_RE.sub("", preview[m.end():]).strip())
        else:
            preview = (preview + f" [{_cfg()['link_display_text']}]").strip()
    else:
        preview = _URL_RE.sub("", preview).strip()
    print(f"✓ Posted {kind}: {preview!r} (link={'yes' if add_link else 'no'})")
    if tags:
        print(f"  Tags: {' '.join('#'+t for t in tags)}")


def post_link_post_to_bluesky(client, caption, tags, card_url=None):
    """Posts a text-only post. If card_url is set, a real social-card
    embed (title/description/thumbnail scraped from that URL) is attached —
    this is what actually produces a rich preview on Bluesky, unlike a plain
    inline hyperlink. Falls back to the account's default inline link (no
    card) if card_url is blank."""
    tb = build_link_post_text(caption, tags, card_url)

    embed = None
    if card_url:
        try:
            embed = build_external_embed(client, card_url)
        except Exception as exc:
            print(f"Warning: could not build link card for {card_url!r}: {exc}. "
                  f"Posting without a card.")

    if embed is not None:
        client.send_post(text=tb, embed=embed)
    else:
        client.send_post(text=tb)

    preview = replace_mentions(caption or "")
    if card_url:
        preview = _URL_RE.sub("", preview).strip()
        preview = f"{preview}  [card: {card_url}]".strip()
    else:
        m = _URL_RE.search(preview)
        if m:
            preview = (preview[:m.start()].rstrip()
                       + f" [{_cfg()['link_display_text']}]"
                       + _URL_RE.sub("", preview[m.end():]).strip())
        else:
            preview = (preview + f" [{_cfg()['link_display_text']}]").strip()

    print(f"✓ Posted link_post: {preview!r}")
    if tags:
        print(f"  Tags: {' '.join('#'+t for t in tags)}")


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN CYCLE
# ═══════════════════════════════════════════════════════════════════════════

def run_once():
    cfg = refresh_account_config()

    if not try_acquire_account_lock():
        raise AccountLockedElsewhereError(
            f"Account row {ACCOUNT_ROW} is locked by another repo right now."
        )

    handle = cfg["handle"]

    print_target_account(handle)
    client = Client()
    try:
        client.login(handle, cfg["app_pw"])
    except Exception as exc:
        err = str(exc)
        if "AccountTakedown" in err or "AccountSuspended" in err:
            raise AccountTakenDownError(f"Account {handle} taken down/suspended.") from exc
        if "AuthenticationRequired" in err or "Invalid identifier or password" in err:
            raise AccountTakenDownError(
                f"Auth failed for {handle} — check BSKY_HANDLE / BSKY_APP_PW in sheet row {ACCOUNT_ROW}."
            ) from exc
        raise

    if cfg["enable_report"]:
        run_report(client, handle, cfg)

    preferred = choose_media_kind()

    # ── link_post branch — handled entirely separately from the Drive
    # image/video flow, since a link post has no media file to fetch. ──
    if preferred == "link_post":
        entries = load_link_post_plan()
        if not entries:
            fallback = random.choice(["image", "video"])
            print(f"link_post chosen but '{get_link_post_tab_name()}' tab has no "
                  f"unposted rows (or doesn't exist yet) — falling back to {fallback}.")
            preferred = fallback
        else:
            entry = random.choice(entries)
            tags  = get_account_hashtags() if cfg["hashtags_enabled_linkpost"] else []
            try:
                post_link_post_to_bluesky(client, entry["caption"], tags, card_url=entry.get("url"))
            except Exception as exc:
                err = str(exc)
                if "AccountTakedown" in err or "AccountSuspended" in err:
                    raise AccountTakenDownError(f"Account {handle} taken down mid-cycle.") from exc
                raise
            mark_link_post_posted(entry["row"])
            return  # link post cycle complete — nothing else to do

    # ── Existing Drive-backed image/video flow ──
    plan = load_post_plan()
    if not plan:
        raise NoMediaFoundError("Post-plan sheet has no usable rows.")

    fallback = "video" if preferred == "image" else "image"

    file, path, kind, caption, row_num = fetch_media_matching_plan(preferred, plan)
    if not file:
        print(f"No {preferred} matched; trying {fallback}.")
        file, path, kind, caption, row_num = fetch_media_matching_plan(fallback, plan)

    if not file:
        raise NoMediaFoundError("No unposted Drive file matching the post-plan sheet.")

    original_name = file["original_name"]

    try:
        if kind == "image":
            path = compress_image_under_limit(path)

        cfg = _cfg()
        hashtags_on = cfg["hashtags_enabled_image"] if kind == "image" else cfg["hashtags_enabled_video"]
        tags = get_account_hashtags() if hashtags_on else []
        add_link = should_add_link(kind)

        post_to_bluesky(client, original_name, path, kind, caption, tags, add_link)

    except Exception as exc:
        err = str(exc)
        if "AccountTakedown" in err or "AccountSuspended" in err:
            release_claim(file["id"], original_name)
            raise AccountTakenDownError(f"Account {handle} taken down mid-cycle.") from exc
        release_claim(file["id"], original_name)
        print(f"Post failed — claim released, file stays in upload folder.")
        raise

    mark_posted(original_name, row_num)
    try:
        move_file(file["id"], restore_name=original_name)
    except Exception as exc:
        print(f"Warning: move_file failed: {exc}. File may still be in upload folder — remove manually.")
    try:
        os.remove(path)
    except OSError:
        pass


def main():
    global ACCOUNT_ROW
    try:
        ACCOUNT_ROW = resolve_account_row()
        load_account_config()
    except Exception as exc:
        print(f"\n{'='*60}\nFATAL: {exc}\n{'='*60}\n")
        sys.exit(1)

    print_config_summary()
    print(f"Starting loop. Loop interval is read from the Settings tab "
          f"(LOOP_INTERVAL_SECONDS) and re-checked at the start of every cycle "
          f"— edit it in Google Sheets any time, no redeploy needed.")

    while True:
        cycle_start = time.time()
        try:
            run_once()
        except AccountLockedElsewhereError as exc:
            print(f"\n{'='*60}\n{exc}\nSkipping — schedule keeps running.\n{'='*60}\n")
            sys.exit(0)
        except NoMediaFoundError as exc:
            print(f"\n{'='*60}\nNO MEDIA: {exc}\nStopping — schedule keeps running.\n{'='*60}\n")
            sys.exit(0)
        except AccountTakenDownError as exc:
            handle  = (_account_config or {}).get("handle", "unknown")
            err_str = str(exc)
            reason  = ("🔑 AUTH FAILED — check handle/app-password in sheet"
                       if "Auth failed" in err_str or "app password" in err_str
                       else "⛔ ACCOUNT TAKEN DOWN / BANNED")
            print(f"\n{'='*60}\n{err_str}\n→ {reason}\n{'='*60}\n")
            log_account_problem(handle, status=reason)
            sys.exit(1)
        except Exception as exc:
            print(f"Error during cycle: {exc}")

        loop_interval = (_account_config or {}).get("loop_interval_seconds", DEFAULT_LOOP_INTERVAL_SECONDS)
        elapsed   = time.time() - cycle_start
        sleep_for = max(0, loop_interval - elapsed)
        print(f"Cycle done in {elapsed:.1f}s. Sleeping {sleep_for:.1f}s "
              f"(interval={loop_interval}s from Settings tab)…")
        time.sleep(sleep_for)


if __name__ == "__main__":
    main()
