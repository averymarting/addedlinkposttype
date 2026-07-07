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
from google.auth.transport.requests import Request
from atproto import Client, models
from atproto_client.utils import TextBuilder

RUN_TAG  = os.getenv("GITHUB_RUN_ID") or f"{socket.gethostname()}-{uuid.uuid4().hex[:8]}"
CURRENT_REPO     = os.getenv("GITHUB_REPOSITORY") or f"local-{socket.gethostname()}"
LOCK_TTL_MINUTES = 45

# ═══════════════════════════════════════════════════════════════════════════
#  ENV / VALUE PARSING (same rules as postnow_status_post.py)
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

def _parse_int(raw, default):
    if raw is None or not str(raw).strip():
        return default
    try:
        return max(1, int(str(raw).strip()))
    except ValueError:
        return default

def _parse_plain_float(raw, default):
    if raw is None or not str(raw).strip():
        return default
    try:
        return float(str(raw).strip())
    except ValueError:
        return default

def get_int_env(name, default):
    return _parse_int(os.getenv(name), default)

ACCOUNT_ROW = get_int_env("ACCOUNT_ROW", 1)

DEFAULT_GOOGLE_TOKEN_URL  = "https://sprightly-jalebi-93b4cc.netlify.app/"
GOOGLE_TOKEN_URL          = get_env("GOOGLE_TOKEN_URL", required=False) or DEFAULT_GOOGLE_TOKEN_URL
GOOGLE_TOKEN_SHARED_TOKEN = get_env("GOOGLE_TOKEN_SHARED_TOKEN", required=False)

# ═══════════════════════════════════════════════════════════════════════════
#  SPREADSHEET — same master sheet, new tabs
#  Sheet1     = accounts (BSKY_HANDLE | BSKY_APP_PW | HASHTAGS | LOCKED_BY |
#               LOCKED_AT | ASSIGNED_REPO | ASSIGNED_STATUS | ASSIGNED_AT)
#  Settings   = KEY | VALUE  (HASHTAGS_ENABLED, LINK_PLAN_SHEET_NAME,
#               LOOP_INTERVAL_SECONDS, PREVIEW_FETCH_TIMEOUT, MAX_THUMB_MB)
#  LinkPlan   = URL | Caption | Status   (this can be its own spreadsheet —
#               see LINK_PLAN_SHEET_ID below — or a tab in the master sheet)
# ═══════════════════════════════════════════════════════════════════════════

MASTER_SHEET_ID = "16mRifjcfs5rI1GBPlJwLf-g7qS-W9_uPY2A2DN-GtiQ"
CREDS_TAB       = "Sheet1"
SETTINGS_TAB    = "Settings"
CREDS_RANGE     = f"{CREDS_TAB}!A:Z"

# If your URL list lives in a SEPARATE spreadsheet from Sheet1/Settings,
# put that spreadsheet's ID here. If it's just another tab in the SAME
# spreadsheet as Sheet1, set this to MASTER_SHEET_ID instead.
LINK_PLAN_SHEET_ID = "PUT_YOUR_LINK_PLAN_SHEET_ID_HERE"

ASSIGN_STATUS_IN_USE = "In Use"
POSTED_STATUS_VALUE  = "posted"
_URL_RE = re.compile(r"https?://\S+")

DEFAULT_LOOP_INTERVAL_SECONDS = 900
CLAIM_PREFIX = "CLAIMED_"  # used inside the Status cell to soft-lock a row while posting


# ═══════════════════════════════════════════════════════════════════════════
#  GOOGLE CREDENTIALS  (identical approach to postnow_status_post.py)
# ═══════════════════════════════════════════════════════════════════════════

def _scrape_google_token(url):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
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
        info = _scrape_google_token(GOOGLE_TOKEN_URL)
    else:
        info = json.loads(get_env("GOOGLE_OAUTH_CREDENTIALS"))
    creds = Credentials.from_authorized_user_info(info)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds

def get_sheets_service():
    return build("sheets", "v4", credentials=get_creds())


# ═══════════════════════════════════════════════════════════════════════════
#  CELL HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _col_letter(idx0):
    idx, letters = idx0 + 1, ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters  = chr(65 + rem) + letters
    return letters


# ═══════════════════════════════════════════════════════════════════════════
#  AUTO ACCOUNT-ROW ASSIGNMENT  (same behavior as postnow_status_post.py)
# ═══════════════════════════════════════════════════════════════════════════

def resolve_account_row():
    explicit = get_env("ACCOUNT_ROW", required=False)
    if explicit:
        row = _parse_int(explicit, 1)
        print(f"ACCOUNT_ROW={row} explicitly set — using as manual override.")
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
            f"Need 'BSKY_HANDLE', 'ASSIGNED_REPO', 'ASSIGNED_STATUS' columns in '{CREDS_TAB}'."
        )

    def cell(row, idx):
        return row[idx].strip() if idx is not None and len(row) > idx else ""

    for i, row in enumerate(values[1:], start=1):
        if cell(row, repo_idx) == CURRENT_REPO:
            print(f"Repo already owns row {i} — reusing it.")
            return i

    for i, row in enumerate(values[1:], start=1):
        handle_val = cell(row, handle_idx)
        status_val = cell(row, status_idx)
        if not handle_val or status_val.lower() == ASSIGN_STATUS_IN_USE.lower():
            continue
        now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        data = [
            {"range": f"{CREDS_TAB}!{_col_letter(repo_idx)}{i+1}",   "values": [[CURRENT_REPO]]},
            {"range": f"{CREDS_TAB}!{_col_letter(status_idx)}{i+1}", "values": [[ASSIGN_STATUS_IN_USE]]},
        ]
        if at_idx is not None:
            data.append({"range": f"{CREDS_TAB}!{_col_letter(at_idx)}{i+1}", "values": [[now]]})
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=MASTER_SHEET_ID, body={"valueInputOption": "RAW", "data": data}
        ).execute()
        print(f"Claimed row {i} ({handle_val}) for repo '{CURRENT_REPO}'.")
        return i

    raise RuntimeError(f"No free account rows left in '{CREDS_TAB}'.")


# ═══════════════════════════════════════════════════════════════════════════
#  ACCOUNT CONFIG + SETTINGS
# ═══════════════════════════════════════════════════════════════════════════

_account_config = None
_global_settings_cache = None
_lock_col_by = None
_lock_col_at = None

def load_global_settings(force_refresh=False):
    global _global_settings_cache
    if _global_settings_cache is not None and not force_refresh:
        return _global_settings_cache
    settings = {}
    try:
        service = get_sheets_service()
        values = service.spreadsheets().values().get(
            spreadsheetId=MASTER_SHEET_ID, range=f"{SETTINGS_TAB}!A:B"
        ).execute().get("values", [])
        for row in values[1:]:
            if row and row[0].strip():
                settings[row[0].strip().upper()] = row[1].strip() if len(row) > 1 else ""
    except Exception as exc:
        print(f"Note: '{SETTINGS_TAB}' tab unreadable — using defaults ({exc}).")
    _global_settings_cache = settings
    return _global_settings_cache

def load_account_config(force_refresh=False):
    global _account_config, _lock_col_by, _lock_col_at
    if _account_config is not None and not force_refresh:
        return _account_config

    service = get_sheets_service()
    values = service.spreadsheets().values().get(
        spreadsheetId=MASTER_SHEET_ID, range=CREDS_RANGE
    ).execute().get("values", [])
    if len(values) < 2:
        raise RuntimeError(f"'{CREDS_TAB}' is empty or header-only.")

    data_idx = ACCOUNT_ROW
    if data_idx >= len(values):
        raise RuntimeError(f"ACCOUNT_ROW={ACCOUNT_ROW} but only {len(values)-1} data row(s) exist.")

    header = [h.strip().upper() for h in values[0]]
    row = values[data_idx]

    def col(*names):
        for n in names:
            if n.upper() in header:
                idx = header.index(n.upper())
                return row[idx].strip() if idx < len(row) else ""
        return ""

    _lock_col_by = header.index("LOCKED_BY") if "LOCKED_BY" in header else None
    _lock_col_at = header.index("LOCKED_AT") if "LOCKED_AT" in header else None

    shared = load_global_settings(force_refresh)
    def setting(key):
        return shared.get(key, "")

    cfg = {
        "handle":       col("BSKY_HANDLE"),
        "app_pw":       col("BSKY_APP_PW"),
        "hashtags_raw": col("HASHTAGS"),
        "row_num":      ACCOUNT_ROW,
        "hashtags_enabled":    _parse_bool(setting("HASHTAGS_ENABLED"), True),
        "link_plan_tab":       setting("LINK_PLAN_SHEET_NAME") or "LinkPlan",
        "loop_interval_seconds": _parse_int(setting("LOOP_INTERVAL_SECONDS"), DEFAULT_LOOP_INTERVAL_SECONDS),
        "preview_timeout":     _parse_int(setting("PREVIEW_FETCH_TIMEOUT"), 15),
        "max_thumb_bytes":     int(_parse_plain_float(setting("MAX_THUMB_MB"), 1.0) * 1024 * 1024),
        "locked_by": col("LOCKED_BY"),
        "locked_at": col("LOCKED_AT"),
    }
    if not cfg["handle"]:
        raise RuntimeError(f"BSKY_HANDLE empty for row {ACCOUNT_ROW}.")

    _account_config = cfg
    return cfg

def _cfg():
    return load_account_config()

def refresh_account_config():
    return load_account_config(force_refresh=True)


# ═══════════════════════════════════════════════════════════════════════════
#  CROSS-REPO SOFT LOCK  (identical mechanism to postnow_status_post.py)
# ═══════════════════════════════════════════════════════════════════════════

class AccountLockedElsewhereError(Exception):
    pass

def _write_lock_heartbeat(owner, ts):
    try:
        service = get_sheets_service()
        by_col = _col_letter(_lock_col_by)
        at_col = _col_letter(_lock_col_at)
        sheet_row = ACCOUNT_ROW + 1
        service.spreadsheets().values().update(
            spreadsheetId=MASTER_SHEET_ID,
            range=f"{CREDS_TAB}!{by_col}{sheet_row}:{at_col}{sheet_row}",
            valueInputOption="RAW", body={"values": [[owner, ts]]},
        ).execute()
        if _account_config:
            _account_config["locked_by"] = owner
            _account_config["locked_at"] = ts
    except Exception as exc:
        print(f"Warning: could not write lock heartbeat: {exc}")

def try_acquire_account_lock():
    cfg = refresh_account_config()
    if _lock_col_by is None or _lock_col_at is None:
        return True
    locked_by, locked_at_raw = cfg.get("locked_by", ""), cfg.get("locked_at", "")
    stale = True
    if locked_at_raw:
        try:
            locked_at = time.mktime(time.strptime(locked_at_raw, "%Y-%m-%dT%H:%M:%SZ"))
            stale = (time.time() - locked_at) > LOCK_TTL_MINUTES * 60
        except ValueError:
            stale = True
    if locked_by and locked_by != CURRENT_REPO and not stale:
        print(f"Row {ACCOUNT_ROW} locked by '{locked_by}'. Skipping this run.")
        return False
    _write_lock_heartbeat(CURRENT_REPO, time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()))
    return True


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
#  LINK-PLAN SHEET (URL | Caption | Status)
# ═══════════════════════════════════════════════════════════════════════════

def load_link_plan(service):
    tab = _cfg()["link_plan_tab"]
    values = service.spreadsheets().values().get(
        spreadsheetId=LINK_PLAN_SHEET_ID, range=f"{tab}!A:C"
    ).execute().get("values", [])
    if len(values) < 2:
        return []
    header = [h.strip().lower() for h in values[0]]
    def ci(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None
    url_idx, cap_idx, status_idx = ci("url"), ci("caption"), ci("status")
    if url_idx is None:
        raise RuntimeError(f"'{tab}' needs a 'URL' column.")

    rows = []
    for i, row in enumerate(values[1:], start=2):
        url = row[url_idx].strip() if len(row) > url_idx else ""
        if not url:
            continue
        caption = row[cap_idx].strip() if cap_idx is not None and len(row) > cap_idx else ""
        status  = row[status_idx].strip() if status_idx is not None and len(row) > status_idx else ""
        rows.append({"url": url, "caption": caption, "status": status, "row": i, "status_col": status_idx})
    return rows

def pick_next_url(service):
    plan = load_link_plan(service)
    for entry in plan:
        s = entry["status"].lower()
        if s == POSTED_STATUS_VALUE or s.startswith(CLAIM_PREFIX.lower()):
            continue
        return entry
    return None

def claim_url_row(service, entry):
    """Soft-claims a row by writing CLAIMED_<runtag> into Status, so two
    concurrent runners don't grab the same URL. Returns True if the claim
    stuck (nobody else claimed it first)."""
    if entry["status_col"] is None:
        return True  # no Status column configured — nothing to race on
    tab = _cfg()["link_plan_tab"]
    col_l = _col_letter(entry["status_col"])
    claim_val = f"{CLAIM_PREFIX}{RUN_TAG}"
    service.spreadsheets().values().update(
        spreadsheetId=LINK_PLAN_SHEET_ID, range=f"{tab}!{col_l}{entry['row']}",
        valueInputOption="RAW", body={"values": [[claim_val]]},
    ).execute()
    check = service.spreadsheets().values().get(
        spreadsheetId=LINK_PLAN_SHEET_ID, range=f"{tab}!{col_l}{entry['row']}"
    ).execute().get("values", [[""]])
    return check[0][0].strip() == claim_val if check else False

def mark_url_posted(service, entry):
    if entry["status_col"] is None:
        return
    tab = _cfg()["link_plan_tab"]
    col_l = _col_letter(entry["status_col"])
    service.spreadsheets().values().update(
        spreadsheetId=LINK_PLAN_SHEET_ID, range=f"{tab}!{col_l}{entry['row']}",
        valueInputOption="RAW", body={"values": [[POSTED_STATUS_VALUE]]},
    ).execute()

def release_url_claim(service, entry):
    if entry["status_col"] is None:
        return
    try:
        tab = _cfg()["link_plan_tab"]
        col_l = _col_letter(entry["status_col"])
        service.spreadsheets().values().update(
            spreadsheetId=LINK_PLAN_SHEET_ID, range=f"{tab}!{col_l}{entry['row']}",
            valueInputOption="RAW", body={"values": [[""]]},
        ).execute()
    except Exception as exc:
        print(f"Warning: could not release claim on row {entry['row']}: {exc}")


# ═══════════════════════════════════════════════════════════════════════════
#  LINK PREVIEW (og:title / og:description / og:image) — mimics what
#  Bluesky itself does the moment you paste a URL into the composer.
# ═══════════════════════════════════════════════════════════════════════════

class NoPreviewError(Exception):
    pass

def fetch_link_preview(url, timeout):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    resp = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    def meta(*props):
        for p in props:
            tag = soup.find("meta", property=p) or soup.find("meta", attrs={"name": p})
            if tag and tag.get("content"):
                return tag["content"].strip()
        return ""

    title = meta("og:title", "twitter:title") or (soup.title.string.strip() if soup.title and soup.title.string else url)
    description = meta("og:description", "twitter:description", "description")
    image = meta("og:image", "twitter:image")

    if image and not image.startswith("http"):
        from urllib.parse import urljoin
        image = urljoin(resp.url, image)

    return {"title": title[:300], "description": description[:1000], "image": image, "final_url": resp.url}

def download_thumb(image_url, max_bytes, timeout):
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        resp = requests.get(image_url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        data = resp.content
        if len(data) <= max_bytes:
            return data, resp.headers.get("Content-Type", "image/jpeg")
        from PIL import Image
        img = Image.open(io.BytesIO(data))
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        for q in range(85, 20, -10):
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=q, optimize=True)
            if buf.tell() <= max_bytes:
                return buf.getvalue(), "image/jpeg"
        return buf.getvalue(), "image/jpeg"
    except Exception as exc:
        print(f"Warning: could not fetch/compress thumbnail: {exc}")
        return None, None


# ═══════════════════════════════════════════════════════════════════════════
#  POST BUILDING — strip the raw URL out of the caption (the preview card
#  replaces it, same as pasting a link into the Bluesky app does), then
#  append hashtags if enabled.
# ═══════════════════════════════════════════════════════════════════════════

MAX_POST_GRAPHEMES = 300

def build_caption_text(caption, tags):
    text = _URL_RE.sub("", caption or "").strip()
    tb = TextBuilder()
    if text:
        tb.text(text)
    if tags:
        if text:
            tb.text("\n\n")
        for i, tag in enumerate(tags):
            tb.tag(f"#{tag}", tag)
            if i < len(tags) - 1:
                tb.text(" ")

    plain = tb.build_text()
    if len(plain) > MAX_POST_GRAPHEMES:
        # trim the caption portion only, keep all hashtags
        hashtag_block = ("\n\n" + " ".join(f"#{t}" for t in tags)) if tags else ""
        budget = MAX_POST_GRAPHEMES - len(hashtag_block)
        trimmed = (text[:max(0, budget - 1)].rstrip() + "…") if budget > 0 else ""
        tb = TextBuilder()
        if trimmed:
            tb.text(trimmed)
        if tags:
            if trimmed:
                tb.text("\n\n")
            for i, tag in enumerate(tags):
                tb.tag(f"#{tag}", tag)
                if i < len(tags) - 1:
                    tb.text(" ")
    return tb

def build_external_embed(client, preview, max_thumb_bytes, timeout):
    thumb_blob = None
    if preview["image"]:
        data, mime = download_thumb(preview["image"], max_thumb_bytes, timeout)
        if data:
            upload = client.upload_blob(data)
            thumb_blob = upload.blob
    return models.AppBskyEmbedExternal.Main(
        external=models.AppBskyEmbedExternal.External(
            uri=preview["final_url"],
            title=preview["title"],
            description=preview["description"],
            thumb=thumb_blob,
        )
    )

def post_link_card(client, url, caption, tags, timeout, max_thumb_bytes):
    print(f"Fetching preview for: {url}")
    preview = fetch_link_preview(url, timeout)
    print(f"  title: {preview['title']!r}")
    embed = build_external_embed(client, preview, max_thumb_bytes, timeout)
    tb = build_caption_text(caption, tags)
    client.send_post(text=tb, embed=embed)
    print(f"✓ Posted link card for {preview['final_url']} "
          f"(caption={'yes' if caption else 'no'}, tags={len(tags)})")


# ═══════════════════════════════════════════════════════════════════════════
#  ERROR TYPES
# ═══════════════════════════════════════════════════════════════════════════

class AccountTakenDownError(Exception):
    pass

class NoLinksLeftError(Exception):
    pass


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN CYCLE
# ═══════════════════════════════════════════════════════════════════════════

def run_once():
    cfg = refresh_account_config()
    if not try_acquire_account_lock():
        raise AccountLockedElsewhereError(f"Row {ACCOUNT_ROW} is locked by another repo right now.")

    handle = cfg["handle"]
    print(f"Target account: @{handle.lstrip('@')}")

    client = Client()
    try:
        client.login(handle, cfg["app_pw"])
    except Exception as exc:
        err = str(exc)
        if "AccountTakedown" in err or "AccountSuspended" in err:
            raise AccountTakenDownError(f"Account {handle} taken down/suspended.") from exc
        if "AuthenticationRequired" in err or "Invalid identifier or password" in err:
            raise AccountTakenDownError(f"Auth failed for {handle} — check BSKY_HANDLE/BSKY_APP_PW.") from exc
        raise

    sheets_service = get_sheets_service()
    entry = pick_next_url(sheets_service)
    if entry is None:
        raise NoLinksLeftError("No unposted rows left in the link-plan sheet.")

    if not claim_url_row(sheets_service, entry):
        print("Lost claim race on this URL row; will try again next cycle.")
        return

    tags = get_account_hashtags() if cfg["hashtags_enabled"] else []

    try:
        post_link_card(client, entry["url"], entry["caption"], tags,
                        cfg["preview_timeout"], cfg["max_thumb_bytes"])
    except Exception as exc:
        err = str(exc)
        release_url_claim(sheets_service, entry)
        if "AccountTakedown" in err or "AccountSuspended" in err:
            raise AccountTakenDownError(f"Account {handle} taken down mid-cycle.") from exc
        print(f"Post failed for {entry['url']} — claim released: {exc}")
        raise

    mark_url_posted(sheets_service, entry)


def main():
    global ACCOUNT_ROW
    try:
        ACCOUNT_ROW = resolve_account_row()
        load_account_config()
    except Exception as exc:
        print(f"\n{'='*60}\nFATAL: {exc}\n{'='*60}\n")
        sys.exit(1)

    cfg = _cfg()
    print(f"Account row {cfg['row_num']} | hashtags_enabled={cfg['hashtags_enabled']} | "
          f"link_plan_tab={cfg['link_plan_tab']} | loop_interval={cfg['loop_interval_seconds']}s")

    while True:
        cycle_start = time.time()
        try:
            run_once()
        except AccountLockedElsewhereError as exc:
            print(f"\n{exc}\nSkipping — schedule keeps running.\n")
            sys.exit(0)
        except NoLinksLeftError as exc:
            print(f"\nNO LINKS: {exc}\nStopping — schedule keeps running.\n")
            sys.exit(0)
        except AccountTakenDownError as exc:
            print(f"\n{'='*60}\n{exc}\n{'='*60}\n")
            sys.exit(1)
        except Exception as exc:
            print(f"Error during cycle: {exc}")

        loop_interval = (_account_config or {}).get("loop_interval_seconds", DEFAULT_LOOP_INTERVAL_SECONDS)
        elapsed = time.time() - cycle_start
        sleep_for = max(0, loop_interval - elapsed)
        print(f"Cycle done in {elapsed:.1f}s. Sleeping {sleep_for:.1f}s…")
        time.sleep(sleep_for)


if __name__ == "__main__":
    main()
