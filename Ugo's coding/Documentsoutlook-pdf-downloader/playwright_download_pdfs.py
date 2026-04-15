import os
import re
import json
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

load_dotenv()

OUTLOOK_URL = os.getenv("OUTLOOK_URL", "https://outlook.office.com/mail/")
OUTLOOK_EMAIL = os.getenv("OUTLOOK_EMAIL", "")
OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD", "")
DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", os.getenv("DOWNLOAD_DIR".lower(), "downloads")))
MAX_EMAILS = int(os.getenv("MAX_EMAILS", "25"))

# We'll save login cookies here so you don't keep logging in every run
STORAGE_STATE_PATH = Path("storage_state.json")
LOG_PATH = Path("download_log.json")

PDF_RE = re.compile(r"\.pdf(\?|$)", re.IGNORECASE)


def load_log():
    if LOG_PATH.exists():
        return json.loads(LOG_PATH.read_text(encoding="utf-8"))
    return {"downloaded": []}


def save_log(log):
    LOG_PATH.write_text(json.dumps(log, indent=2), encoding="utf-8")


def looks_like_pdf(name_or_url: str) -> bool:
    return bool(PDF_RE.search(name_or_url or ""))


def ensure_download_dir():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)


def maybe_login(page):
   
    page.goto(OUTLOOK_URL, wait_until="domcontentloaded")

   
    try:
        page.wait_for_selector('div[role="main"]', timeout=8000)
        return
    except PWTimeoutError:
        pass

   
    try:
        # Email
        if page.locator('input[type="email"]').count() > 0:
            page.fill('input[type="email"]', OUTLOOK_EMAIL)
            page.click('input[type="submit"], button[type="submit"]')
        page.wait_for_timeout(1000)

        # Password
        if page.locator('input[type="password"]').count() > 0:
            page.fill('input[type="password"]', OUTLOOK_PASSWORD)
            page.click('input[type="submit"], button[type="submit"]')
        page.wait_for_timeout(1000)

        # "Stay signed in?" prompt sometimes appears
        if page.locator('input#idBtn_Back').count() > 0:
            # choose "No" to reduce surprises, but either is fine
            page.click('input#idBtn_Back')
        page.wait_for_timeout(1000)
    except Exception:
        pass

    # If MFA / extra prompts happen, you must complete them manually
    print("\nIf you see MFA / extra login steps in the browser, complete them now.")
    print("When your inbox loads, come back here and press ENTER.\n")
    input()

    # Wait for mailbox UI
    page.wait_for_selector('div[role="main"]', timeout=60000)


def open_latest_message(page, index: int):
    """
    Tries to click the message in the list by index.
    Outlook message list is dynamic; this uses a few common patterns.
    """
    # Common: message list items have role="option"
    options = page.locator('[role="option"]')
    if options.count() > index:
        options.nth(index).click()
        return True

    # Fallback: list rows sometimes have role="row"
    rows = page.locator('[role="row"]')
    if rows.count() > index:
        rows.nth(index).click()
        return True

    return False


def download_pdf_attachments_from_open_email(page, log):
    """
    Looks for attachment items in the reading pane and downloads PDFs by clicking.
    """
    downloaded_now = 0

    # Attachment chips/links often show filenames; we click any that end in .pdf
    attachment_candidates = page.locator('a:has-text(".pdf"), span:has-text(".pdf"), div:has-text(".pdf")')
    count = min(attachment_candidates.count(), 30)  # safety
    for i in range(count):
        text = (attachment_candidates.nth(i).inner_text() or "").strip()
        if not looks_like_pdf(text):
            continue

        key = f"att:{text}"
        if key in log["downloaded"]:
            continue

        try:
            with page.expect_download(timeout=15000) as dl_info:
                attachment_candidates.nth(i).click()
            download = dl_info.value
            suggested = download.suggested_filename or "attachment.pdf"
            if not looks_like_pdf(suggested):
                # Still save, but keep extension safe
                suggested = suggested + ".pdf"
            save_path = DOWNLOAD_DIR / suggested
            download.save_as(str(save_path))

            log["downloaded"].append(key)
            downloaded_now += 1
            print(f"✅ Downloaded attachment: {save_path.name}")
        except PWTimeoutError:
            # Sometimes clicking opens preview instead of direct download
            # We'll just skip quietly for now.
            pass
        except Exception:
            pass

    return downloaded_now


def download_pdf_links_in_body(page, log):
    """
    Finds links in the email body that point to PDFs and downloads them (if direct).
    """
    downloaded_now = 0

    links = page.locator('a[href]')
    total = min(links.count(), 200)  # safety
    for i in range(total):
        href = links.nth(i).get_attribute("href") or ""
        if not looks_like_pdf(href):
            continue

        key = f"url:{href}"
        if key in log["downloaded"]:
            continue

        try:
            # Open in a new tab and attempt download
            with page.context.expect_page() as new_page_info:
                links.nth(i).click(button="middle")  # open new tab
            newp = new_page_info.value
            newp.wait_for_timeout(1500)

            # Try a direct download from that page
            try:
                with newp.expect_download(timeout=10000) as dl_info:
                    newp.click("body")  # sometimes triggers download immediately
                download = dl_info.value
                suggested = download.suggested_filename or "linked.pdf"
                if not looks_like_pdf(suggested):
                    suggested += ".pdf"
                save_path = DOWNLOAD_DIR / suggested
                download.save_as(str(save_path))
                log["downloaded"].append(key)
                downloaded_now += 1
                print(f"✅ Downloaded link PDF: {save_path.name}")
            except PWTimeoutError:
                # Not a direct-download link or requires auth; skip for now
                pass
            finally:
                newp.close()

        except Exception:
            pass

    return downloaded_now


def main():
    ensure_download_dir()
    log = load_log()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        if STORAGE_STATE_PATH.exists():
            context = browser.new_context(
                accept_downloads=True,
                storage_state=str(STORAGE_STATE_PATH),
            )
        else:
            context = browser.new_context(accept_downloads=True)

        page = context.new_page()
        maybe_login(page)

        # Save cookies/session so next runs don’t require logging in again
        context.storage_state(path=str(STORAGE_STATE_PATH))

        print(f"\nScanning up to {MAX_EMAILS} emails…\n")

        total_downloads = 0
        opened = 0

        for idx in range(MAX_EMAILS):
            ok = open_latest_message(page, idx)
            if not ok:
                print("Could not find more messages in the list.")
                break

            opened += 1
            page.wait_for_timeout(1200)  # let reading pane load

            # Download PDF attachments + body links
            total_downloads += download_pdf_attachments_from_open_email(page, log)
            total_downloads += download_pdf_links_in_body(page, log)

            save_log(log)

        print(f"\nDone. Opened {opened} emails. Downloaded {total_downloads} PDFs.")
        print(f"Saved to: {DOWNLOAD_DIR.resolve()}")
        browser.close()


if __name__ == "__main__":
    main()