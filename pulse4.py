import streamlit as st
import time
import os
import glob
import zipfile
import tempfile
import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

st.set_page_config(page_title="Sequential Ticket Automation", layout="wide")

TICKET_TYPE_LABELS = {
    "Pulse Tickets": {"Yes": "Normal Ticket", "No": "GXP Standard Ticket"},
    "Other Pulse Tickets": {"Yes": "Standard Ticket", "No": "GXP Normal Ticket"},
}

LOG_FILE_PATH = os.path.join(os.getcwd(), "automation_log.txt")

defaults = {
    "step": 0,
    "sso_id": "",
    "platform": "Pulse Tickets",
    "ticket_type_key": "Yes",
    "tickets": [],
    "logs": [],
    "download_dir": None,
    "headless": False,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

_log_placeholder = None
_image_placeholder = None


def log(msg: str):
    timestamped = f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}"
    st.session_state.logs.append(msg)
    print(msg)
    try:
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            f.write(timestamped + "\n")
    except Exception as e:
        print(f"Could not write to log file: {e}")
    if _log_placeholder is not None:
        _log_placeholder.text("\n".join(st.session_state.logs[-400:]))


def show_screenshot(driver, caption: str = ""):
    if _image_placeholder is None:
        return
    try:
        png_bytes = driver.get_screenshot_as_png()
        _image_placeholder.image(png_bytes, caption=caption, use_container_width=True)
    except Exception as e:
        log(f"  (could not capture screenshot: {e})")


def poll_until(driver, ticket, description, check_fn, timeout=60, interval=2):
    start = time.time()
    while time.time() - start < timeout:
        show_screenshot(driver, f"[{ticket}] {description}...")
        try:
            result = check_fn()
            if result:
                return result
        except StaleElementReferenceException:
            pass
        except Exception:
            pass
        time.sleep(interval)
    return None


# ==========================================================
# Shadow-DOM-aware element finders
# ==========================================================
DEEP_QS_JS = """
function deepQuerySelectorAll(selector, root) {
    root = root || document;
    let results = [];
    try { results = Array.from(root.querySelectorAll(selector)); } catch (e) {}
    const all = root.querySelectorAll('*');
    for (const el of all) {
        if (el.shadowRoot) {
            results = results.concat(deepQuerySelectorAll(selector, el.shadowRoot));
        }
    }
    return results;
}
return deepQuerySelectorAll(arguments[0]);
"""

DEEP_TEXT_JS = """
function getDirectText(el) {
    let text = '';
    for (const node of el.childNodes) {
        if (node.nodeType === Node.TEXT_NODE) { text += node.textContent; }
    }
    return text.trim();
}
function isVisible(el) {
    return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
}
function deepFindByText(text, root) {
    root = root || document;
    const all = root.querySelectorAll('*');
    for (const el of all) {
        if (el.shadowRoot) {
            const res = deepFindByText(text, el.shadowRoot);
            if (res) return res;
        }
        if (getDirectText(el) === text && isVisible(el)) {
            return el;
        }
    }
    return null;
}
return deepFindByText(arguments[0]);
"""

DEEP_ANY_VISIBLE_JS = """
function isVisible(el) {
    return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
}
function deepAnyVisible(selector, root) {
    root = root || document;
    let found = [];
    try { found = Array.from(root.querySelectorAll(selector)); } catch (e) {}
    for (const el of found) {
        if (isVisible(el)) return true;
    }
    const all = root.querySelectorAll('*');
    for (const el of all) {
        if (el.shadowRoot) {
            if (deepAnyVisible(selector, el.shadowRoot)) return true;
        }
    }
    return false;
}
return deepAnyVisible(arguments[0]);
"""


def deep_find_all(driver, css_selector):
    try:
        return driver.execute_script(DEEP_QS_JS, css_selector) or []
    except Exception:
        return []


def deep_find(driver, css_selector):
    els = deep_find_all(driver, css_selector)
    for el in els:
        try:
            if el.is_displayed():
                return el
        except Exception:
            continue
    return els[0] if els else None


def deep_find_by_text(driver, text):
    try:
        return driver.execute_script(DEEP_TEXT_JS, text)
    except Exception:
        return None


def deep_any_visible(driver, css_selector):
    try:
        return bool(driver.execute_script(DEEP_ANY_VISIBLE_JS, css_selector))
    except Exception:
        return False


def safe_click(driver, el):
    try:
        el.click()
        return
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", el)
        return
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(el).click().perform()
    except Exception as e:
        raise e


# ==========================================================
# Iframe-aware search wrapper
# IMPORTANT: iframes are found via deep_find_all() (shadow-DOM aware),
# NOT driver.find_elements(By.TAG_NAME, "iframe"), because ServiceNow's
# classic-UI iframe (gsft_main) is nested inside a shadow root and is
# completely invisible to plain find_elements().
# ==========================================================
def find_across_frames(driver, finder_fn, max_depth=3):
    el = finder_fn(driver)
    if el:
        return el

    driver.switch_to.default_content()
    el = finder_fn(driver)
    if el:
        return el

    def search_frames(depth):
        if depth > max_depth:
            return None
        frames = deep_find_all(driver, "iframe")
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
            except Exception:
                continue
            found = finder_fn(driver)
            if found:
                return found
            nested = search_frames(depth + 1)
            if nested:
                return nested
            driver.switch_to.parent_frame()
        return None

    result = search_frames(1)
    if result:
        return result

    driver.switch_to.default_content()
    return None


# ==========================================================
# Loading-spinner / step-completion helpers
# ==========================================================
SPINNER_SELECTORS = [
    "[class*='spinner' i]",
    "[class*='loading' i]",
    "[aria-busy='true']",
    ".now-loading-indicator",
]


def is_loading(driver):
    for sel in SPINNER_SELECTORS:
        if deep_any_visible(driver, sel):
            return True
    return False


def wait_for_spinner_to_clear(driver, ticket, context_label, timeout=30, interval=1):
    start = time.time()
    while time.time() - start < timeout:
        if not is_loading(driver):
            return True
        show_screenshot(driver, f"[{ticket}] {context_label}: waiting for loading to finish...")
        time.sleep(interval)
    return not is_loading(driver)


def wait_until_gone(driver, ticket, description, css_selector, timeout=30, interval=1):
    start = time.time()
    while time.time() - start < timeout:
        if not deep_any_visible(driver, css_selector):
            return True
        show_screenshot(driver, f"[{ticket}] {description}: waiting to close...")
        time.sleep(interval)
    return not deep_any_visible(driver, css_selector)


# ==========================================================
# Diagnostics — dumps detailed page state to the log (and log file)
# ==========================================================
def debug_dump_page_state(driver, ticket):
    try:
        driver.switch_to.default_content()
        log(f"  [DEBUG] === Diagnostics for ticket {ticket} ===")
        log(f"  [DEBUG] Current URL: {driver.current_url}")
        log(f"  [DEBUG] Page title: {driver.title}")
        log(f"  [DEBUG] Number of window handles: {len(driver.window_handles)}")

        iframes = deep_find_all(driver, "iframe")  # shadow-DOM aware
        log(f"  [DEBUG] Found {len(iframes)} iframe(s) on the page (incl. shadow DOM):")
        for i, f in enumerate(iframes):
            try:
                fid = f.get_attribute("id")
                fsrc = f.get_attribute("src")
                fname = f.get_attribute("name")
                log(f"    iframe[{i}] id={fid!r} name={fname!r} src={str(fsrc)[:150]!r}")
            except Exception as e:
                log(f"    iframe[{i}] (could not read attributes: {e})")

        candidate_selectors = [
            "button.additional-actions-context-menu-button",
            "button[aria-label='additional actions']",
            "button[aria-label*='action' i]",
            "button[title*='action' i]",
            "[aria-haspopup='true']",
            "button.icon-menu",
        ]

        def scan_frame(label):
            for sel in candidate_selectors:
                els = deep_find_all(driver, sel)
                if els:
                    details = []
                    for e in els:
                        try:
                            details.append(
                                f"(visible={e.is_displayed()}, text={e.text!r}, "
                                f"aria-label={e.get_attribute('aria-label')!r})"
                            )
                        except Exception as ex:
                            details.append(f"(error reading: {ex})")
                    log(f"    [{label}] {sel!r} -> {len(els)} match(es): {details}")
                else:
                    log(f"    [{label}] {sel!r} -> 0 matches")

        scan_frame("top-level")
        for i, f in enumerate(iframes):
            try:
                driver.switch_to.frame(f)
                scan_frame(f"iframe[{i}]")
                driver.switch_to.default_content()
            except Exception as e:
                log(f"    [iframe[{i}]] could not switch into frame: {e}")
                driver.switch_to.default_content()

        try:
            html = driver.page_source
            idx = html.lower().find("additional")
            if idx >= 0:
                snippet = html[max(0, idx - 200):idx + 300]
                log(f"  [DEBUG] HTML snippet around 'additional': ...{snippet}...")
            else:
                log("  [DEBUG] The word 'additional' was not found anywhere in top-level page_source.")
        except Exception as e:
            log(f"  [DEBUG] Could not dump page_source: {e}")

        log(f"  [DEBUG] === End diagnostics for {ticket} ===")

    except Exception as e:
        log(f"  [DEBUG] debug_dump_page_state itself failed: {e}")
    finally:
        driver.switch_to.default_content()


# ==========================================================
# Selenium driver setup
# ==========================================================
def build_driver(download_dir: str, headless: bool = False):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    }
    options.add_experimental_option("prefs", prefs)

    chromium_path = "/usr/bin/chromium"
    if os.path.exists(chromium_path):
        options.binary_location = chromium_path

    try:
        driver = webdriver.Chrome(options=options)
    except Exception as e:
        log(f"webdriver.Chrome() failed with Selenium Manager: {e}")
        log("Falling back to webdriver-manager...")
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

    return driver


# ==========================================================
# Login — visual + manual
# ==========================================================
def login(driver, sso_id: str):
    login_url = "https://pulse.service-now.com/now/nav/ui/home"
    expected_domain = "pulse.service-now.com"

    log("Opening login page in the browser window...")
    driver.get(login_url)
    show_screenshot(driver, "Login page loaded")

    try:
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
        )
        username_field.send_keys(sso_id)
        log("Auto-filled SSO ID. Please continue manually in the browser window "
            "(password, MFA, any prompts).")
        show_screenshot(driver, "SSO ID auto-filled — continue manually")
    except TimeoutException:
        log("Could not find username field automatically. "
            "Please log in fully manually in the browser window.")
        show_screenshot(driver, "Manual login required")

    log("Please complete SSO login (password, MFA/authenticator approval, etc.) "
        "yourself. Waiting for login to complete... (up to 5 minutes)")

    start = time.time()
    timeout = 300
    while time.time() - start < timeout:
        if expected_domain in driver.current_url and "login" not in driver.current_url.lower():
            break
        show_screenshot(driver, "Waiting for you to complete login...")
        time.sleep(3)
    else:
        raise TimeoutException("Login was not completed within 5 minutes.")

    log("Detected successful login! Resuming automation...")
    show_screenshot(driver, "Logged in successfully")


# ==========================================================
# Step 1: search box (shadow-DOM aware, top-level document only)
# ==========================================================
def open_search_and_type(driver, ticket):
    def find_input():
        return deep_find(driver, "#sncwsgs-typeahead-input")

    search_box = poll_until(driver, ticket, "Step 1a: finding search input", find_input, timeout=15, interval=1)

    if search_box is None:
        def find_trigger():
            for sel in ["input[placeholder='Search']", "[aria-label='Search']",
                        "button[aria-label*='search' i]", ".search-box", ".global-search"]:
                el = deep_find(driver, sel)
                if el:
                    return el
            return None

        trigger = poll_until(driver, ticket, "Step 1a: finding search trigger", find_trigger, timeout=15, interval=1)
        if trigger is None:
            raise TimeoutException("Could not find search box/icon (even searching shadow DOM)")
        safe_click(driver, trigger)
        show_screenshot(driver, f"[{ticket}] Step 1a: clicked search trigger")

        search_box = poll_until(driver, ticket, "Step 1b: waiting for expanded search input", find_input, timeout=15, interval=1)
        if search_box is None:
            raise TimeoutException("Expanded search input never appeared after clicking search")

    safe_click(driver, search_box)
    search_box.clear()
    search_box.send_keys(ticket)
    search_box.send_keys(Keys.ENTER)
    show_screenshot(driver, f"[{ticket}] Step 1: searched, waiting for results")
    return search_box


# ==========================================================
# Full ticket download flow — shadow-DOM + iframe aware, with diagnostics
# ==========================================================
def download_ticket_pdf(driver, ticket: str):
    driver.switch_to.default_content()

    log("  Step 1: clicking search box and typing ticket")
    open_search_and_type(driver, ticket)

    log("  Step 2: waiting for loading spinner to clear, then results")
    wait_for_spinner_to_clear(driver, ticket, "Step 2 (search)", timeout=30)

    def check_results(d):
        if is_loading(d):
            return None
        el = deep_find(d, "ul[aria-labelledby='section-EXACT_MATCH_SECTION'] li[data-testclass='sn-global-search-record']")
        if el:
            return el
        return deep_find(d, "li[data-testclass='sn-global-search-record']")

    result = poll_until(
        driver, ticket, "Step 2: waiting for search results",
        lambda: check_results(driver), timeout=60, interval=2
    )
    if result is None:
        show_screenshot(driver, f"[{ticket}] Step 2: FAILED — no results after 60s")
        raise TimeoutException(f"No search results appeared for ticket {ticket} within 60s")

    log("  Step 2: result found, clicking it")
    safe_click(driver, result)
    show_screenshot(driver, f"[{ticket}] Step 2: opened ticket")

    time.sleep(3)

    log("  Step 2b: confirming ticket detail page loaded (checking main page AND any iframes)")
    wait_for_spinner_to_clear(driver, ticket, "Step 2b (opening ticket)", timeout=30)

    menu_selectors = [
        "button.additional-actions-context-menu-button",
        "button[aria-label='additional actions']",
        "button[aria-label='Additional actions']",
        "button[aria-label*='additional actions' i]",
        "button[title*='additional actions' i]",
    ]

    def check_menu_button(d):
        if is_loading(d):
            return None
        for sel in menu_selectors:
            el = deep_find(d, sel)
            if el:
                return el
        return None

    menu_btn = poll_until(
        driver, ticket, "Step 2b: confirming ticket page loaded",
        lambda: find_across_frames(driver, check_menu_button),
        timeout=40, interval=2
    )
    if menu_btn is None:
        log(f"  Step 2b: FAILED for {ticket} — running diagnostics (see {LOG_FILE_PATH})...")
        debug_dump_page_state(driver, ticket)
        show_screenshot(driver, f"[{ticket}] Step 2b: FAILED — see automation_log.txt")
        raise TimeoutException(
            f"Ticket detail page never finished loading for {ticket} "
            f"(no 'additional actions' button found in main page or any iframe). "
            f"See {LOG_FILE_PATH} for full diagnostics."
        )

    log("  Step 3: opening 'Additional actions' menu")
    safe_click(driver, menu_btn)
    show_screenshot(driver, f"[{ticket}] Step 3: menu opened")

    time.sleep(1)

    log("  Step 4: hovering 'Export'")

    def check_export_item(d):
        return deep_find_by_text(d, "Export")

    export_item = poll_until(
        driver, ticket, "Step 4: waiting for Export menu item",
        lambda: find_across_frames(driver, check_export_item),
        timeout=20, interval=1
    )
    if export_item is None:
        debug_dump_page_state(driver, ticket)
        raise TimeoutException(f"'Export' menu item never appeared for ticket {ticket}")
    ActionChains(driver).move_to_element(export_item).perform()
    show_screenshot(driver, f"[{ticket}] Step 4: hovering Export")

    log("  Step 5: clicking 'PDF' in flyout")

    def check_pdf_item(d):
        return deep_find_by_text(d, "PDF")

    pdf_item = poll_until(
        driver, ticket, "Step 5: waiting for PDF flyout item",
        lambda: find_across_frames(driver, check_pdf_item),
        timeout=20, interval=1
    )
    if pdf_item is None:
        debug_dump_page_state(driver, ticket)
        raise TimeoutException(f"'PDF' flyout item never appeared for ticket {ticket}")
    safe_click(driver, pdf_item)
    show_screenshot(driver, f"[{ticket}] Step 5: clicked PDF")

    log("  Step 5b: confirming 'Export to PDF' dialog fully opened")

    def check_dialog_open(d):
        return deep_find(d, "#ok_button")

    ok_btn = poll_until(
        driver, ticket, "Step 5b: waiting for Export dialog to appear",
        lambda: find_across_frames(driver, check_dialog_open),
        timeout=20, interval=1
    )
    if ok_btn is None:
        raise TimeoutException(f"Export dialog 'ok_button' never appeared for ticket {ticket}")

    log("  Step 6: clicking Export button in dialog")
    safe_click(driver, ok_btn)
    show_screenshot(driver, f"[{ticket}] Step 6: export triggered")

    log("  Step 6b: confirming orientation dialog closed")
    wait_until_gone(driver, ticket, "Step 6b (orientation dialog)", "#ok_button", timeout=15)

    log("  Step 7: waiting for Download button (PDF generation can take time)")

    def check_download_button(d):
        if is_loading(d):
            return None
        return deep_find(d, "#download_button")

    download_btn = poll_until(
        driver, ticket, "Step 7: waiting for PDF generation",
        lambda: find_across_frames(driver, check_download_button),
        timeout=90, interval=3
    )
    if download_btn is None:
        raise TimeoutException(f"Download button never appeared for ticket {ticket}")
    show_screenshot(driver, f"[{ticket}] Step 7: export ready")
    safe_click(driver, download_btn)
    log("  Step 8: clicked Download")
    show_screenshot(driver, f"[{ticket}] Step 8: downloading...")

    log("  Step 8b: confirming 'Export Complete' dialog closed before moving on")
    wait_until_gone(driver, ticket, "Step 8b (export complete dialog)", "#download_button", timeout=20)
    show_screenshot(driver, f"[{ticket}] Step 8b: dialog closed, ticket fully done")

    driver.switch_to.default_content()


def wait_for_new_download(download_dir: str, before_files: set, timeout: int = 60):
    start = time.time()
    while time.time() - start < timeout:
        current_files = set(glob.glob(os.path.join(download_dir, "*")))
        new_files = current_files - before_files
        finished = [f for f in new_files if not f.endswith(".crdownload")]
        still_downloading = any(f.endswith(".crdownload") for f in current_files)
        if finished and not still_downloading:
            return finished[0]
        time.sleep(1)
    return None


def rename_downloaded_file(filepath: str, ticket_number: str, ticket_type_label: str, download_dir: str):
    if not filepath or not os.path.exists(filepath):
        return None
    safe_label = ticket_type_label.replace(" ", "_")
    new_name = os.path.join(download_dir, f"{ticket_number}_{safe_label}.pdf")
    try:
        os.rename(filepath, new_name)
        return new_name
    except Exception as e:
        log(f"Error renaming file: {e}")
        return None


def run_automation(sso_id, tickets, ticket_type_label, download_dir, headless):
    driver = None
    try:
        try:
            with open(LOG_FILE_PATH, "w", encoding="utf-8") as f:
                f.write(f"=== Automation run started {datetime.datetime.now()} ===\n")
        except Exception:
            pass

        driver = build_driver(download_dir, headless=headless)
        login(driver, sso_id)

        for ticket in tickets:
            try:
                log(f"Processing ticket: {ticket}")
                before_files = set(glob.glob(os.path.join(download_dir, "*")))

                download_ticket_pdf(driver, ticket)

                downloaded_file = wait_for_new_download(download_dir, before_files, timeout=90)
                if downloaded_file:
                    renamed = rename_downloaded_file(downloaded_file, ticket, ticket_type_label, download_dir)
                    if renamed:
                        log(f"Ticket {ticket} downloaded and renamed to: {os.path.basename(renamed)}")
                    else:
                        log(f"Ticket {ticket} downloaded but renaming failed.")
                else:
                    log(f"Ticket {ticket}: download did not complete within timeout.")

            except TimeoutException as e:
                msg = str(e).strip() or "(no additional details)"
                log(f"Ticket {ticket}: TIMEOUT — {msg}")
                show_screenshot(driver, f"[{ticket}] TIMEOUT")
            except Exception as ex:
                log(f"Ticket {ticket}: ERROR — {type(ex).__name__}: {ex}")
                show_screenshot(driver, f"[{ticket}] ERROR")

        log("All tickets processed.")
        log(f"Full diagnostic log saved to: {LOG_FILE_PATH}")

    except Exception as ex:
        log(f"Error: {str(ex)}")
    finally:
        if driver is not None:
            driver.quit()
        log("Automation completed. Browser closed.")


def make_zip(download_dir: str) -> str:
    zip_path = os.path.join(download_dir, "tickets.zip")
    with zipfile.ZipFile(zip_path, "w") as zf:
        for f in glob.glob(os.path.join(download_dir, "*.pdf")):
            zf.write(f, os.path.basename(f))
    return zip_path


# ==========================================================
# UI
# ==========================================================
st.title("Sequential Ticket Automation")

st.caption(f"Detailed logs are also saved to: `{LOG_FILE_PATH}` — open this file in Notepad "
           f"after a run to copy the full diagnostics if something fails.")

step = st.session_state.step

if step == 0:
    st.header("Login")
    st.session_state.sso_id = st.text_input("SSO ID", value=st.session_state.sso_id)

    st.info(
        "You will complete the actual login (password, MFA/authenticator approval, "
        "any prompts) yourself. You'll be able to watch a live screenshot feed of "
        "the browser on the next page."
    )

    st.session_state.headless = st.checkbox(
        "Run headless (you can still watch via the live screenshot feed)",
        value=st.session_state.headless,
    )

    if st.button("Next"):
        if st.session_state.sso_id:
            st.session_state.step = 1
            st.rerun()
        else:
            st.warning("Please enter your SSO ID.")

elif step == 1:
    st.subheader("Step 1: Select platform")
    st.session_state.platform = st.radio(
        "Platform",
        list(TICKET_TYPE_LABELS.keys()),
        index=list(TICKET_TYPE_LABELS.keys()).index(st.session_state.platform),
    )

    st.subheader("Step 2: Select ticket type")
    labels = TICKET_TYPE_LABELS[st.session_state.platform]
    label_values = list(labels.values())
    default_label = labels[st.session_state.ticket_type_key] if st.session_state.ticket_type_key in labels else label_values[0]
    choice_label = st.radio("Ticket type", label_values, index=label_values.index(default_label))
    st.session_state.ticket_type_key = [k for k, v in labels.items() if v == choice_label][0]

    st.subheader("Step 3: Paste TicketNumber column (one per line)")
    raw = st.text_area("Tickets", height=200)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back"):
            st.session_state.step = 0
            st.rerun()
    with col2:
        if st.button("Load Tickets"):
            tickets = [l.strip() for l in raw.splitlines() if l.strip()]
            if tickets:
                st.session_state.tickets = tickets
                st.session_state.step = 2
                st.rerun()
            else:
                st.warning("No tickets found in pasted text.")

elif step == 2:
    st.subheader("Ready to run")

    labels = TICKET_TYPE_LABELS[st.session_state.platform]
    ticket_type_label = labels[st.session_state.ticket_type_key]

    st.write(f"**Platform:** {st.session_state.platform}")
    st.write(f"**Ticket type:** {ticket_type_label}")

    n = len(st.session_state.tickets)
    est = 60 + n * 90
    m, s = divmod(est, 60)
    st.info(f"Estimated time (excluding manual login): up to {m} min {s} sec")

    st.write("**Ticket List**")
    st.write(st.session_state.tickets)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Back"):
            st.session_state.step = 1
            st.rerun()
    with col2:
        start_clicked = st.button("Start Process")

    col_left, col_right = st.columns([1, 1])
    with col_left:
        st.write("**Live Browser View**")
        _image_placeholder = st.empty()
    with col_right:
        st.write("**Process Log**")
        _log_placeholder = st.empty()
        _log_placeholder.text("\n".join(st.session_state.logs[-400:]))

    if start_clicked:
        tmp_dir = tempfile.mkdtemp()
        st.session_state.download_dir = tmp_dir
        st.session_state.logs = []
        run_automation(
            st.session_state.sso_id,
            st.session_state.tickets,
            ticket_type_label,
            tmp_dir,
            st.session_state.headless,
        )

    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
            log_content = f.read()
        st.download_button("Download full diagnostic log (automation_log.txt)",
                            log_content, file_name="automation_log.txt")

    if st.session_state.download_dir and glob.glob(
        os.path.join(st.session_state.download_dir, "*.pdf")
    ):
        zip_path = make_zip(st.session_state.download_dir)
        with open(zip_path, "rb") as f:
            st.download_button("Download all tickets (zip)", f, file_name="tickets.zip")