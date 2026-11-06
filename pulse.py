import streamlit as st
import time
import os
import glob
import zipfile
import tempfile
import datetime
import shutil
import subprocess

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

import io
import re
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

st.set_page_config(page_title="CM Automation", layout="wide")

# ==========================================================
# Global Purple & White Theme (professional, black outlines)
# ==========================================================
st.markdown(
    """
    <style>
    :root {
        --brand-purple: #5B2A86;
        --brand-purple-dark: #3E1F5B;
        --brand-purple-light: #F3EAFB;
        --brand-white: #FFFFFF;
        --brand-black: #000000;
    }

    /* App background */
    .stApp {
        background-color: var(--brand-white);
    }

    /* Header / top banner */
    header[data-testid="stHeader"] {
        background-color: var(--brand-purple);
        border-bottom: 2px solid #5B2A86;
    }

    /* Sidebar (if used) */
    section[data-testid="stSidebar"] {
        background-color: var(--brand-purple-dark);
        border-right: 2px solid #5B2A86;
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }

    /* Titles / headers */
    h1, h2, h3 {
        color: var(--brand-purple-dark) !important;
        font-weight: 700 !important;
    }

    /* Main title banner */
    .stApp h1:first-of-type {
        background-color: var(--brand-purple);
        color: #FFFFFF !important;
        padding: 18px 24px;
        border: 2px solid #5B2A86;
        border-radius: 8px;
    }

    /* Captions / info text */
    .stCaption, .stMarkdown p, label, .stTextInput label, .stTextArea label {
        color: #2B2B2B !important;
    }

    /* Text inputs / text areas */
    .stTextInput input,
    .stTextArea textarea {
        background-color: #FFFFFF !important;
        color: #000000 !important;
        border: 1.5px solid #5B2A86 !important;
        border-radius: 6px !important;
    }
    .stTextInput input:focus,
    .stTextArea textarea:focus {
        border: 1.5px solid var(--brand-purple) !important;
        box-shadow: 0 0 0 1px var(--brand-purple) !important;
    }

    /* Checkbox */
    .stCheckbox label {
        color: #2B2B2B !important;
    }

    /* Buttons */
    .stButton > button {
        background-color: var(--brand-purple) !important;
        color: #FFFFFF !important;
        border: 1.5px solid #5B2A86 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: background-color 0.15s ease-in-out;
    }
    .stButton > button:hover {
        background-color: var(--brand-purple-dark) !important;
        border: 1.5px solid #5B2A86 !important;
        color: #FFFFFF !important;
    }
    .stButton > button:active {
        background-color: #2E1440 !important;
    }

    /* Download buttons */
    .stDownloadButton > button {
        background-color: #FFFFFF !important;
        color: var(--brand-purple-dark) !important;
        border: 1.5px solid #5B2A86 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    .stDownloadButton > button:hover {
        background-color: var(--brand-purple-light) !important;
        color: var(--brand-purple-dark) !important;
        border: 1.5px solid #5B2A86 !important;
    }

    /* Info / warning / success boxes -> unify with purple/white/black outline */
    div[data-testid="stAlert"] {
        border: 1.5px solid #5B2A86 !important;
        border-radius: 8px !important;
        background-color: var(--brand-purple-light) !important;
        color: #2B2B2B !important;
    }
    div[data-testid="stAlert"] p {
        color: #2B2B2B !important;
    }

    /* Expanders (ticket folders) styled like cards */
    div[data-testid="stExpander"] {
        border: 1.5px solid #5B2A86 !important;
        border-radius: 10px !important;
        background-color: #FFFFFF !important;
        overflow: hidden;
    }
    div[data-testid="stExpander"] summary {
        background-color: var(--brand-purple) !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
    }

    /* Text areas / code / log box */
    .stCodeBlock, pre {
        background-color: #FAF7FD !important;
        border: 1.5px solid #5B2A86 !important;
        border-radius: 6px !important;
    }

    /* Columns acting as cards (log / browser view) */
    div[data-testid="column"] > div {
        border-radius: 10px;
    }

    /* Horizontal rule accents */
    hr {
        border-top: 2px solid var(--brand-purple) !important;
    }

    /* JSON / list viewer (st.write on lists/dicts) -> Purple & White theme, black outline */
    div[data-testid="stJson"] {
        background-color: #FFFFFF !important;
        border: 1.5px solid #5B2A86 !important;
        border-radius: 8px !important;
        padding: 10px !important;
    }
    div[data-testid="stJson"] * {
        background-color: transparent !important;
    }
    div[data-testid="stJson"] pre,
    div[data-testid="stJson"] code {
        background-color: #FFFFFF !important;
        color: #2B2B2B !important;
    }
    /* react-json-view (used internally by stJson) text colors */
    div[data-testid="stJson"] .object-key-val,
    div[data-testid="stJson"] .variable-value,
    div[data-testid="stJson"] span {
        color: #2B2B2B !important;
    }
    div[data-testid="stJson"] .object-key {
        color: var(--brand-purple-dark) !important;
        font-weight: 600 !important;
    }
    div[data-testid="stJson"] .string-value {
        color: #B5651D !important; /* keep string values legible, warm tone on white */
    }
    div[data-testid="stJson"] svg {
        fill: var(--brand-purple) !important;
    }
    /* Copy-to-clipboard icon / toolbar inside stJson */
    div[data-testid="stJson"] [class*="icon-container"] {
        background-color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

TICKET_TYPE_LABELS = {
    "Pulse Tickets": {"Yes": "Normal Ticket", "No": "GXP Standard Ticket"},
    "Other Pulse Tickets": {"Yes": "Standard Ticket", "No": "GXP Normal Ticket"},
}

LOG_FILE_PATH = os.path.join(os.getcwd(), "automation_log.txt")

defaults = {
    "step": 0,
    "sso_id": "",
    "password": "",
    "platform": "Pulse Tickets",
    "ticket_type_key": "Yes",
    "tickets": [],
    "logs": [],
    "download_dir": None,
    "headless": True,
    "excel_bytes": None,
    "excel_filename": "change_management_report.xlsx",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

_log_placeholder = None
_image_placeholder = None


def render_process_log(placeholder, logs):
    """Render the process log inside a scrollable, purple-outlined box (display-only helper)."""
    if placeholder is None:
        return
    log_text = "\n".join(logs[-400:])
    escaped = (
        log_text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )
    placeholder.markdown(
        f'''
        <div style="
            background-color:#FFFFFF;
            border:1.5px solid #5B2A86;
            outline:1.5px solid #5B2A86;
            outline-offset:-1.5px;
            border-radius:8px;
            padding:10px 12px;
            height:400px;
            overflow-y:auto;
            font-family:monospace;
            font-size:13px;
            color:#2B2B2B;
            white-space:pre-wrap;
            word-break:break-word;
        ">{escaped}</div>
        ''',
        unsafe_allow_html=True,
    )


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
        render_process_log(_log_placeholder, st.session_state.logs)


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

DEEP_TEXT_CONTAINS_JS = """
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
function deepFindByTextContains(text, root) {
    root = root || document;
    const all = root.querySelectorAll('*');
    for (const el of all) {
        if (el.shadowRoot) {
            const res = deepFindByTextContains(text, el.shadowRoot);
            if (res) return res;
        }
        const t = getDirectText(el);
        if (t.indexOf(text) !== -1 && isVisible(el)) {
            return el;
        }
    }
    return null;
}
return deepFindByTextContains(arguments[0]);
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


def deep_find_by_text_contains(driver, text):
    try:
        return driver.execute_script(DEEP_TEXT_CONTAINS_JS, text)
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
# Iframe-aware search wrapper (iframes found via shadow-DOM-aware search)
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


def find_all_across_frames(driver, finder_fn, max_depth=3):
    """Like find_across_frames, but collects ALL matches from every frame
    (top-level + every nested iframe, found via shadow-DOM-aware search).
    NOTE: elements returned here may go stale once the driver switches
    frames again -- do not read attributes off them after calling this."""
    results = []

    def collect():
        try:
            found = finder_fn(driver)
        except Exception:
            found = None
        if found:
            results.extend(found)

    driver.switch_to.default_content()
    collect()

    def search_frames(depth):
        if depth > max_depth:
            return
        frames = deep_find_all(driver, "iframe")
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
            except Exception:
                continue
            collect()
            search_frames(depth + 1)
            try:
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()

    search_frames(1)
    driver.switch_to.default_content()
    return results


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
# Diagnostics
# ==========================================================
def debug_dump_page_state(driver, ticket):
    try:
        driver.switch_to.default_content()
        log(f"  [DEBUG] === Diagnostics for ticket {ticket} ===")
        log(f"  [DEBUG] Current URL: {driver.current_url}")
        log(f"  [DEBUG] Page title: {driver.title}")
        log(f"  [DEBUG] Number of window handles: {len(driver.window_handles)}")

        iframes = deep_find_all(driver, "iframe")
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

        log(f"  [DEBUG] === End diagnostics for {ticket} ===")

    except Exception as e:
        log(f"  [DEBUG] debug_dump_page_state itself failed: {e}")
    finally:
        driver.switch_to.default_content()


# ==========================================================
# Binary detection (HARDENED: known paths -> PATH -> filesystem-wide search)
# ==========================================================
def _filesystem_search(patterns, roots=("/usr", "/opt", "/snap", "/root", "/home")):
    """Last-resort recursive search for a binary matching any of `patterns`."""
    for root in roots:
        if not os.path.isdir(root):
            continue
        for pattern in patterns:
            try:
                matches = glob.glob(os.path.join(root, "**", pattern), recursive=True)
            except Exception:
                matches = []
            for m in matches:
                if os.path.isfile(m) and os.access(m, os.X_OK) or os.path.isfile(m):
                    return m
    return None


def find_browser_binary():
    candidates = [
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/lib/chromium/chromium",
        "/usr/lib/chromium-browser/chromium-browser",
        "/snap/bin/chromium",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    for name in ["chromium", "chromium-browser", "google-chrome", "google-chrome-stable"]:
        found = shutil.which(name)
        if found:
            return found
    found = _filesystem_search(["chromium", "chromium-browser", "google-chrome*"])
    return found


def find_chromedriver_binary():
    candidates = [
        "/usr/bin/chromedriver",
        "/usr/lib/chromium/chromedriver",
        "/usr/lib/chromium-browser/chromedriver",
        "/snap/bin/chromedriver",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    found = shutil.which("chromedriver")
    if found:
        return found
    found = _filesystem_search(["chromedriver"])
    if found:
        try:
            os.chmod(found, 0o755)
        except Exception:
            pass
    return found


def run_environment_diagnostics():
    log("=== Environment diagnostics ===")
    try:
        which_chromium = shutil.which("chromium") or shutil.which("chromium-browser")
        which_chromedriver = shutil.which("chromedriver")
        log(f"  which chromium: {which_chromium}")
        log(f"  which chromedriver: {which_chromedriver}")
    except Exception as e:
        log(f"  which check failed: {e}")

    try:
        result = subprocess.run(
            ["dpkg", "-l"], capture_output=True, text=True, timeout=10
        )
        lines = [l for l in result.stdout.splitlines() if "chrom" in l.lower()]
        if lines:
            log("  dpkg packages matching 'chrom':")
            for l in lines:
                log(f"    {l}")
        else:
            log("  dpkg: no chromium/chromedriver packages found "
                "(packages.txt may be missing or app wasn't rebooted after adding it)")
    except Exception as e:
        log(f"  dpkg check failed or unavailable: {e}")

    log("=== End environment diagnostics ===")


# ==========================================================
# Selenium driver setup
# ==========================================================
def build_driver(download_dir: str, headless: bool = True):
    run_environment_diagnostics()

    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    chromium_path = find_browser_binary()
    chromedriver_path = find_chromedriver_binary()

    log(f"Detected chromium binary: {chromium_path}")
    log(f"Detected chromedriver binary: {chromedriver_path}")

    if chromium_path:
        options.binary_location = chromium_path

    driver = None

    if chromedriver_path:
        try:
            service = Service(chromedriver_path)
            driver = webdriver.Chrome(service=service, options=options)
            log(f"Using detected chromedriver at {chromedriver_path}")
        except Exception as e:
            log(f"Detected chromedriver failed to launch: {e}")
    else:
        log("No chromedriver binary found via known paths, system PATH, or filesystem search. "
            "This almost certainly means packages.txt is missing 'chromium-driver' "
            "or the app has not been rebooted since adding it.")

    if driver is None:
        try:
            driver = webdriver.Chrome(options=options)
            log("Using Selenium Manager auto-resolved driver")
        except Exception as e:
            log(f"Selenium Manager failed: {e}")

    if driver is None:
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            log("Using webdriver-manager driver")
        except Exception as e:
            log(f"webdriver-manager also failed: {e}")
            raise RuntimeError(
                "Could not start Chrome/Chromium via any method. "
                "Check that packages.txt contains 'chromium' and 'chromium-driver', "
                "and that you have rebooted the app (Manage app -> Reboot) after adding it."
            )

    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_dir
        })
        log(f"Download behavior explicitly set via CDP to: {download_dir}")
    except Exception as e:
        log(f"Warning: could not set CDP download behavior: {e}")

    return driver


# ==========================================================
# Login — fully automated username + password, then wait for MFA push
# ==========================================================
def login(driver, sso_id: str, password: str):
    login_url = "https://pulse.service-now.com/now/nav/ui/home"
    expected_domain = "pulse.service-now.com"

    log("Opening login page...")
    driver.get(login_url)
    show_screenshot(driver, "Login page loaded")

    try:
        username_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))
        )
        username_field.send_keys(sso_id)
        show_screenshot(driver, "Username entered")
        log("Entered SSO ID, clicking Next...")

        next_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Next')]"))
        )
        next_btn.click()
    except TimeoutException:
        show_screenshot(driver, "FAILED at username step")
        raise TimeoutException("Could not find username field or Next button on login page")

    log("Waiting 2 seconds before entering password...")
    time.sleep(2)
    show_screenshot(driver, "Waiting before password entry")

    try:
        password_field = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
        )
        password_field.send_keys(password)
        show_screenshot(driver, "Password entered")
        log("Entered password, submitting...")

        login_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log In & Remember Me')]"))
        )
        login_btn.click()
        log("Clicked 'Log In & Remember Me'")
    except TimeoutException:
        show_screenshot(driver, "FAILED at password step")
        raise TimeoutException("Could not find password field or Login button")

    log("Password submitted. A push notification should now be sent to your "
        "authenticator app. Please approve it on your phone.")
    log("Waiting for you to approve the push notification... (up to 5 minutes)")

    start = time.time()
    timeout = 300
    while time.time() - start < timeout:
        if expected_domain in driver.current_url and "login" not in driver.current_url.lower():
            break
        show_screenshot(driver, "Waiting for MFA push approval on your phone...")
        time.sleep(3)
    else:
        raise TimeoutException(
            "Login was not completed within 5 minutes. "
            "Did you approve the push notification on your phone?"
        )

    log("Login successful! Resuming automation...")
    show_screenshot(driver, "Logged in successfully")


# ==========================================================
# Step 1: search box (shadow-DOM aware)
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
# Per-ticket folder helpers
# ==========================================================
def sanitize_ticket_name(ticket: str) -> str:
    cleaned = "".join(c for c in ticket.strip() if c.isalnum() or c in ("-", "_"))
    return cleaned or "ticket"


def get_ticket_folder(download_dir: str, ticket: str) -> str:
    """Every ticket gets its own dedicated folder: <download_dir>/<TICKET>/
    The ticket's PDF goes directly inside it, and its attachments go inside
    a further 'attachments' subfolder."""
    folder = os.path.join(download_dir, sanitize_ticket_name(ticket))
    os.makedirs(folder, exist_ok=True)
    return folder


def get_ticket_attachments_folder(download_dir: str, ticket: str) -> str:
    folder = os.path.join(get_ticket_folder(download_dir, ticket), "attachments")
    os.makedirs(folder, exist_ok=True)
    return folder


def cleanup_empty_attachments_folder(attachments_dir: str):
    """If a ticket genuinely had zero attachments, the 'attachments'
    subfolder was still created up-front (needed as a destination in case
    there WERE any). Remove it afterward if it ended up empty, so tickets
    with no attachments don't leave a clutter empty folder behind in the
    final ZIP/download listing."""
    try:
        if os.path.isdir(attachments_dir) and not os.listdir(attachments_dir):
            os.rmdir(attachments_dir)
    except Exception:
        pass


def build_attachment_filename(ticket, original_name):
    """Prefix the ticket number onto the attachment's original filename,
    unless it's already present (ServiceNow sometimes already includes it)."""
    original_name = (original_name or "attachment").strip()
    if original_name.upper().startswith(ticket.strip().upper()):
        return original_name
    return f"{ticket.strip()}_{original_name}"


# ==========================================================
# Attachment download helpers
# ==========================================================
ATTACHMENT_LINK_SELECTOR_PRIMARY = "a[aria-label^='Download ' i]"
ATTACHMENT_LINK_SELECTOR_FALLBACK = "a[href*='sys_attachment.do']"


def _attachment_link_name(el):
    """Derive a filename for an attachment <a> element. MUST be called
    immediately after locating the element, in the same frame context it
    was found in."""
    name = None
    try:
        aria = el.get_attribute("aria-label") or ""
    except Exception:
        aria = ""
    if aria.lower().startswith("download "):
        name = aria[len("Download "):].strip()
    if not name:
        try:
            text = (el.text or "").strip()
        except Exception:
            text = ""
        if text:
            name = text
    if not name:
        try:
            href = el.get_attribute("href") or ""
        except Exception:
            href = ""
        if href:
            name = href.split("/")[-1].split("?")[0]
    return name or None


def _is_view_link(el):
    """Filters out the separate 'View <filename>' links that sit next to
    each 'Download <filename>' link -- we only want to click Download."""
    try:
        aria = (el.get_attribute("aria-label") or "").lower()
        if aria.startswith("view "):
            return True
    except Exception:
        pass
    try:
        cls = (el.get_attribute("class") or "").lower()
        if "view_" in cls or "view-" in cls:
            return True
    except Exception:
        pass
    return False


def find_next_unclicked_attachment(driver, downloaded_names):
    """Finds the first attachment Download link not yet in `downloaded_names`,
    returning it WITHOUT switching frames afterward, so it's still valid for
    an immediate click()."""

    def finder(d):
        els = deep_find_all(d, ATTACHMENT_LINK_SELECTOR_PRIMARY)
        if not els:
            els = [e for e in deep_find_all(d, ATTACHMENT_LINK_SELECTOR_FALLBACK)
                   if not _is_view_link(e)]
        for el in els:
            name = _attachment_link_name(el)
            if name and name not in downloaded_names:
                el._attachment_name_cache = name
                return el
        return None

    return find_across_frames(driver, finder)


def open_manage_attachments(driver, ticket):
    """Some tickets show every attachment inline in the header already
    (no popup needed); others truncate the list and require clicking
    'Manage Attachments (N)' to reveal the rest. Returns True if any
    Download link is present anywhere on the page afterward, False if the
    ticket has zero attachments (this is a NORMAL, expected outcome, not
    an error)."""

    def check_any_link(d):
        el = deep_find(d, ATTACHMENT_LINK_SELECTOR_PRIMARY)
        if el:
            return el
        for c in deep_find_all(d, ATTACHMENT_LINK_SELECTOR_FALLBACK):
            if not _is_view_link(c):
                return c
        return None

    already_visible = find_across_frames(driver, check_any_link)
    if already_visible:
        log(f"  Step 2c: attachment link(s) already visible inline for {ticket}")
        driver.switch_to.default_content()
        return True

    def find_trigger(d):
        el = deep_find(d, "a#header_attachment_list_label")
        if el:
            return el
        el = deep_find_by_text_contains(d, "Manage Attachments")
        if el:
            return el
        return deep_find_by_text_contains(d, "Attachments (")

    trigger = poll_until(
        driver, ticket, "Step 2c: locating 'Manage Attachments' link",
        lambda: find_across_frames(driver, find_trigger),
        timeout=15, interval=1
    )
    if trigger is None:
        log(f"  Step 2c: 'Manage Attachments' link NOT FOUND for {ticket} "
            f"(this ticket has 0 attachments -- normal, not an error)")
        driver.switch_to.default_content()
        return False

    log(f"  Step 2c: found 'Manage Attachments' trigger for {ticket}, clicking it")
    try:
        safe_click(driver, trigger)
    except Exception as e:
        log(f"  Step 2c: could not click 'Manage Attachments' trigger for {ticket}: {e}")
        driver.switch_to.default_content()
        return False
    show_screenshot(driver, f"[{ticket}] Step 2c: opened Manage Attachments dialog")

    opened = poll_until(
        driver, ticket, "Step 2c: waiting for Attachments dialog to load",
        lambda: find_across_frames(driver, check_any_link),
        timeout=20, interval=1
    )
    driver.switch_to.default_content()
    if not opened:
        log(f"  Step 2c: Attachments popup opened but no Download links "
            f"rendered within 20s for {ticket} (ticket may genuinely have 0 "
            f"attachments, or the popup listed them differently than expected)")
    return bool(opened)


def close_manage_attachments_dialog(driver, ticket):
    """Closes the popup via an explicit close/X button only. Deliberately
    never sends ESCAPE -- in ServiceNow's classic UI, ESC can navigate away
    from the ticket entirely."""

    def check_open(d):
        el = deep_find(d, ATTACHMENT_LINK_SELECTOR_PRIMARY)
        if el:
            return el
        return deep_find(d, "div.modal-dialog, div[role='dialog']")

    still_open = find_across_frames(driver, check_open)
    if not still_open:
        driver.switch_to.default_content()
        return

    def find_close_btn(d):
        for sel in ["button[aria-label='Close']", "button[title='Close']",
                    ".close", "[aria-label*='close' i]", "button.btn-close",
                    "img[alt='Close']", "a[title='Close']", ".glide_dialog_close"]:
            el = deep_find(d, sel)
            if el:
                return el
        return None

    btn = find_across_frames(driver, find_close_btn)
    if btn:
        try:
            safe_click(driver, btn)
            log(f"  Closed Attachments dialog for {ticket}")
        except Exception as e:
            log(f"  Could not close Attachments dialog for {ticket}: {e} (leaving it open)")
    else:
        log(f"  No explicit close button found for Attachments dialog on {ticket}; "
            f"leaving it open (non-blocking) rather than sending ESCAPE.")

    driver.switch_to.default_content()
    time.sleep(1)


def download_ticket_attachments(driver, ticket, download_dir, attachments_dir):
    """Finds and downloads every attachment for the ticket ONE AT A TIME.
    If the ticket has zero attachments, this cleanly logs that and returns
    -- no error, no retries, no impact on the PDF that was already
    downloaded before this step runs."""

    opened = open_manage_attachments(driver, ticket)
    if not opened:
        log(f"  No attachments found for {ticket} (skipping attachment download)")
        return

    log(f"  Step 2c: attachment link(s) confirmed present for {ticket}, downloading...")
    downloaded_names = set()
    max_iterations = 25
    consecutive_click_failures = 0

    for _ in range(max_iterations):
        target = find_next_unclicked_attachment(driver, downloaded_names)

        if target is None:
            if not downloaded_names:
                log(f"  Step 2c: WARNING -- no attachment Download links could be "
                    f"matched for {ticket} even though the trigger/popup opened.")
            break

        target_name = getattr(target, "_attachment_name_cache", None) or "attachment"

        log(f"  Step 2c: downloading attachment '{target_name}' for {ticket}")
        before_files = _list_files_only(get_candidate_download_dirs(download_dir))

        try:
            safe_click(driver, target)
            consecutive_click_failures = 0
        except Exception as e:
            log(f"    Could not click attachment link '{target_name}' for {ticket}: {e}")
            downloaded_names.add(target_name)
            consecutive_click_failures += 1
            driver.switch_to.default_content()
            if consecutive_click_failures >= 3:
                log(f"    Too many consecutive click failures for {ticket}, stopping")
                break
            continue

        driver.switch_to.default_content()

        show_screenshot(driver, f"[{ticket}] downloading attachment: {target_name}")
        downloaded_file = wait_for_new_download(download_dir, before_files, timeout=90)

        if downloaded_file:
            new_name = build_attachment_filename(ticket, target_name)
            dest_path = os.path.join(attachments_dir, new_name)
            if os.path.exists(dest_path):
                root_name, ext = os.path.splitext(new_name)
                dest_path = os.path.join(attachments_dir, f"{root_name}_{int(time.time())}{ext}")
            try:
                shutil.move(downloaded_file, dest_path)
                log(f"    Attachment saved as: {os.path.relpath(dest_path, download_dir)}")
            except Exception as e:
                log(f"    Attachment downloaded but move/rename failed for {ticket}: {e}. "
                    f"File left at: {downloaded_file}")
        else:
            log(f"    Attachment '{target_name}' did not finish downloading within 90s for {ticket}")

        downloaded_names.add(target_name)
        time.sleep(1)
    else:
        log(f"  Reached max attachment-download iterations ({max_iterations}) for {ticket}")

    log(f"  Finished attachment downloads for {ticket}: {len(downloaded_names)} attachment(s) handled")
    close_manage_attachments_dialog(driver, ticket)


# ==========================================================
# Full ticket download flow
# ==========================================================
def download_ticket_pdf(driver, ticket: str, download_dir: str):
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

    log("  Step 2b: confirming ticket detail page loaded")
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
        log(f"  Step 2b: FAILED for {ticket} — running diagnostics...")
        debug_dump_page_state(driver, ticket)
        show_screenshot(driver, f"[{ticket}] Step 2b: FAILED — see automation_log.txt")
        raise TimeoutException(
            f"Ticket detail page never finished loading for {ticket}. "
            f"See {LOG_FILE_PATH} for diagnostics."
        )

    log("  Step 3: opening 'Additional actions' menu")
    fresh_menu_btn = find_across_frames(driver, check_menu_button) or menu_btn
    safe_click(driver, fresh_menu_btn)
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

    log("  Step 7: waiting for Download button to appear AND become enabled "
        "(PDF generation can take time)")

    _logged_once = {"done": False}

    def check_download_button(d):
        if is_loading(d):
            return None
        el = deep_find(d, "#download_button")
        if el is None:
            return None
        try:
            disabled_attr = el.get_attribute("disabled")
            aria_disabled = el.get_attribute("aria-disabled")
            classes = (el.get_attribute("class") or "").lower()
            if not _logged_once["done"]:
                log(f"    [debug] download_button found: disabled={disabled_attr!r} "
                    f"aria-disabled={aria_disabled!r} class={classes!r} "
                    f"is_enabled={el.is_enabled()}")
                _logged_once["done"] = True
            if disabled_attr or aria_disabled == "true" or "disabled" in classes:
                return None
            if not el.is_enabled():
                return None
        except Exception:
            return None
        return el

    download_btn = poll_until(
        driver, ticket, "Step 7: waiting for PDF generation to finish (button enabled)",
        lambda: find_across_frames(driver, check_download_button),
        timeout=120, interval=3
    )
    if download_btn is None:
        raise TimeoutException(f"Download button never became enabled for ticket {ticket}")
    show_screenshot(driver, f"[{ticket}] Step 7: export ready (button enabled)")

    time.sleep(1)

    safe_click(driver, download_btn)
    log("  Step 8: clicked Download")
    show_screenshot(driver, f"[{ticket}] Step 8: downloading...")

    log("  Step 8b: waiting for 'Export Complete' dialog to close on its own "
        "(this can take a while depending on file size)")
    dialog_closed = wait_until_gone(
        driver, ticket, "Step 8b (export complete dialog)", "#download_button", timeout=90
    )
    if dialog_closed:
        log(f"  Step 8b: dialog closed normally for {ticket}")
    else:
        log(f"  Step 8b: dialog did NOT close within 90s for {ticket} — "
            f"will check if the file downloaded anyway, and try to force-close the dialog.")
    show_screenshot(driver, f"[{ticket}] Step 8b: done waiting on dialog")

    driver.switch_to.default_content()

    log("  Step 9: downloading all attachments (if any) after PDF export")
    attachments_dir = get_ticket_attachments_folder(download_dir, ticket)
    try:
        download_ticket_attachments(driver, ticket, download_dir, attachments_dir)
    except Exception as ex:
        log(f"  Attachment download step failed for {ticket}: {type(ex).__name__}: {ex}")
        show_screenshot(driver, f"[{ticket}] Attachment download step failed")
    finally:
        # If the ticket had 0 attachments, don't leave a clutter empty
        # "attachments" folder behind in the final output.
        cleanup_empty_attachments_folder(attachments_dir)

    driver.switch_to.default_content()


# ==========================================================
# Download detection — checks multiple candidate folders as a safety net
# ==========================================================
def get_candidate_download_dirs(download_dir: str):
    candidates = [download_dir]
    home = os.path.expanduser("~")
    default_downloads = os.path.join(home, "Downloads")
    if default_downloads not in candidates:
        candidates.append(default_downloads)
    return candidates


def _list_files_only(dirs):
    """Returns only FILES (never directories) found directly inside any of
    `dirs`. Critical: per-ticket folders (e.g. <download_dir>/CHG12345/ and
    its attachments/ subfolder) get created via os.makedirs() -- a plain
    glob.glob(dir/"*") would also match those new folders, which could get
    mistaken for the downloaded PDF file itself. Filtering to
    os.path.isfile() only avoids that."""
    result = set()
    for d in dirs:
        for p in glob.glob(os.path.join(d, "*")):
            if os.path.isfile(p):
                result.add(p)
    return result


def wait_for_new_download(download_dir: str, before_files: set, timeout: int = 120):
    candidate_dirs = get_candidate_download_dirs(download_dir)
    start = time.time()
    while time.time() - start < timeout:
        current_files = _list_files_only(candidate_dirs)
        new_files = current_files - before_files
        finished = [f for f in new_files if not f.endswith(".crdownload") and not f.endswith(".tmp")]
        still_downloading = any(
            f.endswith(".crdownload") or f.endswith(".tmp") for f in current_files
        )
        if finished and not still_downloading:
            return finished[0]
        time.sleep(1)
    return None


def rename_downloaded_file(filepath: str, ticket_number: str, ticket_type_label: str, download_dir: str):
    """Moves the exported PDF into the ticket's dedicated folder
    (<download_dir>/<TICKET>/<TICKET>_<label>.pdf), separate from its
    attachments (which live in <download_dir>/<TICKET>/attachments/)."""
    if not filepath or not os.path.exists(filepath):
        return None
    safe_label = ticket_type_label.replace(" ", "_")
    ticket_folder = get_ticket_folder(download_dir, ticket_number)
    new_name = os.path.join(ticket_folder, f"{ticket_number}_{safe_label}.pdf")
    try:
        shutil.move(filepath, new_name)
        return new_name
    except Exception as e:
        log(f"Error moving/renaming file: {e}")
        return None


def close_export_dialog(driver, ticket):
    def check_still_open(d):
        return deep_find(d, "#download_button")

    still_open = find_across_frames(driver, check_still_open)
    if not still_open:
        driver.switch_to.default_content()
        return

    def find_cancel_or_close(d):
        el = deep_find_by_text(d, "Cancel")
        if el:
            return el
        for sel in ["button[aria-label='Close']", "button[title='Close']", ".close", "[aria-label*='close' i]"]:
            el = deep_find(d, sel)
            if el:
                return el
        return None

    btn = find_across_frames(driver, find_cancel_or_close)
    if btn:
        try:
            safe_click(driver, btn)
            log(f"  Force-closed 'Export Complete' dialog for {ticket}")
            show_screenshot(driver, f"[{ticket}] Dialog force-closed")
        except Exception as e:
            log(f"  Could not click Cancel/Close on dialog for {ticket}: {e}")
    else:
        log(f"  Dialog still open for {ticket} but no Cancel/Close button found")
    driver.switch_to.default_content()


def run_automation(sso_id, password, tickets, ticket_type_label, download_dir, headless):
    driver = None
    try:
        try:
            with open(LOG_FILE_PATH, "w", encoding="utf-8") as f:
                f.write(f"=== Automation run started {datetime.datetime.now()} ===\n")
        except Exception:
            pass

        driver = build_driver(download_dir, headless=headless)
        login(driver, sso_id, password)

        for ticket in tickets:
            try:
                log(f"Processing ticket: {ticket}")
                before_files = _list_files_only(get_candidate_download_dirs(download_dir))

                download_ticket_pdf(driver, ticket, download_dir)

                downloaded_file = wait_for_new_download(download_dir, before_files, timeout=120)

                close_export_dialog(driver, ticket)

                if downloaded_file:
                    renamed = rename_downloaded_file(downloaded_file, ticket, ticket_type_label, download_dir)
                    if renamed:
                        log(f"Ticket {ticket} downloaded and renamed to: {os.path.relpath(renamed, download_dir)}")
                    else:
                        log(f"Ticket {ticket} downloaded but renaming failed. File was at: {downloaded_file}")
                else:
                    log(f"Ticket {ticket}: download did not complete within 120s. "
                        f"Checked folders: {get_candidate_download_dirs(download_dir)}")

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


# ==========================================================
# Excel report generation (Change Management report) — extracted from each
# downloaded ticket PDF, merged into this tool so the report is produced
# automatically as part of the same run, no separate upload step needed.
# ==========================================================
def extract_field_value(text, field_label, length_limit=None):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        value = text[start_index:end_index].strip()
        if length_limit:
            value = value[:length_limit]
        return value
    except ValueError:
        return None


def extract_requestor_value(text, field_label):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        value = text[start_index:end_index].strip()
        end_paren_index = value.find(")")
        if end_paren_index != -1:
            value = value[: end_paren_index + 1]
        if "Type:" in value:
            value = value.split("Type:")[0].strip()
        return value
    except ValueError:
        return None


def extract_value_after_line(text, field_label):
    try:
        start_index = text.index(field_label) + len(field_label)
        end_index = text.index("\n", start_index)
        next_line_start = end_index + 1
        next_line_end = text.find("\n", next_line_start)
        if next_line_end == -1:
            next_line_end = len(text)
        value = text[next_line_start:next_line_end].strip()
        return value
    except ValueError:
        return None


APPROVAL_TS_IN_COMMENT = re.compile(r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s*-\s*")
ANY_TS = re.compile(r"\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}")


def _norm_header(h):
    return "".join(str(h or "").split()).lower()


def _find_approval_columns(header_row):
    """Locate the columns of the Approval List table by header NAME instead of
    a hard-coded position (so a shifted/merged column can never make us read
    the wrong one, e.g. 'Approval for' = CHG number instead of a date)."""
    cols = {}
    for idx, h in enumerate(header_row):
        n = _norm_header(h)
        if n == "state":
            cols["state"] = idx
        elif n == "approver":
            cols["approver"] = idx
        elif n == "assignmentgroup":
            cols["group"] = idx
        elif n == "comments":
            cols["comments"] = idx
        elif n == "approvalfor":
            cols["approval_for"] = idx
        elif n == "created":
            cols["created"] = idx
    return cols


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ""
    return (row[idx] or "").strip() if isinstance(row[idx], str) or row[idx] is None else str(row[idx]).strip()


def _extract_approval_datetime(row, cols):
    """Real approval date/time of an 'Approved' row.
    1) ServiceNow writes '<YYYY-MM-DD HH:MM:SS> - System (Comments)' in the
       Comments cell at the moment the approval is given -> use that.
    2) If there is no such comment, fall back to the 'Created' column.
    3) Last resort: any date-time found anywhere in the row.
    A CHG number / free text is NEVER returned."""
    comment = _cell(row, cols.get("comments"))
    m = APPROVAL_TS_IN_COMMENT.search(comment) or ANY_TS.search(comment)
    if m:
        return " ".join((m.group(1) if m.groups() else m.group(0)).split())

    created = _cell(row, cols.get("created"))
    m = ANY_TS.search(created)
    if m:
        return " ".join(m.group(0).split())

    m = ANY_TS.search(" ".join(str(c or "") for c in row))
    if m:
        return " ".join(m.group(0).split())
    return ""


def parse_approval_table(table):
    """Returns a list of dicts (approver, group, approved_on) for every
    row whose State is 'Approved'."""
    results = []
    if not table:
        return results
    header = table[0]
    cols = _find_approval_columns(header)
    # defaults = ServiceNow Approval List order
    cols.setdefault("state", 0)
    cols.setdefault("approver", 1)
    cols.setdefault("group", 2)
    cols.setdefault("comments", 3)
    cols.setdefault("approval_for", 4)
    cols.setdefault("created", 5)
    for row in table[1:]:
        if not row:
            continue
        if _cell(row, cols["state"]).lower() != "approved":
            continue
        results.append({
            "approver": " ".join(_cell(row, cols["approver"]).split()),
            "group": " ".join(_cell(row, cols["group"]).split()),
            "approved_on": _extract_approval_datetime(row, cols),
        })
    return results


def extract_data_from_pdf(file_bytes: bytes, filename: str):
    text = ""
    approvers = []
    approval_dates = []
    approver_details = []
    change_developer = ""
    change_implementor = ""
    change_implemented_on = ""
    cab_approval_provided = "False"
    cab_approver = None
    cab_approval_date = None

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for pdf_page in pdf.pages:
                for table in pdf_page.extract_tables():
                    for row in table:
                        if len(row) >= 10:
                            type_value = (
                                (" ".join((row[2] or "").split())).lower() if row[2] else None
                            )
                            assigned_to_value = (
                                " ".join((row[9] or "").split()) if row[9] else None
                            )
                            assigned_end_date = row[8].strip() if row[8] else None

                            if type_value and type_value.startswith("planning"):
                                change_developer = assigned_to_value
                            elif type_value and type_value.startswith("implement"):
                                change_implementor = assigned_to_value
                                change_implemented_on = assigned_end_date
    except Exception as e:
        log(f"    [Excel] pdfplumber failed to extract tables from {filename}: {e}")

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for pdf_page in pdf.pages:
                page_text = pdf_page.extract_text()
                if page_text:
                    text += page_text + "\n"

                for table in pdf_page.extract_tables():
                    if table and any(_norm_header(h) == "approver" for h in table[0]):
                        for rec in parse_approval_table(table):
                            approvers.append(rec["approver"])
                            approval_dates.append(rec["approved_on"])
                            approver_details.append(rec["group"])

                            if (not cab_approver) and "it cab approvers" in rec["group"].lower():
                                cab_approval_provided = "Yes"
                                cab_approver = rec["approver"]
                                cab_approval_date = rec["approved_on"]
    except Exception as e:
        log(f"    [Excel] pdfplumber failed to extract text from {filename}: {e}")

    ticket_no = extract_field_value(text, "Number:", length_limit=10)
    requestor = extract_requestor_value(text, "Requested by:")
    planned_start_date = extract_field_value(text, "Planned start date:", length_limit=19)
    change_type = extract_field_value(text, "Type:")
    change_description = extract_value_after_line(text, "Short description:")

    change_approved_by = ", ".join(approvers)
    change_approved_on = ", ".join(approval_dates)

    attribute_checks = {
        "Approvals Obtained": "Approval: Approved" in text,
        "Segregation of Duty": change_implementor.strip().lower() != change_developer.strip().lower(),
        "Exception Noted": "",
        "CAB Approval Provided (Yes/No)?": cab_approval_provided,
    }

    return (
        ticket_no,
        requestor,
        planned_start_date,
        change_type,
        change_description,
        change_approved_by,
        change_approved_on,
        change_developer,
        change_implementor,
        change_implemented_on,
        attribute_checks,
        cab_approver,
        cab_approval_date,
        cab_approval_provided,
    )


def generate_excel_report(data_list) -> bytes:
    df = pd.DataFrame(
        data_list,
        columns=[
            "Sl No", "Change Request Ticket", "Change Description", "Change Requestor",
            "Planned Start Date", "Change Type", "Change Approved By", "Change Approved On",
            "Approvals are obtained for all changes?",
            "Change Implemented By", "Change Implemented On", "Change Developed By",
            "Whether SOD is maintained between the developer and implementor? (Yes/No)",
            "CAB Approval Provided (Yes/No)?",
            "CAB Approver", "CAB Approval Obtained on", "Exception Noted",
        ],
    )

    df["Whether SOD is maintained between the developer and implementor? (Yes/No)"] = df.apply(
        lambda row: "True"
        if str(row["Change Implemented By"]).strip().lower() != str(row["Change Developed By"]).strip().lower()
        else "False",
        axis=1,
    )

    df["Exception Noted"] = df.apply(
        lambda row: "Yes" if row.astype(str).str.contains("False", case=False).any() else "No",
        axis=1,
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Testing Table"
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
    control_cell = ws.cell(row=1, column=1, value="Control Description: Testing & approval is required...")
    control_cell.font = Font(name="Source Sans Pro", size=12, bold=True)

    thin_border = Border(
        left=Side(style="thin", color="5B2A86"),
        right=Side(style="thin", color="5B2A86"),
        top=Side(style="thin", color="5B2A86"),
        bottom=Side(style="thin", color="5B2A86"),
    )

    attributes = [
        "Attribute 1", "Approvals are obtained for all changes",
        "Attribute 2", "Segregation of duty is ensured between developers and implementors",
        "Attribute 3", "CAB Approval Provided (Yes/No)?",
    ]
    for i in range(0, len(attributes), 2):
        attr_number = ws.cell(row=3 + i // 2, column=1, value=attributes[i])
        attr_number.font = Font(name="Source Sans Pro", bold=True, color="FFFFFF")
        attr_number.fill = PatternFill(start_color="5B2A86", end_color="5B2A86", fill_type="solid")
        attr_number.border = thin_border
        attr_description = ws.cell(row=3 + i // 2, column=2, value=attributes[i + 1])
        attr_description.border = thin_border
        attr_description.font = Font(name="Source Sans Pro")

    start_row = 8

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
        for c_idx, value in enumerate(row, 1):
            if r_idx > start_row and c_idx == 14:
                value = ""
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == start_row:
                cell.font = Font(bold=True, color="FFFFFF", name="Source Sans Pro")
                cell.fill = PatternFill(start_color="5B2A86", end_color="5B2A86", fill_type="solid")
            else:
                cell.font = Font(name="Source Sans Pro")
                cell.border = Border(
                    left=Side(style="thin", color="5B2A86"),
                    right=Side(style="thin", color="5B2A86"),
                    top=Side(style="thin", color="5B2A86"),
                    bottom=Side(style="thin", color="5B2A86"),
                )
                if c_idx == 14 and r_idx >= start_row + 1:
                    cell.value = f'=IF(O{r_idx}<>"", "Yes", "No")'

    for col in ws.iter_cols(min_row=3, max_row=start_row):
        col_values = [len(str(cell.value)) for cell in col if cell.value is not None]
        if col_values:
            ws.column_dimensions[col[0].column_letter].width = max(col_values) + 2

    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    ws.protection.sheet = True
    ws.protection.enable()

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_excel_report_for_tickets(download_dir: str, tickets, ticket_type_label: str):
    """Reads back each ticket's already-downloaded PDF (from its per-ticket
    folder) and runs it through the same extraction/report logic as the
    standalone Change Management Excel tool -- so the report is produced
    automatically in the SAME run, with no separate upload step. Returns
    the .xlsx bytes, or None if no ticket PDFs could be read."""
    data_list = []
    for i, ticket in enumerate(tickets, start=1):
        ticket_folder = get_ticket_folder(download_dir, ticket)
        pdf_files = sorted(glob.glob(os.path.join(ticket_folder, "*.pdf")))
        if not pdf_files:
            log(f"  [Excel] No PDF found for {ticket}, skipping in Excel report")
            continue
        try:
            with open(pdf_files[0], "rb") as f:
                file_bytes = f.read()
            (
                ticket_no, requestor, planned_start_date, change_type,
                change_description, change_approved_by, change_approved_on,
                change_developer, change_implementor, change_implemented_on,
                attribute_checks, cab_approver, cab_approval_date, cab_approval_provided,
            ) = extract_data_from_pdf(file_bytes, os.path.basename(pdf_files[0]))

            # Fall back to the originally-searched ticket number if the PDF
            # text extraction didn't find one (keeps the report row usable
            # even if the ticket's PDF layout differs slightly).
            ticket_no = ticket_no or ticket

            data_list.append(
                [
                    i, ticket_no, change_description, requestor, planned_start_date,
                    change_type, change_approved_by, change_approved_on,
                    attribute_checks["Approvals Obtained"],
                    change_implementor, change_implemented_on, change_developer,
                    attribute_checks["Segregation of Duty"], cab_approval_provided,
                    cab_approver, cab_approval_date, "",
                ]
            )
            log(f"  [Excel] Extracted report data for {ticket}")
        except Exception as e:
            log(f"  [Excel] Failed to process {ticket}: {type(e).__name__}: {e}")

    if not data_list:
        log("  [Excel] No valid ticket data to generate report.")
        return None

    return generate_excel_report(data_list)


def make_zip(download_dir: str, excel_bytes: bytes = None) -> str:
    """Zips the entire per-ticket folder structure, PLUS the Excel
    change-management report at the top level (if available):
    tickets.zip
      change_management_report.xlsx
      CHG0096025/
        CHG0096025_Normal_Ticket.pdf
        attachments/
          CHG0096025_COO File.zip
          ...
      CHG0070418/
        CHG0070418_Normal_Ticket.pdf
        (no attachments/ subfolder if this ticket had none)
    """
    zip_path = os.path.join(download_dir, "tickets.zip")
    with zipfile.ZipFile(zip_path, "w") as zf:
        for ticket_folder in sorted(glob.glob(os.path.join(download_dir, "*"))):
            if not os.path.isdir(ticket_folder):
                continue
            for root, _dirs, files in os.walk(ticket_folder):
                for fname in files:
                    if fname == "tickets.zip":
                        continue
                    full_path = os.path.join(root, fname)
                    arcname = os.path.relpath(full_path, download_dir)
                    zf.write(full_path, arcname)
        if excel_bytes:
            zf.writestr("change_management_report.xlsx", excel_bytes)
    return zip_path


# ==========================================================
# UI
# ==========================================================
st.markdown(
    """
    <div style="
        background-color:#5B2A86;
        border:2px solid #5B2A86;
        border-radius:10px;
        padding:22px 28px;
        margin-bottom:18px;
        display:flex;
        align-items:center;
        gap:16px;
    ">
        <div style="
            width:48px;height:48px;border-radius:50%;
            background-color:#FFFFFF; border:2px solid #5B2A86;
            display:flex;align-items:center;justify-content:center;
            font-weight:800;font-size:20px;color:#5B2A86;
        ">GE</div>
        <div>
            <div style="color:#E8DAF5;font-size:12px;letter-spacing:1px;text-transform:uppercase;">
                Change Management &middot; IT SOx Team Tool
            </div>
            <div style="color:#FFFFFF;font-size:26px;font-weight:800;">
                CM Automation
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

#st.caption(f"Detailed logs are also saved to: `{LOG_FILE_PATH}`")

step = st.session_state.step

if step == 0:
    st.header("Login")
    st.session_state.sso_id = st.text_input("SSO ID", value=st.session_state.sso_id)
    st.session_state.password = st.text_input("Password", type="password", value=st.session_state.password)

    st.info(
        "Your SSO ID and password will be entered automatically. After that, "
        "you'll just need to approve the push notification on your authenticator "
        "app (e.g. Okta Verify / Duo / Microsoft Authenticator) on your phone — "
        "no browser interaction needed."
    )

    st.session_state.headless = st.checkbox(
        "Run headless (leave checked for cloud deployment)",
        value=st.session_state.headless,
    )

    if st.button("Next"):
        if st.session_state.sso_id and st.session_state.password:
            st.session_state.step = 1
            st.rerun()
        else:
            st.warning("Please enter your SSO ID and password.")

elif step == 1:
    labels = TICKET_TYPE_LABELS[st.session_state.platform]
    label_values = list(labels.values())
    default_label = labels[st.session_state.ticket_type_key] if st.session_state.ticket_type_key in labels else label_values[0]
    choice_label = default_label
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

    n = len(st.session_state.tickets)
    est = 60 + n * 150
    m, s = divmod(est, 60)
    st.info(f"Estimated time (excluding manual MFA approval): up to {m} min {s} sec")

    st.markdown("**Ticket List**")
    _ticket_rows = "".join(
        f'<div style="padding:6px 12px;border-bottom:1px solid #E5D6F2;'
        f'display:flex;align-items:center;gap:10px;">'
        f'<span style="color:#5B2A86;font-weight:700;font-size:13px;">{i+1}.</span>'
        f'<span style="color:#2B2B2B;font-family:monospace;font-size:14px;">{t}</span>'
        f'</div>'
        for i, t in enumerate(st.session_state.tickets)
    )
    st.markdown(
        "<div style=\"background-color:#FFFFFF;border:1.5px solid #5B2A86;"
        "border-radius:8px;overflow:hidden;margin-bottom:14px;\">"
        + _ticket_rows +
        "</div>",
        unsafe_allow_html=True,
    )

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
        render_process_log(_log_placeholder, st.session_state.logs)

    if start_clicked:
        tmp_dir = tempfile.mkdtemp()
        st.session_state.download_dir = tmp_dir
        st.session_state.logs = []
        st.session_state.excel_bytes = None
        run_automation(
            st.session_state.sso_id,
            st.session_state.password,
            st.session_state.tickets,
            ticket_type_label,
            tmp_dir,
            st.session_state.headless,
        )
        st.session_state.password = ""

        log("Generating Excel change-management report from downloaded ticket PDFs...")
        with st.spinner("Generating Excel report..."):
            try:
                st.session_state.excel_bytes = build_excel_report_for_tickets(
                    tmp_dir, st.session_state.tickets, ticket_type_label
                )
                if st.session_state.excel_bytes:
                    log("Excel report generated successfully: change_management_report.xlsx")
                else:
                    log("Excel report could not be generated (no ticket PDFs were readable).")
            except Exception as ex:
                log(f"Excel report generation failed: {type(ex).__name__}: {ex}")
                st.session_state.excel_bytes = None

    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
            log_content = f.read()
        st.download_button("Download full diagnostic log (automation_log.txt)",
                            log_content, file_name="automation_log.txt")

    if st.session_state.excel_bytes:
        st.write("### 📊 Change Management Excel Report")
        st.download_button(
            "⬇ Download Excel Report (change_management_report.xlsx)",
            st.session_state.excel_bytes,
            file_name="change_management_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_excel_report",
        )

    if st.session_state.download_dir:
        ticket_folders = sorted([
            d for d in glob.glob(os.path.join(st.session_state.download_dir, "*"))
            if os.path.isdir(d)
        ])

        if ticket_folders:
            st.write("### Downloaded Tickets")
            st.write(f"{len(ticket_folders)} ticket folder(s) ready "
                     f"(each contains the ticket PDF and its attachments, kept separate):")

            for ticket_folder in ticket_folders:
                ticket_name = os.path.basename(ticket_folder)
                pdf_files = sorted(glob.glob(os.path.join(ticket_folder, "*.pdf")))
                attachments_dir = os.path.join(ticket_folder, "attachments")
                attachment_files = sorted(glob.glob(os.path.join(attachments_dir, "*"))) \
                    if os.path.isdir(attachments_dir) else []

                with st.expander(
                    f"📁 {ticket_name}  "
                    f"({len(pdf_files)} PDF, {len(attachment_files)} attachment(s))",
                    expanded=False,
                ):
                    if pdf_files:
                        st.caption("Ticket PDF")
                        for pdf_path in pdf_files:
                            fname = os.path.basename(pdf_path)
                            fsize_kb = os.path.getsize(pdf_path) / 1024
                            with open(pdf_path, "rb") as f:
                                st.download_button(
                                    f"⬇ {fname} ({fsize_kb:.0f} KB)",
                                    f,
                                    file_name=fname,
                                    mime="application/pdf",
                                    key=f"dl_pdf_{ticket_name}_{fname}",
                                )
                    else:
                        st.caption("No PDF found for this ticket.")

                    if attachment_files:
                        st.caption("Attachments")
                        for att_path in attachment_files:
                            fname = os.path.basename(att_path)
                            fsize_kb = os.path.getsize(att_path) / 1024
                            with open(att_path, "rb") as f:
                                st.download_button(
                                    f"⬇ {fname} ({fsize_kb:.0f} KB)",
                                    f,
                                    file_name=fname,
                                    key=f"dl_att_{ticket_name}_{fname}",
                                )
                    else:
                        st.caption("No attachments for this ticket.")

            zip_path = make_zip(st.session_state.download_dir, st.session_state.excel_bytes)
            zip_label = (
                "⬇ Download ALL tickets + attachments + Excel report as ZIP "
                "(organized in per-ticket folders)"
                if st.session_state.excel_bytes else
                "⬇ Download ALL tickets + attachments as ZIP "
                "(organized in per-ticket folders)"
            )
            with open(zip_path, "rb") as f:
                st.download_button(
                    zip_label,
                    f,
                    file_name="tickets.zip",
                    mime="application/zip",
                    key="dl_zip_all",
                )
        else:
            st.info("No tickets downloaded yet in this session.")
