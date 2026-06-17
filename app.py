"""
Regal Rexnord — General Specifications PDF Downloader
========================================================
Streamlit app: upload an Excel file of product URLs, the app opens
each page with a stealth-patched headless Chromium browser, finds the
General Specifications section, and saves the page as a PDF.

Run with:
    streamlit run app.py

See README.md for full setup instructions and notes on anti-bot blocking.
"""

import io
import re
import sys
import random
import zipfile
import asyncio
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ── Windows needs the Proactor loop for asyncio subprocess support ──
# (Playwright launches Chromium as a subprocess; the Selector loop
#  policy does NOT support subprocesses on Windows.)
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

PLAYWRIGHT_OK = True
PLAYWRIGHT_IMPORT_ERROR = ""
try:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeout
    from playwright_stealth import Stealth
except Exception as e:  # noqa: BLE001
    PLAYWRIGHT_OK = False
    PLAYWRIGHT_IMPORT_ERROR = str(e)

# ════════════════════════════════════════════════════════════════
#  CONSTANTS
# ════════════════════════════════════════════════════════════════

OUT_BASE    = Path("Regal Rexnord Corporation")
PDF_DIR     = OUT_BASE / "PDFs"
ERR_DIR     = OUT_BASE / "Errors"
LOG_DIR     = OUT_BASE / "Logs"
MASTER_PATH = OUT_BASE / "Master_Final.xlsx"
PROFILE_DIR = OUT_BASE / ".browser_profile"   # persists cookies between runs

RUN_DATE      = datetime.now().strftime("%Y%m%d")
TODAY_DISPLAY = datetime.now().strftime("%Y-%m-%d")

HOMEPAGE = "https://www.regalrexnord.com/"

BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/137.0.0.0 Safari/537.36"
)

SPEC_SELECTORS = [
    'button:has-text("General Specifications")',
    'button:has-text("Specifications")',
    'a:has-text("General Specifications")',
    'a:has-text("Specifications")',
    '[role="tab"]:has-text("Specifications")',
    '[class*="tab"]:has-text("Specifications")',
    '[class*="accordion"]:has-text("Specifications")',
]

STATUS_COLORS = {
    "Success":   "C6EFCE",
    "Duplicate": "FFEB9C",
    "Failed":    "FFC7CE",
    "Pending":   "FFFFFF",
}
HEADERS = ["Row_ID", "URL", "Part_Number", "Duplicate_Flag",
           "Status", "PDF_File_Name", "Note"]

# ════════════════════════════════════════════════════════════════
#  PURE HELPERS
# ════════════════════════════════════════════════════════════════

def normalize_url(url: str) -> str:
    url = url.strip().lower()
    url = url.rstrip("/")
    url = re.sub(r"#.*$", "", url)
    return url

def part_from_url(url: str) -> str:
    slug = url.rstrip("/").split("/")[-1]
    m = re.search(r"(\d{4}-\d{5})$", slug)                  # bevel: 0270-09576
    if m:
        return m.group(1).upper()
    m = re.search(r"([A-Za-z]{1,4}\d{5,}-\d{2,})$", slug)   # EL8420267-00
    if m:
        return m.group(1).upper()
    # Fallback for bearings/accessories/bare numeric IDs, e.g. .../products/1234459,
    # .../...-lock-767774, .../shaftmount-accessory-xp9202 — prefer the final
    # segment alone when it already looks like a standalone identifier.
    parts = slug.split("-")
    last = parts[-1]
    if (last.isdigit() and len(last) >= 4) or re.match(r"^[A-Za-z0-9]{3,}$", last):
        return last.upper()
    return ("-".join(parts[-2:]) if len(parts) >= 2 else slug).upper()

def get_max_pdf_number() -> int:
    """Scan PDF_DIR for existing PdfFileNNNNNN_*.pdf to continue numbering on resume."""
    max_n = 0
    if PDF_DIR.exists():
        for f in PDF_DIR.glob("PdfFile*.pdf"):
            m = re.match(r"PdfFile(\d{6})_", f.name)
            if m:
                max_n = max(max_n, int(m.group(1)))
    return max_n

def detect_url_column(df: pd.DataFrame):
    for candidate in ["COnlineUrl", "URL", "Url", "url",
                       "Link", "link", "Product URL", "ProductURL"]:
        if candidate in df.columns:
            return candidate
    for col in df.columns:
        if any(str(v).startswith("http") for v in df[col].dropna().head(3)):
            return col
    return None

def build_zip_bytes(folder: Path) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in folder.rglob("*"):
            if path.is_file() and ".browser_profile" not in path.parts:
                zf.write(path, path.relative_to(folder.parent))
    buf.seek(0)
    return buf.getvalue()

# ════════════════════════════════════════════════════════════════
#  EXCEL OUTPUT
# ════════════════════════════════════════════════════════════════

def _header(ws, cols, color="003087"):
    fill = PatternFill("solid", fgColor=color)
    font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 20

def _autowidth(ws, min_w=10, max_w=80):
    for col in ws.columns:
        ltr = get_column_letter(col[0].column)
        ml = max((len(str(c.value)) for c in col if c.value), default=0)
        ws.column_dimensions[ltr].width = max(min_w, min(max_w, ml + 2))

def save_master(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    _header(ws, HEADERS)
    for r in rows:
        ws.append([r["row_id"], r["url"], r["part"], r["duplicate"],
                   r["status"], r.get("pdf_name", "NULL"), r.get("note", "")])
        ri = ws.max_row
        fc = STATUS_COLORS.get(r["status"], "FFFFFF")
        for ci in range(1, len(HEADERS) + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.fill = PatternFill("solid", fgColor=fc)
            cell.font = Font(name="Arial", size=9)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    _autowidth(ws)
    ws.column_dimensions["B"].width = 80
    wb.save(path)

def save_errors(rows, path):
    failed = [r for r in rows if r["status"] == "Failed"]
    if not failed:
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Failed URLs"
    _header(ws, ["Row_ID", "URL", "Error", "Timestamp"])
    for r in failed:
        ws.append([r["row_id"], r["url"], r.get("note", ""), TODAY_DISPLAY])
    _autowidth(ws)
    ws.column_dimensions["B"].width = 80
    wb.save(path)

def save_log(rows, pdf_n, start, end, path):
    total   = len(rows)
    unique  = sum(1 for r in rows if r["duplicate"] == "No")
    success = sum(1 for r in rows if r["status"] == "Success")
    failed  = sum(1 for r in rows if r["status"] == "Failed")
    pending = sum(1 for r in rows if r["status"] == "Pending")
    dur     = round((end - start).total_seconds(), 1)

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    _header(ws, ["Metric", "Value"])
    for k, v in [
        ("Total URLs",        total),
        ("Unique",            unique),
        ("Duplicates",        total - unique),
        ("Successful",        success),
        ("Failed",            failed),
        ("Not yet processed", pending),
        ("PDFs in folder",    pdf_n),
        ("Start Time",        start.strftime("%Y-%m-%d %H:%M:%S")),
        ("End Time",          end.strftime("%Y-%m-%d %H:%M:%S")),
        ("Duration (s)",      dur),
        ("Duration (min)",    round(dur / 60, 1)),
        ("Run Date",          TODAY_DISPLAY),
    ]:
        ws.append([k, v])
        ri = ws.max_row
        ws.cell(ri, 1).font = Font(name="Arial", bold=True, size=10)
        ws.cell(ri, 2).font = Font(name="Arial", size=10)
        ws.cell(ri, 2).alignment = Alignment(horizontal="right")
    _autowidth(ws)
    wb.save(path)

def load_previous_results(path):
    """Load a previous Master_Final.xlsx -> {normalized_url: {status, pdf_name, note}}"""
    result = {}
    try:
        wb = load_workbook(path)
        ws = wb["Results"] if "Results" in wb.sheetnames else wb.active
        headers = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(headers)}
        if not all(k in idx for k in ("URL", "Status", "PDF_File_Name", "Note")):
            return result
        for row in ws.iter_rows(min_row=2, values_only=True):
            url = row[idx["URL"]]
            if not url:
                continue
            result[normalize_url(str(url))] = {
                "status":   row[idx["Status"]],
                "pdf_name": row[idx["PDF_File_Name"]],
                "note":     row[idx["Note"]] or "",
            }
    except Exception:
        pass
    return result

# ════════════════════════════════════════════════════════════════
#  ASYNC CORE: open URL -> click Specs button -> print to PDF
# ════════════════════════════════════════════════════════════════

async def save_pdf(page, url: str, pdf_path: str):
    """
    Returns (result, note, http_status)
      result: "success" | "403" | "error"
    """
    try:
        response = await page.goto(
            url, wait_until="domcontentloaded", timeout=60_000, referer=HOMEPAGE
        )
        status = response.status if response else 0
    except PWTimeout:
        return "error", "Timeout", 0
    except Exception as e:  # noqa: BLE001
        return "error", str(e), 0

    if status == 403:
        return "403", "HTTP 403", 403
    if status >= 400:
        return "error", f"HTTP {status}", status

    await page.wait_for_timeout(3_000)   # let JS finish rendering

    clicked = False
    for sel in SPEC_SELECTORS:
        try:
            await page.click(sel, timeout=3_000)
            await page.wait_for_timeout(1_500)
            clicked = True
            break
        except Exception:  # noqa: BLE001
            pass

    try:
        await page.pdf(
            path=pdf_path,
            format="Letter",
            print_background=True,
            margin={"top": "0.5in", "bottom": "0.5in",
                    "left": "0.5in", "right": "0.5in"},
        )
        note = "specs button clicked" if clicked else "no specs button found"
        return "success", note, status
    except Exception as e:  # noqa: BLE001
        return "error", f"PDF print failed: {e}", status

async def warmup(page, log):
    """Visit the homepage first so the session looks like a normal visitor."""
    log("Warming up session...")
    try:
        await page.goto(HOMEPAGE, wait_until="domcontentloaded", timeout=30_000)
        await page.wait_for_timeout(3_000)
        log("Session ready.")
    except Exception as e:  # noqa: BLE001
        log(f"Warmup warning: {e}")

# ════════════════════════════════════════════════════════════════
#  MAIN ASYNC PROCESSOR  (drives the Streamlit UI as it runs)
# ════════════════════════════════════════════════════════════════

async def run_processing(url_list, settings, ui):
    """
    settings: dict(delay_min, delay_max, soft_threshold, soft_seconds,
                    hard_threshold, hard_seconds, giveup_threshold)
    ui:       dict of callables/placeholders used to push live updates:
                  ui['log'](msg)            -> append a line to the log box
                  ui['progress'](fraction)  -> update the progress bar
                  ui['metrics'](success, failed, dupes, remaining) -> update counters
    """
    start = datetime.now()

    for d in [PDF_DIR, ERR_DIR, LOG_DIR]:
        d.mkdir(parents=True, exist_ok=True)

    # ── Deduplicate ─────────────────────────────────────────
    seen, rows = {}, []
    for idx, raw in enumerate(url_list, 1):
        url  = str(raw).strip()
        norm = normalize_url(url)
        part = part_from_url(url)
        if norm in seen:
            rows.append(dict(row_id=idx, url=url, norm=norm, part=part,
                              duplicate="Yes", status="Duplicate",
                              pdf_name="NULL", note=""))
        else:
            seen[norm] = idx
            rows.append(dict(row_id=idx, url=url, norm=norm, part=part,
                              duplicate="No", status="Pending",
                              pdf_name="", note=""))

    total  = len(rows)
    unique = sum(1 for r in rows if r["duplicate"] == "No")
    dupes  = total - unique

    # ── Resume from a previous run, if Master_Final.xlsx exists ──
    resumed = 0
    if MASTER_PATH.exists():
        previous = load_previous_results(MASTER_PATH)
        for r in rows:
            if r["duplicate"] == "No" and r["norm"] in previous:
                prev = previous[r["norm"]]
                if prev["status"] == "Success":
                    r.update(status="Success", pdf_name=prev["pdf_name"], note=prev["note"])
                    resumed += 1
        if resumed:
            ui["log"](f"Found previous progress — {resumed} URL(s) already "
                       f"completed and will be skipped.")

    pdf_n = get_max_pdf_number()
    remaining = sum(1 for r in rows if r["duplicate"] == "No" and r["status"] == "Pending")

    ui["log"](f"Total: {total}  |  Unique: {unique}  |  Duplicates: {dupes}  "
              f"|  Already done: {resumed}  |  Remaining: {remaining}")
    ui["metrics"](resumed, 0, dupes, remaining)

    if remaining == 0:
        ui["log"]("Nothing left to process — all URLs already completed.")
        end = datetime.now()
        save_master(rows, MASTER_PATH)
        save_errors(rows, ERR_DIR / "Failed_Invalid_URLs.xlsx")
        save_log(rows, pdf_n, start, end, LOG_DIR / "Processing_Log.xlsx")
        return dict(rows=rows, pdf_n=pdf_n, dupes=dupes, gave_up=False,
                    start=start, end=end)

    consecutive_403 = 0
    gave_up = False
    processed_count = 0

    async with async_playwright() as pw:
        context = await pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=True,   # required: Playwright/Chromium only supports
                              # page.pdf() in headless mode (non-headless
                              # silently fails PDF generation).
            user_agent=BROWSER_UA,
            locale="en-US",
            viewport={"width": 1280, "height": 900},
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
            ],
        )
        page = context.pages[0] if context.pages else await context.new_page()

        await Stealth().apply_stealth_async(page)
        await context.add_init_script(
            "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"
        )

        await warmup(page, ui["log"])

        try:
            for row in rows:
                idx = row["row_id"]
                url = row["url"]

                if row["duplicate"] == "Yes":
                    continue
                if row["status"] == "Success":
                    continue
                if not re.match(r"https?://", url):
                    row.update(status="Failed", note="Invalid URL")
                    ui["log"](f"[{idx}] INVALID URL")
                    continue

                pdf_n += 1
                pdf_name = f"PdfFile{pdf_n:06d}_{RUN_DATE}_{row['part']}.pdf"
                pdf_path = str(PDF_DIR / pdf_name)

                result, note, _status_code = await save_pdf(page, url, pdf_path)

                if result == "success":
                    row.update(status="Success", pdf_name=pdf_name, note=note)
                    consecutive_403 = 0
                    ui["log"](f"[{idx}] OK -> {pdf_name}  ({note})")
                elif result == "403":
                    pdf_n -= 1
                    consecutive_403 += 1
                    row.update(status="Failed",
                               note=f"HTTP 403 (consecutive #{consecutive_403})")
                    ui["log"](f"[{idx}] BLOCKED (HTTP 403) — {consecutive_403} in a row")
                else:
                    pdf_n -= 1
                    consecutive_403 = 0
                    row.update(status="Failed", note=note)
                    ui["log"](f"[{idx}] FAILED — {note}")

                processed_count += 1

                # Incremental save — progress survives crashes / interruptions
                if processed_count % 5 == 0:
                    save_master(rows, MASTER_PATH)

                success_n = sum(1 for r in rows if r["status"] == "Success")
                failed_n  = sum(1 for r in rows if r["status"] == "Failed")
                remaining_n = sum(1 for r in rows
                                  if r["duplicate"] == "No" and r["status"] == "Pending")
                ui["metrics"](success_n, failed_n, dupes, remaining_n)
                ui["progress"]((unique - remaining_n) / unique if unique else 1.0)

                # Escalating response to repeated blocks
                if consecutive_403 == settings["soft_threshold"]:
                    ui["log"](f"{consecutive_403} blocks in a row — pausing "
                              f"{settings['soft_seconds']}s and refreshing session...")
                    await asyncio.sleep(settings["soft_seconds"])
                    await warmup(page, ui["log"])
                elif consecutive_403 == settings["hard_threshold"]:
                    mins = settings["hard_seconds"] // 60
                    ui["log"](f"{consecutive_403} blocks in a row — pausing "
                              f"{mins} minutes and refreshing session...")
                    await asyncio.sleep(settings["hard_seconds"])
                    await warmup(page, ui["log"])
                elif consecutive_403 >= settings["giveup_threshold"]:
                    ui["log"](f"{consecutive_403} blocks in a row — the site is "
                              f"rate-limiting this connection. Progress saved.")
                    ui["log"]("Wait 15-30 minutes, then click Start again — "
                              "it will resume automatically.")
                    gave_up = True
                    break

                await asyncio.sleep(random.uniform(settings["delay_min"], settings["delay_max"]))

        finally:
            await context.close()

    end = datetime.now()

    save_master(rows, MASTER_PATH)
    save_errors(rows, ERR_DIR / "Failed_Invalid_URLs.xlsx")
    save_log(rows, pdf_n, start, end, LOG_DIR / "Processing_Log.xlsx")

    return dict(rows=rows, pdf_n=pdf_n, dupes=dupes, gave_up=gave_up,
                start=start, end=end)

def run_async(coro):
    """Run an async coroutine from Streamlit's synchronous script context."""
    try:
        return asyncio.run(coro)
    except RuntimeError as e:
        if "cannot be called from a running event loop" in str(e):
            import nest_asyncio
            nest_asyncio.apply()
            loop = asyncio.get_event_loop()
            return loop.run_until_complete(coro)
        raise

# ════════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Regal Rexnord PDF Downloader", page_icon="🏭", layout="wide")

st.title("🏭 Regal Rexnord — Specifications PDF Downloader")
st.caption(
    "Upload an Excel file of product URLs. The app opens each page in a "
    "stealth-patched headless browser, finds the General Specifications "
    "section, and saves it as a PDF."
)

if not PLAYWRIGHT_OK:
    st.error(
        "Playwright isn't set up yet.\n\n"
        "Run these two commands in your terminal, then restart this app:\n\n"
        "```\npip install -r requirements.txt\nplaywright install chromium\n```\n\n"
        f"Import error: {PLAYWRIGHT_IMPORT_ERROR}"
    )
    st.stop()

with st.sidebar:
    st.header("⚙️ Settings")
    st.caption(
        "PDF export only works in headless mode (a Playwright/Chromium "
        "limitation), so the browser always runs in the background."
    )
    delay_min, delay_max = st.slider(
        "Delay between requests (seconds)", 3, 30, (8, 15),
        help="Longer delays look more human and reduce the chance of being "
             "rate-limited, at the cost of a slower run."
    )
    st.divider()
    st.subheader("Response to blocking")
    soft_threshold = st.number_input("Short pause after N blocks in a row", 1, 10, 3)
    soft_seconds   = st.number_input("Short pause duration (sec)", 10, 600, 60)
    hard_threshold = st.number_input("Long pause after N blocks in a row", 1, 20, 6)
    hard_seconds   = st.number_input("Long pause duration (sec)", 60, 1800, 300)
    giveup_threshold = st.number_input("Stop & save after N blocks in a row", 5, 50, 10)

    st.divider()
    st.subheader("🗑️ Reset")
    if MASTER_PATH.exists():
        if st.button("Clear progress tracking (keep PDFs)"):
            MASTER_PATH.unlink(missing_ok=True)
            st.success("Progress tracking cleared. PDFs were kept.")
            st.rerun()
        confirm_wipe = st.checkbox("I understand — delete ALL output files")
        if st.button("Delete everything and start fresh", disabled=not confirm_wipe):
            import shutil
            shutil.rmtree(OUT_BASE, ignore_errors=True)
            st.success("All output deleted.")
            st.rerun()
    else:
        st.caption("Nothing to reset yet.")

# ── Resume banner ──────────────────────────────────────────────
if MASTER_PATH.exists():
    prev = load_previous_results(MASTER_PATH)
    completed = sum(1 for v in prev.values() if v["status"] == "Success")
    if completed:
        st.info(f"📁 Found previous progress: **{completed}** URL(s) already "
                f"completed. They'll be skipped automatically when you start.")

# ── File upload ───────────────────────────────────────────────
uploaded_file = st.file_uploader("Select Excel file (.xlsx)", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype=str)
    df.columns = df.columns.str.strip()
    url_col = detect_url_column(df)

    if not url_col:
        st.error(f"No URL column found. Columns in file: {list(df.columns)}")
    else:
        url_list = [u for u in df[url_col].dropna() if str(u).strip()]
        st.success(f'Column "{url_col}" detected — **{len(url_list)}** URLs loaded.')

        start_clicked = st.button("🚀 Start Processing", type="primary", use_container_width=True)

        if start_clicked:
            status_box = st.empty()
            progress_bar = st.progress(0.0)
            c1, c2, c3, c4 = st.columns(4)
            m_success   = c1.empty()
            m_failed    = c2.empty()
            m_dupes     = c3.empty()
            m_remaining = c4.empty()

            log_lines = []
            log_container = st.expander("📋 Live log", expanded=True)
            log_placeholder = log_container.empty()

            def log(msg: str):
                log_lines.append(msg)
                log_placeholder.code("\n".join(log_lines[-300:]), language=None)

            def metrics(success, failed, dupes, remaining):
                m_success.metric("✅ Success", success)
                m_failed.metric("❌ Failed", failed)
                m_dupes.metric("⚠️ Duplicates", dupes)
                m_remaining.metric("⏳ Remaining", remaining)

            def progress(fraction):
                progress_bar.progress(min(max(fraction, 0.0), 1.0))

            ui = dict(log=log, metrics=metrics, progress=progress)
            settings = dict(
                delay_min=delay_min, delay_max=delay_max,
                soft_threshold=soft_threshold, soft_seconds=soft_seconds,
                hard_threshold=hard_threshold, hard_seconds=hard_seconds,
                giveup_threshold=giveup_threshold,
            )

            status_box.info("Processing started — this window updates live below.")
            result = run_async(run_processing(url_list, settings, ui))

            dur = round((result["end"] - result["start"]).total_seconds(), 1)
            if result["gave_up"]:
                status_box.warning(
                    f"⏸️ Paused after repeated blocks ({dur}s elapsed). "
                    f"Progress was saved — wait 15-30 minutes and click "
                    f"Start again to resume."
                )
            else:
                status_box.success(f"✅ Complete! ({dur}s elapsed)")
            st.rerun()

# ── Download section ────────────────────────────────────────────
if MASTER_PATH.exists():
    st.divider()
    st.subheader("📥 Download Results")

    col1, col2, col3 = st.columns(3)

    with col1:
        with open(MASTER_PATH, "rb") as f:
            st.download_button(
                "📄 Master_Final.xlsx", f,
                file_name="Master_Final.xlsx",
                use_container_width=True,
            )

    err_path = ERR_DIR / "Failed_Invalid_URLs.xlsx"
    with col2:
        if err_path.exists():
            with open(err_path, "rb") as f:
                st.download_button(
                    "⚠️ Failed_Invalid_URLs.xlsx", f,
                    file_name="Failed_Invalid_URLs.xlsx",
                    use_container_width=True,
                )
        else:
            st.button("⚠️ No failures yet", disabled=True, use_container_width=True)

    with col3:
        if st.button("📦 Build full ZIP", use_container_width=True):
            with st.spinner("Zipping PDFs, Excel files, and logs..."):
                zip_bytes = build_zip_bytes(OUT_BASE)
            st.download_button(
                "⬇️ Download ZIP", zip_bytes,
                file_name=f"Regal_Rexnord_Output_{RUN_DATE}.zip",
                mime="application/zip",
                use_container_width=True,
            )

    n_pdfs = len(list(PDF_DIR.glob("*.pdf"))) if PDF_DIR.exists() else 0
    st.caption(
        f"PDFs are also saved directly on disk at: `{PDF_DIR.resolve()}` "
        f"({n_pdfs} files) — no need to download the ZIP if you're running "
        f"this app on your own computer."
    )
