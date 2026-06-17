# Regal Rexnord — General Specifications PDF Downloader

A Streamlit app that takes an Excel file of Regal Rexnord product URLs, opens
each page in a stealth-patched headless browser, finds the **General
Specifications** section, and saves the page as a PDF — with deduplication,
a full audit trail, and automatic resume if the connection gets temporarily
rate-limited.

## Features

- Drag-and-drop Excel upload (auto-detects the URL column — `COnlineUrl`,
  `URL`, `Link`, or any column full of `http` links)
- Live progress bar, success/failure/duplicate counters, and a scrolling log
  while it runs
- Automatic resume — if you stop the app or it gets rate-limited, just click
  **Start** again later and it picks up exactly where it left off, skipping
  every URL already marked `Success`
- One-click downloads when finished: `Master_Final.xlsx`, the failed-URL
  report, or a ZIP of the entire output folder
- Adjustable pacing and cooldown behavior from the sidebar, no code editing
  required

## Requirements

- Python 3.10 or newer
- About 300 MB free disk space for the Chromium browser Playwright installs

## Setup

```bash
git clone <your-repo-url>
cd <your-repo-folder>
pip install -r requirements.txt
playwright install chromium
```

## Run

```bash
streamlit run app.py
```

This opens the app in your browser, usually at `http://localhost:8501`.

## Using the app

1. Upload an `.xlsx` file with a column of product URLs.
2. (Optional) Open **Settings** in the sidebar to adjust delay timing or how
   the app responds to blocking.
3. Click **Start Processing**. PDFs save as they're generated; the page
   stays on this screen with a live log and progress bar.
4. When it finishes (or pauses — see below), download `Master_Final.xlsx`,
   the failed-URL report, or the full output as a ZIP.

All output is also written directly to disk next to `app.py`:

```
Regal Rexnord Corporation/
├── PDFs/                       ← every generated PDF
├── Errors/Failed_Invalid_URLs.xlsx
├── Logs/Processing_Log.xlsx
└── Master_Final.xlsx           ← full audit trail (Row_ID, URL, status, etc.)
```

If you're running this on your own machine, the ZIP download is just a
convenience — the files are already sitting in that folder.

## How resume works

Every 5 URLs, the app saves its progress to `Master_Final.xlsx`. When you
click **Start** again — whether that's because the run finished partway,
the app crashed, or you closed the terminal — it reads that file first and
skips any URL already marked `Success`. PDF file numbering also continues
from the highest existing file, so nothing gets overwritten.

To wipe progress and start completely over, use the **Reset** section in
the sidebar.

## About rate limiting (please read before a large run)

Regal Rexnord's site has bot-detection that can return `HTTP 403` after a
handful of requests, even with realistic delays and a stealth-patched
browser. This app handles that as gracefully as it reasonably can:

- A short pause (default 60s) after 3 blocks in a row, then a session
  refresh
- A longer pause (default 5 minutes) after 6 in a row
- After 10 in a row, it stops on its own, saves everything, and tells you
  to wait before trying again

**There's no setting that guarantees this won't happen**, and that's
intentional — the techniques that exist to defeat this more aggressively
(rotating proxies, residential IP networks, CAPTCHA-solving services) are
the kind of thing this project deliberately doesn't include. They cross
from "a patient, realistic browser" into "designed specifically to evade a
site's protection at scale," which isn't something to build casually.

If you're regularly hitting the give-up threshold on a large batch (the
original ask was 30,000+ URLs), the most reliable path forward is slowing
the delay sliders down further, running it in smaller daily batches, or
reaching out to Regal Rexnord directly — distributors and large customers
can often get a bulk product-data export or API access, which is a more
sustainable way to pull thousands of spec sheets than browser automation.

### One technical note

`page.pdf()` in Playwright only works when Chromium runs in headless mode —
a visible browser window will not export a PDF at all. So unlike many
scraping guides, "show a real browser window" isn't an option here; the
app always runs headless, and relies instead on stealth patches, a
persistent cookie profile, realistic delays, and a `Referer` header to look
like ordinary traffic.

## Project structure

```
.
├── app.py              ← the Streamlit app
├── requirements.txt
├── README.md
└── .gitignore
```

## Adjustable settings (sidebar)

| Setting | What it does | Default |
|---|---|---|
| Delay between requests | Random pause before each URL | 8–15s |
| Short pause threshold / duration | Pause + session refresh after N blocks in a row | 3 blocks / 60s |
| Long pause threshold / duration | Longer pause after N blocks in a row | 6 blocks / 300s |
| Stop threshold | Give up and save after N blocks in a row | 10 blocks |

## License

Use and modify freely for your own internal purposes.
