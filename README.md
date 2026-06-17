# Regal Rexnord — Smart Specifications PDF Downloader

A Streamlit app that takes an Excel file of Regal Rexnord product URLs, opens
each page in a stealth-patched headless browser, finds the **General
Specifications** section, and saves the page as a PDF — with deduplication,
a full audit trail, and resume support.

## Files in this project

```
.
├── app.py              ← the Streamlit app
├── requirements.txt    ← pip dependencies (playwright is pinned — see below)
├── packages.txt         ← apt dependencies, only used by Streamlit Cloud
├── README.md
└── .gitignore
```

---

## Option A — Run locally

```bash
git clone <your-repo-url>
cd <your-repo-folder>
pip install -r requirements.txt
playwright install chromium
streamlit run app.py
```

This opens the app at `http://localhost:8501`. Output is written to
`Regal Rexnord Corporation/` next to `app.py`, and it persists there
normally — like any other file on your computer.

---

## Option B — Deploy to Streamlit Community Cloud

Push this repo to GitHub, then deploy it from
[share.streamlit.io](https://share.streamlit.io). Community Cloud reads
`requirements.txt` and `packages.txt` automatically — no extra setup needed
beyond having both files at the repo root, which they already are here.

### Why this needed two extra files

Locally, you run `playwright install chromium` yourself in a terminal.
Community Cloud has no terminal step between installing `requirements.txt`
and starting the app — so `app.py` does that step itself on first run
(cached, so it only happens once per container, not on every interaction).
That's the `ensure_chromium_installed()` function near the top of the file.

But downloading the browser isn't enough — Chromium also needs several
system-level libraries to actually *run*, which Python's package manager
can't install. That's what `packages.txt` is for; Community Cloud installs
everything listed there via `apt-get` during the build.

### Why `playwright` is pinned to `1.49.0`

Community Cloud's official docs state it runs on **Debian 11 ("bullseye")**
— that base image is from 2021. Newer Playwright releases bundle Chromium
builds that expect newer system libraries than Bullseye ships (and some
even renamed the libraries Bullseye uses, e.g. `libasound2` became
`libasound2t64` on more recent Debian/Ubuntu — so a `packages.txt`
generated on a newer machine will list packages that don't exist on
Bullseye and silently fail to install). `1.49.0` is a version confirmed
to work against Bullseye's libraries. If Streamlit eventually updates
Community Cloud's base image, a newer pin may start working too — if you
want to try, bump the version in `requirements.txt`, watch the build logs
for apt errors, and adjust `packages.txt` accordingly.

### Important: Community Cloud's disk is temporary

This is the part that's easy to miss. Files written while the app is
running **do not necessarily survive**:

- a new deployment (`git push`)
- the app going to sleep after 12 hours with no visitors, then waking back up
- the app being rebooted from "Manage app"

For a run that finishes in one sitting, this doesn't matter. For a very
large batch spread across multiple sessions, it does — there's a real
chance of losing progress you thought was saved. The app shows a warning
above the download buttons for exactly this reason: **download your
results before closing the tab**, don't assume they'll still be there
tomorrow.

If you need true persistence across sessions (large batches, multi-day
runs), running locally — or self-hosting via Docker on something like
Render, Railway, Fly.io, or a small VPS, where you control the disk and
the OS version — is more reliable than Community Cloud for this specific
workload. Community Cloud also caps resources fairly low (around 1
CPU core, limited memory), and a real browser doing PDF rendering is
heavier than the typical Streamlit app it's designed for.

---

## Using the app

1. Upload an `.xlsx` file with a column of product URLs (`COnlineUrl`,
   `URL`, `Link`, or any column full of `http` links — auto-detected).
2. Review the preview and detected URL count.
3. (Optional) Adjust delay timing or blocking-response thresholds in the
   sidebar.
4. Click **Start Processing**. Watch the live log, progress bar, and
   counters while it runs.
5. Download `Master_Final.xlsx`, the failed-URL report, or a full ZIP when
   done — **immediately**, if you're on Community Cloud.

## How resume works

The app saves progress to `Master_Final.xlsx` periodically while running
(configurable via "Save progress every N processed URLs" in the sidebar).
Starting a new run reads that file first and skips any URL already marked
`Success`, continuing PDF numbering from the highest existing file. This
works within a single long-running local session reliably; on Community
Cloud it only works if the underlying container hasn't restarted since
your last run (see above).

## About rate limiting

Regal Rexnord's site has bot-detection that can return `HTTP 403` after a
handful of requests, even with a stealth-patched browser and realistic
delays. The app responds with a short pause and session refresh after a
few blocks in a row, a longer pause after more, and stops on its own
(saving everything first) past a configurable threshold — all adjustable
in the sidebar.

There's no setting that guarantees this won't happen, and that's
intentional: techniques that exist to defeat this more aggressively
(rotating proxies, residential IP networks, CAPTCHA-solving services)
aren't included here. If you're hitting the give-up threshold often on a
large batch, slow the delay sliders down further, run smaller batches over
multiple days, or contact Regal Rexnord directly — distributors and large
customers can often get a bulk product-data export, which is a more
sustainable route for thousands of spec sheets than browser automation.

One related technical note: `page.pdf()` in Playwright only works when
Chromium runs in **headless** mode — a visible browser window can't export
a PDF at all. So unlike a lot of general scraping advice, "show a real
window" isn't an available lever here; the app always runs headless and
relies on stealth patches, a persistent cookie profile, a `Referer` header,
and pacing instead.
