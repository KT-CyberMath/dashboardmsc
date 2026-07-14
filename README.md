# Dashboard MSC

An internal desktop dashboard built for MSC, put together after the company's COO asked for a single, simple tool to replace the usual scatter of spreadsheets, email threads, and sticky notes used to run day-to-day operations.

## Why it exists

Day-to-day office work tends to live in too many places at once supplier quotes buried in email, meeting notes on paper, stock prices in a dozen different spreadsheet versions. The idea behind this dashboard is to pull the essentials into one lightweight tool that anyone on the team can open and use without training: track tasks straight from incoming emails, keep supplier contacts and price quotes in one place, log team members and meetings, and import/compare stock pricing all backed by a local database so nothing depends on a shared server or someone's personal spreadsheet copy.

It's intentionally simple by design: a native desktop app, no login, no cloud dependency, built to be extended as the team's needs grow.

## Features

- **Tasks** — an email inbox pane next to a to-do list; drag an email's subject or body straight into your tasks.
- **Suppliers** — contact details, attachments (contracts, emails, price quotes), and a price-quote log per supplier with Min/Max/Average pricing, sortable alphabetically or by registration date.
- **Members** — a simple team roster with role and notes.
- **Meetings** — log meetings by date with notes, and browse by day.
- **Stocklist** — import an Excel stocklist, auto-detect the group/price/name columns, and see per-group pricing stats.

## Requirements

- Python 3.14 or later, with a Tk 9.x build (older Tk versions, especially on macOS, have known rendering bugs buttons/entries can render blank or with wrong colors)
- `pandas` (for the Stocklist Excel import): `pip install pandas`
- `tkinterdnd2` (optional — enables dragging an email file from Finder/Explorer/Mail/Outlook straight onto the app; without it, the app still runs fine, you just use the "Browse" buttons instead): `pip install tkinterdnd2`
- `Pillow` (optional — shows the MSC logo in the dashboard header, resized from `logo.png`; without it, the header falls back to a plain text label instead): `pip install Pillow`

## Running it

```
python3 dashboard.py
```

Data is stored locally in `dashboard.db` (SQLite), created automatically the first time the app runs.

`logo.png` needs to sit next to `dashboard.py` for the header logo to show up — it's a committed asset, not generated at runtime.

## Making a double-clickable desktop app (macOS)

Instead of running `python3 dashboard.py` from Terminal every time, you can build a small `.app` wrapper that launches the dashboard from a double-click, like any other Mac app — using your project's own `.venv`, so it always runs with the same setup as Terminal would. It's personal to your machine (it hardcodes your own path), so it's not committed to git — build it once locally:

```
mkdir -p "Dashboard MSC.app/Contents/MacOS" "Dashboard MSC.app/Contents/Resources"
```

Then create `Dashboard MSC.app/Contents/Info.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key><string>Dashboard MSC</string>
    <key>CFBundleDisplayName</key><string>Dashboard MSC</string>
    <key>CFBundleIdentifier</key><string>com.msc.dashboard</string>
    <key>CFBundleVersion</key><string>1.0</string>
    <key>CFBundleShortVersionString</key><string>1.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>CFBundleExecutable</key><string>DashboardMSC</string>
    <key>CFBundleIconFile</key><string>icon.icns</string>
    <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
```

And `Dashboard MSC.app/Contents/MacOS/DashboardMSC` (then `chmod +x` it):

```bash
#!/bin/bash
APP_DIR="/absolute/path/to/dashboardmsc"
"$APP_DIR/.venv/bin/python3" "$APP_DIR/dashboard.py" 2> "$APP_DIR/launch_error.log"
```

Drop an `icon.icns` into `Dashboard MSC.app/Contents/Resources/` (you can convert `logo.png` to `.icns` with Preview.app, or `sips`/`iconutil` on the command line) for a proper app icon.

Drag the finished `.app` to your Desktop or `/Applications`. First launch will likely trigger a macOS Gatekeeper warning since it's unsigned — right-click it and choose "Open" once to allow it. If it doesn't launch, check `launch_error.log` next to `dashboard.py`.

## Notes for contributors

- `tk.Button` on macOS ignores custom colors due to a long-standing platform limitation in Tk's Aqua theme — this project uses a small `make_button()` helper (styled `Label` widgets) instead, so buttons look the same on macOS, Windows, and Linux. Please use `make_button()` rather than `tk.Button` for anything new.
- `dashboard.db` is local runtime data and differs machine to machine — avoid relying on its committed state for anything important.
