# Dashboard MSC

An internal desktop dashboard built for MSC, put together after the company's COO asked for a single, simple tool to replace the usual scatter of spreadsheets, email threads, and sticky notes used to run day-to-day operations.

## Why it exists

Day-to-day office work tends to live in too many places at once — supplier quotes buried in email, meeting notes on paper, stock prices in a dozen different spreadsheet versions. The idea behind this dashboard is to pull the essentials into one lightweight tool that anyone on the team can open and use without training: track tasks straight from incoming emails, keep supplier contacts and price quotes in one place, log team members and meetings, and import/compare stock pricing — all backed by a local database so nothing depends on a shared server or someone's personal spreadsheet copy.

It's intentionally simple by design: a native desktop app, no login, no cloud dependency, built to be extended as the team's needs grow.

## Features

- **Tasks** — an email inbox pane next to a to-do list; drag an email's subject or body straight into your tasks.
- **Suppliers** — contact details, attachments (contracts, emails, price quotes), and a price-quote log per supplier with Min/Max/Average pricing, sortable alphabetically or by registration date.
- **Members** — a simple team roster with role and notes.
- **Meetings** — log meetings by date with notes, and browse by day.
- **Stocklist** — import an Excel stocklist, auto-detect the group/price/name columns, and see per-group pricing stats.

## Requirements

- Python 3.14 or later, with a Tk 9.x build (older Tk versions, especially on macOS, have known rendering bugs — buttons/entries can render blank or with wrong colors)
- `pandas` (for the Stocklist Excel import): `pip install pandas`

## Running it

```
python3 dashboard.py
```

Data is stored locally in `dashboard.db` (SQLite), created automatically the first time the app runs.

## Notes for contributors

- `tk.Button` on macOS ignores custom colors due to a long-standing platform limitation in Tk's Aqua theme — this project uses a small `make_button()` helper (styled `Label` widgets) instead, so buttons look the same on macOS, Windows, and Linux. Please use `make_button()` rather than `tk.Button` for anything new.
- `dashboard.db` is local runtime data and differs machine to machine — avoid relying on its committed state for anything important.
