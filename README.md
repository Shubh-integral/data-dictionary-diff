# Data Dictionary Diff

Browser-based tool for comparing a baseline data dictionary with a new one. Open the [live app](https://shubh-integral.github.io/data-dictionary-diff/), select your spreadsheets, and compare their tables and columns.

Spreadsheet processing happens locally in your browser. SheetJS 0.20.3 and the fonts are bundled in this repository; the app makes no third-party requests. A Content Security Policy blocks network API calls, remote resources, inline scripts, and form submissions. Local scripts and styles have SHA-384 integrity checks. These controls reduce risk; they are not a guarantee against every possible browser, dependency, or hosting compromise.

The table/column matching, source scoping, baseline markup, and Excel export rules are unchanged. Clipboard TSV output quotes fields, prefixes formula-like values with an apostrophe, replaces control characters (including tabs and line breaks) with spaces, and removes invisible direction/format controls. This is intended for pasting into spreadsheets; use the Excel download to preserve original strings and multiline sample values. Spreadsheet applications differ in how they interpret clipboard text; preserve text formatting and do not remove the protective apostrophes when re-exporting clipboard data.

Imports run in a cancellable worker with a 30-second timeout. The blob worker inherits the page's CSP, including its network restrictions. Safety limits are 5 files, 20 MiB per file, 50 MiB combined, 50 tabs per workbook, 250,000 data rows per workbook, about 500,000 rows combined, 256 columns per tab, 2 million cells per workbook, and 100 MiB declared expanded ZIP content. Oversized files are rejected with an explanation, rather than silently truncated. These limits mitigate resource exhaustion; browsers do not provide a strict per-worker memory cap.

## Local development and validation

Serve this directory over HTTP, for example with `python3 -m http.server 8000 --bind 127.0.0.1`. The app uses a worker, so opening `index.html` through `file://` is not supported.

After changing a script or stylesheet, run `node scripts/update-integrity.cjs` to refresh the integrity attributes in `index.html`. Run `node --test tests/security.test.cjs` for security and before/after comparison checks. The browser test also exercises the CSP, worker timeout/cancellation, imports, clipboard, and real downloads: `node tests/browser-security.cjs` (requires Chrome and permission to launch a temporary isolated profile).

Bundled dependencies and their download sources are recorded in `vendor/README.md`. When upgrading SheetJS, update both script paths, regenerate integrity hashes, and rerun the tests. Font files and their licenses are in `vendor/fonts/`.

## Deployment

GitHub Pages serves `index.html` directly from the root of the `main` branch. Updates pushed to `main` are published automatically.
