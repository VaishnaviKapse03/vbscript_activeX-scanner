# vbscript_activeX-scanner
AI agent that scans Classic ASP and ASP.NET pages for  client-side VBScript blocks and ActiveX COM control  dependencies — built to support legacy IE-to-modern-browser  migration projects

## What it does

- Walks a directory of `.asp`, `.aspx`, `.html`, `.htm` files
- Finds pages with real client-side VBScript code
- Detects ActiveX objects used via `CreateObject()` in VBScript or `<object>` tags in HTML
- Identifies each control by name, category, and description
- For unknown controls, falls back to the Claude API and caches the result for future scans
- Writes two CSV reports into an `output/` folder created inside the scanned directory

## Setup

```bash
pip install beautifulsoup4 lxml anthropic
```

Set your Anthropic API key:

```bash
# macOS / Linux
export ANTHROPIC_API_KEY="sk-ant-..."

# Windows
set ANTHROPIC_API_KEY=sk-ant-...
```

## Usage

```bash
python reporter.py <path_to_directory>
```

The tool will create an `output/` folder inside the directory you provide and write two files there:

- `vbscript_report.csv` — all pages that contain real client-side VBScript
- `activex_report.csv` — all ActiveX controls found, one row per control, with name, category, and description

## Project structure

```
├── scanner.py          # Walks files, detects VBScript and ActiveX
├── activex_known.py    # Built-in control registry and learned cache
├── reporter.py         # Orchestrates the scan and writes CSVs
└── learned_cache.csv   # Auto-generated; saves Claude API results
```

## How unknown controls are identified

If a control isn't in the built-in registry, the tool calls the Claude API to identify it and saves the result to `learned_cache.csv`. On future scans, the cache is checked first so the API isn't called again for the same control.

## License

MIT
