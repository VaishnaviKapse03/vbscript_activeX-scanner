# reporter.py
# ─────────────────────────────────────────────
# Takes scan results from scanner.py and writes
# them out as two CSV files:
#   1. vbscript_report.csv  — List A (pages with client VBScript)
#   2. activex_report.csv   — List B (pages with ActiveX details)
# ─────────────────────────────────────────────

import csv
import sys
import scanner
import activex_known


def run_scan(folder):
    """
    Triggers the actual scan by calling scanner's functions directly,
    since importing scanner.py alone does NOT run its __main__ block.
    """
    activex_known.load_cache()

    file_list = scanner.traverse(folder)
    print(f"Found {len(file_list)} files to scan\n")

    for file_path in file_list:
        scanner.file_analyzer(file_path)

    print(f"Files with real client-side VBScript: {len(scanner.vbscript_pages)}")

    pages_with_activex = [f for f in scanner.activex_pages if scanner.activex_pages[f]["activex_found"]]
    print(f"Files with ActiveX: {len(pages_with_activex)}\n")


def write_vbscript_report(output_path="output/vbscript_report.csv"):
    """
    Writes List A — every page with real client-side VBScript.
    One row per file.
    """
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file_path"])

        for page in sorted(scanner.vbscript_pages):
            writer.writerow([page])

    print(f"VBScript report written to {output_path}")


def write_activex_report(output_path="output/activex_report.csv"):
    """
    Writes List B — every ActiveX element found, one row per element.
    A page with 2 ActiveX controls will appear as 2 rows.
    """
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "file_path",
            "detection_type",
            "progid_or_clsid",
            "friendly_name",
            "category",
            "description"
        ])

        for page in sorted(scanner.activex_pages):
            findings = scanner.activex_pages[page]["activex_found"]
            for item in findings:
                writer.writerow([
                    page,
                    item["type"],
                    item["progid"],
                    item["friendly_name"],
                    item["category"],
                    item["description"]
                ])

    print(f"ActiveX report written to {output_path}")


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else "test_corpus"

    run_scan(folder)
    write_vbscript_report()
    write_activex_report()