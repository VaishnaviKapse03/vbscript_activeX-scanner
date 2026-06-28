# activex_known.py
# ─────────────────────────────────────────────
# Responsibilities:
#   1. Store known ActiveX ProgIDs and CLSIDs
#   2. Provide a lookup function for the scanner
#   3. Manage learned_cache.csv for unknown
#      controls discovered by Claude API
# ─────────────────────────────────────────────

import csv
import os

# ─────────────────────────────────────────────
# CACHE FILE PATH
# ─────────────────────────────────────────────
CACHE_FILE = os.path.join(os.path.dirname(__file__), "learned_cache.csv")

# ─────────────────────────────────────────────
# KNOWN PROGIDS
# key   → lowercase ProgID
# value → (friendly_name, category, description)
# ─────────────────────────────────────────────
KNOWN_PROGIDS = {

    # PRINT
    "massprint.document":            ("MassPrint Document Control",        "PRINT",    "Batch document printing — queues and prints multiple documents silently"),
    "massprint.labelengine":         ("MassPrint Label Engine",            "PRINT",    "Label printing — renders and prints label templates"),
    "massprint.spoolmanager":        ("MassPrint Spool Manager",           "PRINT",    "Print spooler management — controls and monitors print queue jobs"),
    "pdfcreator.clspdfcreator":      ("PDFCreator",                        "PRINT",    "PDF creation — prints documents to PDF via virtual printer"),
    "fineprint.sdk":                 ("FinePrint SDK",                     "PRINT",    "FinePrint print preview and management SDK"),

    # REPORT
    "launchview.viewer":             ("LaunchView Report Viewer",          "REPORT",   "Launches and displays reports embedded inside the browser page"),
    "crystalruntime.application":    ("Crystal Reports Runtime",           "REPORT",   "Opens, renders and prints Crystal Reports .rpt files"),
    "crystalruntime.printengine":    ("Crystal Reports Print Engine",      "REPORT",   "Handles Crystal Reports output to printer or file"),
    "cognosreportviewer.control":    ("IBM Cognos Report Viewer",          "REPORT",   "IBM Cognos embedded report viewer control"),
    "saproireports.application":     ("SAP ROI Reports",                   "REPORT",   "Connects to SAP and runs report objects"),
    "activereports.activereport":    ("GrapeCity ActiveReports",           "REPORT",   "Viewer control for ActiveReports .NET report files"),

    # OFFICE
    "excel.application":             ("Microsoft Excel",                   "OFFICE",   "Creates, reads and writes Excel workbooks client-side"),
    "excel.sheet":                   ("Microsoft Excel Sheet",             "OFFICE",   "Embeds a live Excel spreadsheet in the page"),
    "excel.chart":                   ("Microsoft Excel Chart",             "OFFICE",   "Renders an Excel chart client-side"),
    "word.application":              ("Microsoft Word",                    "OFFICE",   "Creates and manipulates Word documents client-side"),
    "word.document":                 ("Microsoft Word Document",           "OFFICE",   "Opens and renders a Word document in the page"),
    "outlook.application":           ("Microsoft Outlook",                 "OFFICE",   "Sends email and creates appointments via client Outlook"),
    "powerpoint.application":        ("Microsoft PowerPoint",              "OFFICE",   "Creates and opens presentations client-side"),
    "access.application":            ("Microsoft Access",                  "OFFICE",   "Opens and queries Access databases client-side"),
    "owc11.spreadsheet":             ("Office Web Components Spreadsheet", "OFFICE",   "Interactive spreadsheet embedded in browser — Office 2003"),
    "owc11.chart":                   ("Office Web Components Chart",       "OFFICE",   "Interactive chart embedded in browser — Office 2003"),
    "owc11.pivottable":              ("Office Web Components PivotTable",  "OFFICE",   "Pivot table embedded in browser — Office 2003"),

    # DATA
    "adodb.connection":              ("ADO Database Connection",           "DATA",     "Opens a database connection client-side"),
    "adodb.recordset":               ("ADO Recordset",                     "DATA",     "Holds and navigates query results client-side"),
    "adodb.command":                 ("ADO Command",                       "DATA",     "Executes parameterised SQL client-side"),
    "adodb.stream":                  ("ADO Stream",                        "DATA",     "Reads and writes binary or text streams client-side"),
    "rds.datacontrol":               ("Remote Data Service Control",       "DATA",     "Binds HTML elements directly to server-side recordsets"),
    "rds.dataspace":                 ("Remote Data Service DataSpace",     "DATA",     "Creates server-side business objects from the client"),
    "microsoft.xmldom":              ("Microsoft XML DOM",                 "DATA",     "Legacy XML DOM parser — loads and queries XML documents"),
    "msxml2.domdocument":            ("MSXML2 DOM Document",               "DATA",     "Microsoft XML 2.0 DOM parser"),
    "msxml2.domdocument.6.0":        ("MSXML2 DOM Document 6.0",           "DATA",     "Microsoft XML 6.0 DOM parser — latest MSXML DOM"),

    # NETWORK
    "microsoft.xmlhttp":             ("Microsoft XMLHTTP",                 "NETWORK",  "Makes HTTP requests from client-side VBScript — pre-AJAX era"),
    "msxml2.xmlhttp":                ("MSXML2 XMLHTTP",                    "NETWORK",  "Sends HTTP GET/POST requests from client VBScript"),
    "msxml2.xmlhttp.6.0":            ("MSXML2 XMLHTTP 6.0",                "NETWORK",  "Latest version of the XMLHTTP request control"),
    "wscript.network":               ("Windows Script Host Network",       "NETWORK",  "Maps drives, connects printers, reads network info client-side"),
    "mapi.session":                  ("MAPI Mail Session",                 "NETWORK",  "Logs into mail server and sends messages via client mail profile"),
    "cdo.message":                   ("CDO Message",                       "NETWORK",  "Sends email via SMTP client-side"),
    "imcontrol.messenger":           ("Instant Messenger Control",         "NETWORK",  "Connects to intranet IM server and sends messages"),

    # FILE
    "scripting.filesystemobject":    ("FileSystemObject",                  "FILE",     "Reads, writes, creates and deletes files and folders on the client machine"),
    "scripting.dictionary":          ("Scripting Dictionary",              "FILE",     "Key-value store used client-side in VBScript"),

    # SYSTEM
    "wscript.shell":                 ("Windows Script Host Shell",         "SYSTEM",   "Runs executables, reads registry, creates shortcuts client-side"),
    "wbemscripting.swbemlocator":    ("WMI Scripting Locator",             "SYSTEM",   "Connects to WMI to query system information client-side"),
    "shell.application":             ("Shell Application",                 "SYSTEM",   "Browses folders, launches files, manages shell operations"),
    "vmware.vmcom":                  ("VMware VmCOM",                      "SYSTEM",   "Automates VMware virtual machine operations client-side"),

    # MEDIA
    "wmplayer.ocx":                  ("Windows Media Player",              "MEDIA",    "Plays audio and video streams inside the browser"),
    "wmplayer.ocx.7":                ("Windows Media Player 7",            "MEDIA",    "Windows Media Player 7 ActiveX control"),
    "shockwaveflash.shockwaveflash": ("Adobe Flash Player",                "MEDIA",    "Plays SWF Flash animations and applications"),
    "realplayer.realplayer":         ("RealPlayer",                        "MEDIA",    "RealNetworks player — plays streaming audio and video"),
    "quicktime.quicktime":           ("Apple QuickTime",                   "MEDIA",    "Plays MOV and multimedia files"),

    # UI
    "msflexgrid.msflexgrid":         ("MSFlexGrid",                        "UI",       "Scrollable data grid control for tabular data display"),
    "mscomctllib.treectrl":          ("TreeView Control",                  "UI",       "Hierarchical tree navigation control"),
    "mscomctllib.listviewctrl":      ("ListView Control",                  "UI",       "Multi-column list control with sorting"),
    "mscomctllib.tabstrip":          ("TabStrip Control",                  "UI",       "Tabbed panel navigation control"),
    "mscomctllib.progressbar":       ("ProgressBar Control",               "UI",       "Shows task completion percentage"),
    "mscal.calendar":                ("Microsoft Calendar Control",        "UI",       "Date picker and calendar display control"),
    "mscomdlg.commondialog":         ("Common Dialog Control",             "UI",       "Shows Windows Open/Save/Print/Color dialog boxes"),
    "agent.control":                 ("Microsoft Agent",                   "UI",       "Animated character control — Merlin, Clippy etc."),
    "marqueetag.marqueetag":         ("Marquee Control",                  "UI",       "Scrolling text marquee control"),
    # SECURITY
    "capicom.certificate":           ("CAPICOM Certificate",               "SECURITY", "Accesses Windows certificate store for digital signing"),
    "capicom.signeddata":            ("CAPICOM Signed Data",               "SECURITY", "Signs data using a certificate from the Windows store"),

    # MAPPING
    "mapviewer.control":             ("MapViewer Control",                 "MAPPING",  "Loads and displays GIS map files in the browser"),
    "mapx.map":                      ("MapX Mapping Control",              "MAPPING",  "Pitney Bowes MapX interactive GIS mapping control"),

    # FINANCE
    "bloomberg.data":                ("Bloomberg Data Feed",               "FINANCE",  "Retrieves real-time market data from Bloomberg terminal"),
    "reuters.rtfeed":                ("Reuters Real-Time Feed",            "FINANCE",  "Subscribes to real-time Reuters financial data feeds"),

    # SCANNER / HARDWARE
    "barcodepro.scanner":            ("BarcodePro Scanner",                "SCANNER",  "Reads barcodes from attached hardware"),
    "signaturepad.sigctl":           ("Topaz Signature Pad",               "SCANNER",  "Captures handwritten signatures from Topaz signature pad"),

    # PDF
    "acropdf.pdf":                   ("Adobe Acrobat PDF Viewer",          "PDF",      "Embeds and displays PDF documents inside the browser"),
    "acropdf.pdf.1":                 ("Adobe Acrobat PDF Viewer 1",        "PDF",      "Legacy Adobe Acrobat viewer ActiveX control"),
    "pdf.pdfctrl":                   ("Adobe PDF Control",                 "PDF",      "Older Adobe Acrobat browser viewer ActiveX"),
}

# ─────────────────────────────────────────────
# KNOWN CLSIDS
# key   → lowercase CLSID (no braces)
# value → lowercase ProgID string
# ─────────────────────────────────────────────
KNOWN_CLSIDS = {
    # PRINT
    "5220cb21-c88d-11cf-b347-00aa00a28331": "massprint.document",
    "09ea5560-2d41-4cb4-9d93-c6b8b9dca2ec": "barcodepro.scanner",

    # REPORT
    "9cb1b849-8e2a-4c6a-a428-39a20a45a8b3": "crystalruntime.application",
    "22d6f312-b0f6-11d0-94ab-0080c74c7e95": "launchview.viewer",

    # OFFICE
    "0002e541-0000-0000-c000-000000000046": "owc11.spreadsheet",
    "0002e500-0000-0000-c000-000000000046": "owc11.chart",

    # DATA
    "f6d90f11-9c73-11d3-b32e-00c04f990bb4": "msxml2.domdocument",
    "f5078f32-c551-11d3-89b9-0000f81fe221": "microsoft.xmlhttp",
    "bd96c556-65a3-11d0-983a-00c04fc29e33": "rds.datacontrol",

    # MEDIA
    "6bf52a52-394a-11d3-b153-00c04f79faa6": "wmplayer.ocx",
    "22d6f311-b0f6-11d0-94ab-0080c74c7e95": "wmplayer.ocx",
    "d27cdb6e-ae6d-11cf-96b8-444553540000": "shockwaveflash.shockwaveflash",

    # UI
    "8e27c92b-1264-101c-8a2f-040224009c02": "mscal.calendar",
    "6262d3a0-531b-11cf-91f6-c2863c385e30": "msflexgrid.msflexgrid",
    "c74190b6-8589-11d1-b16a-00c0f0283628": "mscomctllib.treectrl",
    "bdd1f04b-858b-11d1-b16a-00c0f0283628": "mscomctllib.listviewctrl",
    "0713e8d2-850a-101b-afc0-4210102a8da7": "mscomctllib.progressbar",
    "d45fd31b-5c6e-11d1-9ec1-00c04fd7081f": "agent.control",
    "1a4da620-6217-11cf-be62-0080c72edd2d": "marqueetag.marqueetag",

    # SYSTEM
    "76a64158-cb41-11d1-8b02-00600806d9b6": "wbemscripting.swbemlocator",

    # MAPPING
    "22d6f315-b0f6-11d0-94ab-0080c74c7e95": "mapx.map",

    # PDF
    "ca8a9780-280d-11cf-a24d-444553540000": "acropdf.pdf",
}

# ─────────────────────────────────────────────
# IN-MEMORY CACHE
# populated by load_cache() at start of run
# ─────────────────────────────────────────────
_learned_cache = {}


def load_cache():
    global _learned_cache
    _learned_cache = {}

    if not os.path.exists(CACHE_FILE):
        return

    with open(CACHE_FILE, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            key = row["progid"].strip().lower()
            _learned_cache[key] = (
                row["friendly_name"],
                row["category"],
                row["description"]
            )


def save_to_cache(progid, friendly_name, category, description):
    key = progid.strip().lower()
    _learned_cache[key] = (friendly_name, category, description)

    file_exists = os.path.exists(CACHE_FILE)

    with open(CACHE_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["progid", "friendly_name", "category", "description"]
        )
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "progid":        progid,
            "friendly_name": friendly_name,
            "category":      category,
            "description":   description
        })


def lookup(identifier):
    clean = identifier.strip().lower().strip("{}")

    if "-" in clean:
        # ── CLSID path ──
        progid = KNOWN_CLSIDS.get(clean)
        if progid:
            result = KNOWN_PROGIDS.get(progid)
            if result:
                return result
            return _learned_cache.get(progid)
        return None
    else:
        # ── ProgID path ──
        result = KNOWN_PROGIDS.get(clean)
        if result:
            return result
        return _learned_cache.get(clean)


