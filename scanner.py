import os
from bs4 import BeautifulSoup
import re
import activex_known
import anthropic

vbscript_pages = set()   
activex_pages  = {}

def traverse(root_folder):
    file_list = []
    for folder, subfolders, files in os.walk(root_folder):      
        for file in files:
            if file.endswith((".asp", ".aspx", ".html", ".htm")):  
                file_list.append(os.path.join(folder, file))
    return file_list                                             


def has_real_code(script_text):
    for line in script_text.splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("'"):
            return True
    return False

def ask_claude(progid):
    try:
        client = anthropic.Anthropic() 
        #API Key needs to be set. Once the API key is set this line above will communicate with the system to find that key and make claude calls.
        
        prompt = f"""You are an expert in Windows ActiveX and COM components.
I found this ProgID in client-side VBScript code: "{progid}"

Tell me:
1. Is this a known Windows ActiveX or COM control?
2. If yes, provide:
   - friendly_name: short human readable name
   - category: one of PRINT, REPORT, OFFICE, DATA, NETWORK, FILE, SYSTEM, MEDIA, UI, SECURITY, MAPPING, FINANCE, SCANNER, PDF
   - description: one sentence explaining what it does

If this appears to be a custom internal object (not a standard ActiveX control), respond with category: CUSTOM

Respond in this exact format:
friendly_name: <name>
category: <category>
description: <description>"""

        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response = message.content[0].text
        
        # parse the response
        lines = response.strip().splitlines()
        data = {}
        for line in lines:
            if ":" in line:
                key, value = line.split(":", 1)
                data[key.strip().lower()] = value.strip()
        
        friendly_name = data.get("friendly_name", progid)
        category      = data.get("category", "UNKNOWN").upper()
        description   = data.get("description", "Identified by Claude API")

        # skip custom/internal objects
        if category in ("CUSTOM", "UNKNOWN"):
            return None
            
        return (friendly_name, category, description)

    except Exception as e:
        print(f"  Claude API error for {progid}: {e}")
        return None


def chk_activeX(file_path, script_text, soup):
    if file_path not in activex_pages:
        activex_pages[file_path] = {"activex_found": []}

    # ── TYPE 1 — CreateObject() inside client-side VBScript ──
    progids_found = re.findall(r'CreateObject\("([^"]+)"\)', script_text)

    for progid in progids_found:
        result = activex_known.lookup(progid)

        if result:
            friendly_name, category, description = result
            activex_pages[file_path]["activex_found"].append({
                "progid":        progid,
                "type":          "Script",
                "friendly_name": friendly_name,
                "category":      category,
                "description":   description
            })
        else:
            api_result = ask_claude(progid)
            if api_result:
                friendly_name, category, description = api_result
                activex_known.save_to_cache(progid, friendly_name, category, description)
                activex_pages[file_path]["activex_found"].append({
                    "progid":        progid,
                    "type":          "Script",
                    "friendly_name": friendly_name,
                    "category":      category,
                    "description":   description
                })

    # ── TYPE 2 — <object classid="..."> or <object progid="..."> in HTML ──
    object_tags = soup.find_all("object")

    for obj in object_tags:
        classid = obj.get("classid", "").replace("clsid:", "").replace("CLSID:", "")
        progid  = obj.get("progid", "")

        identifier = ""
        if classid:
            identifier = classid
        elif progid:
            identifier = progid

        if identifier:
            result = activex_known.lookup(identifier)

            if result:
                friendly_name, category, description = result
                activex_pages[file_path]["activex_found"].append({
                    "progid":        identifier,
                    "type":          "HTML",
                    "friendly_name": friendly_name,
                    "category":      category,
                    "description":   description
                })
            else:
                api_result = ask_claude(identifier)
                if api_result:
                    friendly_name, category, description = api_result
                    activex_known.save_to_cache(identifier, friendly_name, category, description)
                    activex_pages[file_path]["activex_found"].append({
                        "progid":        identifier,
                        "type":          "HTML",
                        "friendly_name": friendly_name,
                        "category":      category,
                        "description":   description
                    })



def file_analyzer(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    soup = BeautifulSoup(content, "lxml")
    script_tags = soup.find_all("script")
    
    for script in script_tags:
        language = script.get("language", "").strip().lower()
        runat = script.get("runat", "").strip().lower()
        
        if language == "vbscript" and runat != "server":
            script_text = script.get_text()
            if has_real_code(script_text):
                vbscript_pages.add(file_path)
                chk_activeX(file_path, script_text, soup)









if __name__ == "__main__":
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else "test_corpus"

    activex_known.load_cache()

    file_list = traverse(folder)
    print(f"Found {len(file_list)} files to scan\n")

    for file_path in file_list:
        file_analyzer(file_path)

    print(f"Files with real client-side VBScript: {len(vbscript_pages)}")

    pages_with_activex = [f for f in activex_pages if activex_pages[f]["activex_found"]]
    print(f"Files with ActiveX: {len(pages_with_activex)}\n")

    for page in sorted(pages_with_activex):
        print(f"\n{page}")
        for item in activex_pages[page]["activex_found"]:
            print(f"    [{item['type']}] {item['progid']} -> {item['friendly_name']} ({item['category']})")