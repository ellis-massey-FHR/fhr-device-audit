import os
import json
from pathlib import Path
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image as XLImage

# -----------------------Load environment----------------------

env_path = Path(__file__).resolve().parent / "servicenow.env"
load_dotenv(dotenv_path=env_path)

instance = os.getenv("SERVICENOW_INSTANCE")
username = os.getenv("SERVICENOW_USER")
password = os.getenv("SERVICENOW_PASS")

print("ENV path:", env_path)
print("Instance loaded:", instance)
print("Username loaded:", username)
print("Password loaded:", bool(password))

#-----------------------Fetch Full User and Device Info----------------

def fetch_details(link):
    if not link:
        return {}
    res = requests.get(link, auth=HTTPBasicAuth(username, password), params={"sysparm_display_value": "true"})
    if res.status_code == 200:
        return res.json().get("result", {})
    return {}

#-------------------------Clean formatting of workstation strings-----------------

def format_workstation_string(ws_string):
    return ws_string.replace("\\n", ", ").replace("\\r", "").replace("\\n", ", ").replace("\\r", "").strip().lstrip(", ")

#-------------------------Pull workstation CI's for a user-----------------------

def get_current_workstation_from_ci(user_sys_id):
    ci_url = f"{instance}/api/now/table/cmdb_ci_computer"
    params = {
        "sysparm_query": (
            f"assigned_to={user_sys_id}"
            "^model_category.name=Computer"
            "^ORDERBYDESClast_discovered"
        ),
        "sysparm_display_value": "true",
        "sysparm_limit": "50"
    }

    response = requests.get(ci_url, auth=HTTPBasicAuth(username, password), params=params)

    if response.status_code != 200:
        print(f"❌ Error fetching CIs: {response.status_code}")
        return None

    records = response.json().get("result", [])
    all_devices = []

    for r in records:
        name = r.get("name")
        class_name = r.get("sys_class_name")

        # model_id may be str or dict depending on sysparm_display_value
        model_id = r.get("model_id")
        if isinstance(model_id, dict):
            model_name = (model_id.get("display_value") or "").lower()
        elif isinstance(model_id, str):
            model_name = model_id.lower()
        else:
            model_name = ""

        # --- OS handling: include Windows or UNKNOWN; skip only if we know it's not Windows
        os_raw = r.get("os")
        os_name = os_raw.lower() if isinstance(os_raw, str) else ""

        # Fallbacks in case OS is populated under different fields
        if not os_name:
            alt_os = r.get("operating_system") or r.get("u_os_full_name")
            if isinstance(alt_os, str):
                os_name = alt_os.lower()
            elif isinstance(alt_os, dict):
                os_name = (alt_os.get("display_value") or alt_os.get("value") or "").lower()

        # Skip if we positively know it's not Windows
        if os_name and "windows" not in os_name:
            print(f"🛑 Skipping non-Windows device: {name} (OS: {os_name})")
            continue

        #-----------------Sanitize name to prevent newlines from breaking formatting-----------

        if isinstance(name, str):
            name = name.replace("\n", " ").replace("\r", " ").strip()

        if name and class_name:
            device_str = name
            if device_str not in all_devices:
                all_devices.append(device_str)

    return format_workstation_string(", ".join(all_devices)) if all_devices else None


#----------------------Query ServiceNow--------------------------

url = f"{instance}/api/now/table/sc_req_item"
params = {
    "sysparm_query": (
        "company.name=Flint Hills Resources"
        "^u_new_hire=false"
        "^cat_item.nameLIKElaptop"
        "^ORcat_item.nameLIKEsurface"
        "^opened_atRELGTjavascript:gs.daysAgoStart(60)^ORclosed_atRELGTjavascript:gs.daysAgoStart(60)"
        "^ORDERBYDESCclosed_at"
    ),
    "sysparm_display_value": "all",
    "sysparm_limit": "115"
}

response = requests.get(url, auth=HTTPBasicAuth(username, password), params=params)

if response.status_code == 200:
    print("✅ Success!")
    data = response.json()
    print(f"Total records returned: {len(data['result'])}")

    rows = []
    for item in data["result"]:
        print("→ Processing item ID:", item.get("sys_id"))

        requested_for_info = fetch_details(item.get("requested_for", {}).get("link"))

        user_location = (
            requested_for_info.get("location", {}).get("display_value")
            or requested_for_info.get("u_location", {}).get("display_value")
            or requested_for_info.get("u_site_id")
            or "Unknown"
        )

        user_sys_id = requested_for_info.get("sys_id")
        recent_ws = get_current_workstation_from_ci(user_sys_id)

        if not recent_ws:
            recent_ws = format_workstation_string(requested_for_info.get("u_workstations", ""))

        rows.append({
            "Requested For": item.get("requested_for", {}).get("display_value", "Unknown"),
            "Requested For Email": requested_for_info.get("email", "Unknown"),
            "Requested For Location": user_location,
            "Catalog Item": item.get("cat_item", {}).get("display_value", "Unknown"),
            "Created": item.get("opened_at", {}).get("display_value", "Unknown") if isinstance(item.get("opened_at"), dict) else item.get("opened_at", "Unknown"),
            "Closed": item.get("closed_at", {}).get("display_value", "Unknown") if isinstance(item.get("closed_at"), dict) else item.get("closed_at", "Unknown"),
            "Workstations": recent_ws or "Unknown"
        })

    df = pd.DataFrame(rows)

    #------------------------------Group by Location--------------------

    def normalize_location(loc):
        if not loc:
            return "Unknown"
        loc = loc.upper()
        if "CORPUSCHRISTI" in loc or "CORPUS CHRISTI" in loc:
            return "Corpus Christi"
        elif "ROSEMOUNT" in loc or "PINE BEND" in loc:
            return "Rosemount"
        elif "WICHITA" in loc:
            return "Wichita"
        else:
            return "Other"

    df["Location Group"] = df["Requested For Location"].apply(normalize_location)

    #--------------------------Write to CSV and Excel-------------------------

    from pathlib import Path
    output_dir = Path(os.getenv("OUTPUT_DIR", "output"))   # default: ./output
    output_dir.mkdir(parents=True, exist_ok=True)
    excel_path = output_dir / "fhr_computer_requests.xlsx"
    
# df.to_excel(excel_path, index=False, engine="openpyxl")

    #---------------------------Format Excel Header-------------------------

    wb = load_workbook(excel_path)
    ws = wb.active

    ws.insert_rows(1)
    header_text = "FHR Laptop Requests Report"
    total_columns = ws.max_column
    end_col_letter = get_column_letter(total_columns)
    ws.merge_cells(f"A1:{end_col_letter}1")
    header_cell = ws["A1"]
    header_cell.value = header_text
    header_cell.font = Font(size=16, bold=True)
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    #-------------------------Adjust column widths------------------------

    for col in ws.columns:
        max_length = 0
        column = col[0].column
        col_letter = get_column_letter(column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        if ws.cell(row=2, column=column).value == "Workstations":
            ws.column_dimensions[col_letter].width = 40
        else:
            ws.column_dimensions[col_letter].width = max_length + 2

    from io import BytesIO
    import matplotlib.pyplot as plt
    from openpyxl.drawing.image import Image as XLImage

    #------------------------Embed pie chart--------------------------

    location_counts = df["Location Group"].value_counts()

    def make_autopct(values):
        def my_autopct(pct):
            total = sum(values)
            count = int(round(pct * total / 100.0))
            return f"{count}"
        return my_autopct

    pie_buffer = BytesIO()
    plt.figure(figsize=(6, 6))
    plt.pie(location_counts, labels=location_counts.index, autopct=make_autopct(location_counts), startangle=140)
    plt.title(f"Laptop Requests by Location (Total: {location_counts.sum()})")
    plt.axis("equal")
    plt.tight_layout()
    plt.savefig(pie_buffer, format="png")
    plt.close()

    pie_buffer.seek(0)
    img = XLImage(pie_buffer)
    img.width = 350
    img.height = 350
    ws.add_image(img, "J6")
    print("📊 Pie chart embedded in Excel at J6")

    wb.save(excel_path)
    print(f"📁 Excel with embedded chart saved to {excel_path}")

else:
    print(f"❌ Error: {response.status_code}")
    print(response.text)
