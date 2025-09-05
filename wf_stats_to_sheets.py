#!/usr/bin/env python3
import os
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import statistics
from datetime import datetime
from dotenv import load_dotenv

# --- Load .env ---
load_dotenv()

# --- Config ---
API_URL =  os.getenv("API_URL")
SHEET_NAME = "Warframe Stats"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, os.getenv("CRED_FILE"))

# --- Google Sheets Setup ---
scope = [
    os.getenv("SCOPE_SHEET_URL"),
    os.getenv("SCOPE_DRIVE_URL"),
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

# ⚡ Worksheet direkt holen, damit wir später den Namen ändern können
worksheet = client.open(SHEET_NAME).sheet1

# --- Helper ---
def col_to_letter(col_index):
    """0-based index -> A, B, ..., Z, AA, AB, ..."""
    n = col_index
    letters = ""
    while n >= 0:
        letters = chr(ord('A') + (n % 26)) + letters
        n = n // 26 - 1
    return letters

def scale_stats(wf, rank=30):
    """Berechnet die Rank-30 Stats basierend auf DE Growth-Formeln."""
    health = wf.get("health", 0)
    shield = wf.get("shield", 0)
    armor  = wf.get("armor", 0)
    power  = wf.get("power", 0)

    parent = wf.get("uniqueName", "")

    if parent in ["/Lotus/Powersuits/Sandman/SandmanBaseSuit",  # Inaros
                  "/Lotus/Powersuits/Devourer/DevourerBaseSuit"]:  # Grendel
        health += (rank + 2) // 3 * 10
        health += (rank + 1) // 3 * 10
        power  += rank // 3 * 5

    elif parent == "/Lotus/Powersuits/Infestation/InfestationBaseSuit":  # Nidus
        health += (rank + 2) // 3 * 10
        armor  += (rank + 4) // 6 * 20
        power  += rank // 6 * 10
        # heal_rate + ability_strength lassen wir weg, brauchst du nicht in deiner Tabelle

    elif parent == "/Lotus/Powersuits/PaxDuviricus/PaxDuviricusBaseSuit":  # Kullervo
        health += (rank + 2) // 3 * 20
        armor  += (rank + 1) // 3 * 10
        power  += rank // 3 * 5

    elif parent == "/Lotus/Powersuits/IronFrame/IronFrameBaseSuit":  # Hildryn
        health += (rank + 2) // 3 * 10
        shield += (rank + 1) // 3 * 25
        shield += rank // 3 * 25

    elif parent == "/Lotus/Powersuits/BrokenFrame/BrokenFrameBaseSuit":  # Xaku
        health += (rank + 2) // 3 * 9
        shield += (rank + 1) // 3 * 9
        power  += rank // 3 * 7

    elif parent == "/Lotus/Powersuits/Alchemist/AlchemistBaseSuit":  # Lavos
        health += (rank + 2) // 3 * 20
        shield += (rank + 1) // 3 * 10
        armor  += rank // 3 * 10

    elif parent == "/Lotus/Powersuits/Berserker/BerserkerBaseSuit":  # Valkyr
        health += (rank + 2) // 3 * 10
        shield += (rank + 1) // 3 * 5
        power  += rank // 3 * 5

    elif parent in [
        "/Lotus/Powersuits/Pacifist/PacifistBaseSuit",  # Baruuk
        "/Lotus/Powersuits/Garuda/GarudaBaseSuit",     # Garuda
        "/Lotus/Powersuits/Wisp/WispBaseSuit",         # Wisp
        "/Lotus/Powersuits/Yareli/YareliBaseSuit"      # Yareli
    ]:
        health += (rank + 2) // 3 * 10
        shield += (rank + 1) // 3 * 10
        power  += rank // 3 * 10

    else:  # Default
        health += (rank + 2) // 3 * 10
        shield += (rank + 1) // 3 * 10
        power  += rank // 3 * 5

    return health, shield, armor, power

# --- Fetch Data ---
print("Fetching Warframe data...")
resp = requests.get(API_URL)
resp.raise_for_status()
warframes = resp.json()

# --- Prepare Data ---
headers = [
    "Name", "Health", "Armor", "Effective Health",
    "Shields", "Energy Cap", "Energy At Spawn", "Sprint",
    "Max Overshields", "EHP w/ Shields", "EHP w/ Shields & Overshields"
]

numeric_cols_idx = list(range(1, len(headers)))  # columns to average/compare

# --- Custom Warframes hinzufügen ---
custom_frames = [
    {
        "name": "Caliban Prime",
        "health": 370,
        "shields": 740,
        "armor": 290,
        "energy": 225,
        "energy_spawn": 100,
        "sprint": 1.1
    }
]

# --- Exclude Warframes ---
EXCLUDE_FRAMES = ["Voidrig", "Bonewidow", "Helminth"]
frame_rows = []
seen_names = set()  # um doppelte Namen zu verhindern

for wf in warframes:
    if wf.get("type") != "Warframe":
        continue
    name = wf.get("name")
    if name in EXCLUDE_FRAMES or name in seen_names:
        continue
    seen_names.add(name)

    health, shields, armor, energy = scale_stats(wf, rank=30)
    sprint = wf.get("sprintSpeed", 0)

    # --- Overshield berechnen ---
    if shields == 0:
        max_overshields = 0
    elif name in ["Harrow", "Harrow Prime"]:
        max_overshields = 2400
    else:
        max_overshields = 1200

    effective_health = health * (1 + armor / 300)
    ehp_with_shields = effective_health + shields
    ehp_with_overshields = ehp_with_shields + max_overshields
    energy_spawn = energy * 0.5

    row = [
        name,
        round(health, 2),
        round(armor, 2),
        round(effective_health, 2),
        round(shields, 2),
        round(energy, 2),
        round(energy_spawn, 2),
        round(sprint, 2),
        round(max_overshields, 2),
        round(ehp_with_shields, 2),
        round(ehp_with_overshields, 2)
    ]
    frame_rows.append(row)

# --- Custom Frames hinzufügen, falls nicht schon da ---
for wf in custom_frames:
    if wf["name"] not in seen_names:
        seen_names.add(wf["name"])
        if wf["shields"] == 0:
            max_overshields = 0
        elif wf["name"] in ["Harrow", "Harrow Prime"]:
            max_overshields = 2400
        else:
            max_overshields = 1200        
        health, shields, armor, energy = scale_stats(wf, rank=30)
        effective_health = health * (1 + armor / 300)
        ehp_with_shields = effective_health + shields
        ehp_with_overshields = ehp_with_shields + max_overshields
        row = [
            wf["name"],
            round(health, 2),
            round(armor, 2),
            round(effective_health, 2),
            round(shields, 2),
            round(energy, 2),
            round(wf["energy_spawn"], 2),
            round(wf["sprint"], 2),
            round(max_overshields, 2),
            round(ehp_with_shields, 2),
            round(ehp_with_overshields, 2)
        ]
        frame_rows.append(row)

# --- Compute medians ---
average_row = ["Median"]
for col_idx in numeric_cols_idx:
    col_values = [r[col_idx] for r in frame_rows if isinstance(r[col_idx], (int, float))]
    med = round(statistics.median(col_values), 2) if col_values else 0
    average_row.append(med)

# --- Sortiere frame_rows alphabetisch nach Name ---
frame_rows.sort(key=lambda x: x[0])

rows = [average_row, headers] + frame_rows

# --- Write all values first ---
worksheet.clear()
worksheet.update(range_name="A1", values=rows)

# --- Format each cell with batch formatting ---
batch_formats = []

base_bg = {"red": 6/255, "green": 15/255, "blue": 20/255}
base_text = {"red": 244/255, "green": 241/255, "blue": 208/255}
green_text = {"red": 0.6, "green": 0.9, "blue": 0.6}
red_text = {"red": 0.95, "green": 0.6, "blue": 0.6}

for r_idx, row in enumerate(rows):
    for c_idx, val in enumerate(row):
        fmt = {
            "backgroundColor": base_bg,
            "textFormat": {
                "foregroundColor": base_text,
                "fontSize": 16
            }
        }
        if r_idx >= 2 and c_idx in numeric_cols_idx:
            avg_val = average_row[c_idx]
            tolerance = avg_val * 0.10
            if val > avg_val + tolerance:
                fmt["textFormat"]["foregroundColor"] = green_text
            elif val < avg_val - tolerance:
                fmt["textFormat"]["foregroundColor"] = red_text

        cell_label = f"{col_to_letter(c_idx)}{r_idx+1}"
        batch_formats.append({"range": cell_label, "format": fmt})

# --- Freeze + Filter ---
worksheet.freeze(rows=2)
worksheet.set_basic_filter(f"A2:K{len(rows)}")

# --- ✅ Info-Block vorbereiten (L4–Q10) ---
title_cells = ["M4", "M7", "M11"]
info_text = [
    ["Last Updated At:"],
    [f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"],
    [""],
    ["Effective Health (EHP) Formula:"],
    ["Health x (1 + Armor ÷ 300)"],
    ["Shields and Overshields not included"],
    [""],
    ["Additional Info:"],
    ["We use the Median instead of the"],
    ["Average because extreme values"],
    ["(like Inaros with huge Health or"],
    ["Hildryn with massive Shields)"],
    ["distort the results. The Median"],
    ["represents the 'average Warframe'"],
    ["better."],
    [""],
    ["Copyright © TheNerdwork"]
]
worksheet.update(range_name="M4", values=info_text)

# --- ✅ Formatierung für Info-Block ---
info_range = worksheet.range("M4:M21")
for cell in info_range:
    cell_label = f"{col_to_letter(cell.col-1)}{cell.row}"
    # Prüfen, ob Zelle ein Titel ist
    if cell_label in title_cells:
        fmt = {
            "backgroundColor": base_bg,
            "wrapStrategy": "WRAP",
            "textFormat": {
                "foregroundColor": base_text,
                "fontSize": 16,
                "bold": True
            }
        }
    else:
        fmt = {
            "backgroundColor": base_bg,
            "wrapStrategy": "WRAP",
            "textFormat": {
                "foregroundColor": base_text,
                "fontSize": 16,
                "bold": False
            }
        }
    batch_formats.append({
        "range": cell_label,
        "format": fmt
    })

# --- Batch updates: 50 Zellen pro Request ---
for i in range(0, len(batch_formats), 50):
    worksheet.batch_format(batch_formats[i:i+50])

# --- ✅ Worksheet sinnvoll benennen ---
worksheet.update_title("Warframe Overview")

print(f"✅ Inserted {len(frame_rows)} Warframes with median row, info block, and formatting")