import pandas as pd
import folium
import os
import numpy as np
import re
import json
from folium import Element


# CONFIGURATION

# This ensures the script works inside your 'wind-asset-map' folder structure
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Define Data Path (looks in 'data' folder first)
p_source = os.path.join(DATA_DIR, "Merged data.xlsx")
if not os.path.exists(p_source):
    # Fallback to current folder just in case
    p_source = os.path.join(BASE_DIR, "Merged data.xlsx")

OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Interactive_Wind_Map.html")

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

SHEET_PARKS = "Windmill park clustering"
SHEET_TURBINES = "All active turbines"
SHEET_OWNERS = "Owner infomraiton"


# 1. HELPER FUNCTIONS

def clean_address(addr):
    if pd.isna(addr) or str(addr).strip() == 'nan': return "-"
    addr = str(addr).strip()
    addr = re.sub(r'(\d+)(?=\d{4}\b)', r'\1 ', addr)
    addr = re.sub(r'(\d+)(?=\d{4}\s)', r'\1 ', addr)
    addr = re.sub(r'(A/S|ApS)(?=[A-Za-z])', r'\1 ', addr)
    return addr

def format_phone(val):
    if pd.isna(val) or str(val).strip() == 'nan' or str(val).strip() == '-':
        return "none listed"
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    if not s.startswith('+'): s = "+" + s
    return s

def format_legal_owners(val):
    if pd.isna(val) or str(val).strip() == 'nan' or str(val).strip() == '-':
        return "none listed"
    text = str(val).strip()
    if len(text) < 50: return text
    truncated = text[:50] + "..."
    return f'<details style="cursor:pointer;outline:none;"><summary>{truncated}</summary><div style="margin-top:5px;padding:5px;background:#fff;border:1px solid #eee;">{text}</div></details>'

def get_historical_production(row, year_cols, month_cols_by_year):
    history = {}
    for y in year_cols:
        val = row.get(y, 0)
        if pd.notna(val) and val != 0: history[int(y)] = val
    for y, cols in month_cols_by_year.items():
        yearly_sum = 0
        has_data = False
        for c in cols:
            val = row.get(c, 0)
            if pd.notna(val):
                yearly_sum += val
                has_data = True
        if has_data and yearly_sum > 0: history[int(y)] = yearly_sum
    return history

def generate_dropdown_html_str(unique_id, history_dict, capacity_kw):
    if not history_dict: return "<span>No Data</span>", 0
    valid_years = [y for y in history_dict.keys() if 2020 <= y <= 2024]
    if not valid_years: return "<span>No Data (2020-24)</span>", 0
    years = sorted(valid_years, reverse=True)
    current_year = years[0]
    latest_prod = history_dict[current_year]
    
    options = ""
    for y in years:
        prod = history_dict[y]
        cf = prod / (capacity_kw * 8760) if capacity_kw > 0 else 0
        options += f"<option value='{prod:,.0f}|{cf:.1%}'>{y}</option>"
    
    select_html = f"""<select onchange="updatePopupData(this, 'prod_{unique_id}', 'cf_{unique_id}')" style="font-size:11px;padding:2px;">{options}</select>"""
    return select_html, latest_prod

def get_structures_list(row, structure_cols):
    found = []
    for col in structure_cols:
        try:
            val = float(row.get(col, 0))
            if val == 1.0:
                clean = col.replace("Type_", "").replace("_", " ")
                found.append(clean)
        except: continue
    if not found: return "none listed", []
    html = "<ul style='margin-top:2px;padding-left:20px;margin-bottom:0;color:#333;'>" + "".join([f"<li>{s}</li>" for s in found]) + "</ul>"
    return html, found

def format_email(email):
    if pd.isna(email) or str(email).lower() == 'nan' or str(email).strip() == '-': return "none listed"
    return f'<a href="mailto:{email}">{email}</a>'


# 2. LOAD & PREP DATA

print("Loading Data...")

if not os.path.exists(p_source):
    print(f"CRITICAL ERROR: Could not find file at: {p_source}")
    exit()

try:
    # A. LOAD PARKS
    df_parks = pd.read_excel(p_source, sheet_name=SHEET_PARKS, header=1)
    cluster_col = next((c for c in df_parks.columns if "Cluster ID" in str(c)), 'Cluster ID')
    df_parks = df_parks.rename(columns={
        cluster_col: 'Cluster_ID', 'Converted coordinates (Latitude)': 'Latitude', 'Converted coordinates (Longitude)': 'Longitude',
        'Total Capacity (kW)': 'Capacity_kW', 'Number of Turbines': 'Turbine_Count', 'Commune': 'Municipality', 'Location (Hav/Land)': 'Location'
    })

    # B. LOAD OWNERS
    df_owners = pd.read_excel(p_source, sheet_name=SHEET_OWNERS, header=1)
    df_owners.columns = df_owners.columns.str.strip()
    structure_cols = [c for c in df_owners.columns if str(c).startswith("Type_")]

    # Updated Financial Map (Supports English & Danish)
    fin_map = {
        'Resultat': ['Profit before tax', 'Resultat før skat', 'Resultat'],
        'Egenkapital': ['Equity', 'Egenkapital'],
        'Brutto': ['Gross profit', 'Bruttofortjeneste', 'Brutto'],
        'Gæld': ['Liabilities', 'Gældsforpligtelser', 'Gæld'],
        'Tilgode': ['Receivables', 'Tilgodehavender'],
        'Likvid': ['Liquidity ratio', 'Likviditetsgrad'],
        'Afkast': ['Return on assets', 'Afkastningsgrad'],
        'Solid': ['Equity ratio', 'Soliditetsgrad']
    }
    resolved_cols = {}
    for key, candidates in fin_map.items():
        for cand in candidates:
            if cand in df_owners.columns: resolved_cols[key] = cand; break
        if key not in resolved_cols: resolved_cols[key] = None

    # C. LOAD TURBINES
    df_raw = pd.read_excel(p_source, sheet_name=SHEET_TURBINES, header=0)

except Exception as e:
    print(f"ERROR LOADING SHEETS: {e}")
    exit()

print("Processing History...")
# Identify Year Columns
all_cols = df_raw.columns.tolist()
year_cols = [str(c) for c in all_cols if re.match(r'^\d{4}$', str(c))]
month_cols_by_year = {}
for c in all_cols:
    if re.match(r'^\d{4}-\d{2}$', str(c)):
        y = str(c)[:4]
        if y not in month_cols_by_year: month_cols_by_year[y] = []
        month_cols_by_year[y].append(c)

df_raw['History'] = df_raw.apply(lambda row: get_historical_production(row, year_cols, month_cols_by_year), axis=1)

# Clean Data
if 'Dato for afmeldning' in df_raw.columns:
    df_raw = df_raw[df_raw['Dato for afmeldning'].isna()].copy()
df_raw = df_raw.dropna(subset=['Latitude', 'Longitude'])
df_raw['Cluster_ID'] = pd.to_numeric(df_raw['Cluster_ID'], errors='coerce').fillna(-1).astype(int)
df_parks['Cluster_ID'] = pd.to_numeric(df_parks['Cluster_ID'], errors='coerce').fillna(-1).astype(int)

# Map park sizes for filters
park_counts = df_parks.set_index('Cluster_ID')['Turbine_Count'].to_dict()

# Aggregate Data
park_histories = {}
park_compositions = {}
for pid in df_parks['Cluster_ID'].unique():
    if pid == -1: continue
    turbines_in_cluster = df_raw[df_raw['Cluster_ID'] == pid]
    
    # History
    cluster_history = {}
    for _, turb in turbines_in_cluster.iterrows():
        th = turb['History']
        for year, val in th.items():
            cluster_history[year] = cluster_history.get(year, 0) + val
    park_histories[pid] = cluster_history
    
    # Composition
    comp = turbines_in_cluster.groupby(['Fabrikat', 'Typebetegnelse', 'Kapacitet (kW)']).size().reset_index(name='Count')
    comp = comp.sort_values('Count', ascending=False)
    comp_html = "<ul style='margin-top:2px;padding-left:20px;margin-bottom:0;color:#333;'>"
    if comp.empty: comp_html += "<li>No turbine data</li>"
    else:
        for _, row in comp.iterrows():
            make = str(row['Fabrikat']).replace(" Wind Systems A/S", "").replace(" A/S", "")
            model = str(row['Typebetegnelse'])
            cap = row['Kapacitet (kW)']
            comp_html += f"<li><b>{row['Count']}x</b> {make} {model} ({cap} kW)</li>"
    comp_html += "</ul>"
    park_compositions[pid] = comp_html


# 3. BUILD JSON DATA

print("Building Application Data...")

# 1. TURBINES
turbines_data = []
for idx, row in df_raw.iterrows():
    cap = row.get('Kapacitet (kW)', 0)
    pid = row.get('Cluster_ID', -1)
    
    # Single Turbine Logic: If park size <= 1 or no park ID
    cluster_size = park_counts.get(pid, 1) if pid != -1 else 1
    
    is_offshore = "HAV" in str(row.get('Type af placering', '')).upper()
    color = "#005a7d" if is_offshore else "#3b5c00"
    
    unique_id = f"tb{idx}"
    dropdown, latest_prod = generate_dropdown_html_str(unique_id, row['History'], cap)
    latest_cf = (latest_prod / (cap * 8760)) if (cap > 0 and latest_prod > 0) else 0
    
    popup = f"""<div style="font-size:12px;width:240px;font-family:sans-serif;">
        <b style="font-size:13px;">{row.get('Fabrikat', '')} {row.get('Typebetegnelse', '')}</b><br>
        <span style="color:#666;">Capacity: {cap} kW</span>
        <hr style="margin:8px 0;border:0;border-top:1px solid #eee;">
        <table style="width:100%;">
            <tr><td>Year:</td><td align="right">{dropdown}</td></tr>
            <tr><td>Productivity:</td><td align="right"><b id="prod_{unique_id}">{latest_prod:,.0f} kWh</b></td></tr>
            <tr><td>Capacity Factor:</td><td align="right"><span id="cf_{unique_id}">{latest_cf:.1%}</span></td></tr>
        </table></div>"""
        
    turbines_data.append({
        "lat": row['Latitude'], "lon": row['Longitude'],
        "cap": cap, "size": cluster_size, "color": color, "popup": popup
    })

# 2. PARKS
parks_data = []
for idx, row in df_parks.iterrows():
    if pd.isna(row['Latitude']): continue
    
    # Filter: Don't show park markers for single turbines
    t_count = row.get('Turbine_Count', 0)
    if t_count <= 1: continue 
    
    cap = row.get('Capacity_kW', 0)
    pid = row.get('Cluster_ID', -1)
    is_offshore = "HAV" in str(row.get('Location', '')).upper()
    color = "#0078A8" if is_offshore else "#568203"
    radius = ((cap/1000) ** 0.5) * 1.5 
    radius = max(radius, 5)
    
    unique_id = f"pk{idx}"
    history = park_histories.get(pid, {})
    dropdown, latest_prod = generate_dropdown_html_str(unique_id, history, cap)
    latest_cf = (latest_prod / (cap * 8760)) if (cap > 0 and latest_prod > 0) else 0
    comp_html = park_compositions.get(pid, "<li>No Data</li>")
    
    popup = f"""<div style="width:320px;font-family:sans-serif;color:#333;">
        <h4 style="margin:0 0 5px 0;color:{color};border-bottom:1px solid #ccc;">{row.get('Municipality', 'Unknown')} <span style="font-size:12px;color:#666;">(ID: {pid})</span></h4>
        <table style="width:100%;font-size:12px;margin-bottom:8px;">
            <tr><td><b>Turbines:</b></td><td align="right">{t_count}</td></tr>
            <tr><td><b>Capacity:</b></td><td align="right">{cap:,.0f} kW</td></tr>
            <tr><td><b>Year:</b></td><td align="right">{dropdown}</td></tr>
            <tr><td><b>Productivity:</b></td><td align="right"><b id="prod_{unique_id}">{latest_prod:,.0f} kWh</b></td></tr>
            <tr><td><b>Capacity Factor:</b></td><td align="right"><span id="cf_{unique_id}">{latest_cf:.1%}</span></td></tr>
        </table>
        <div style="font-size:12px;border-top:1px solid #eee;padding-top:5px;"><b>Turbine Composition:</b>{comp_html}</div></div>"""
    
    parks_data.append({
        "lat": row['Latitude'], "lon": row['Longitude'],
        "cap": cap, "color": color, "radius": radius, "popup": popup
    })

# 3. OWNERS
owners_data = []
all_structures_set = set()

for idx, row in df_owners.iterrows():
    if pd.isna(row.get('Latitude')): continue
    structures_html, structure_list = get_structures_list(row, structure_cols)
    for s in structure_list: all_structures_set.add(s)
    
    name = row.get('Company Name', 'Unknown')
    link = row.get('Link', '#')
    cvr = row.get('CVR', '')
    comp_addr = clean_address(row.get('Address', '-'))
    prop_addr = clean_address(row.get('Property Address', '-'))
    email_html = format_email(row.get('Email', '-'))
    phone_val = format_phone(row.get('Phone', '-'))
    
    directors = str(row.get('Directors', '-')).replace(';', '<br>')
    if pd.isna(row.get('Directors', '-')) or str(row.get('Directors', '-')) == 'nan': 
        directors = "none listed"
        
    legal_owners = format_legal_owners(row.get('Legal Owners', '-'))
    
    prop_link = row.get('Specific Property Link', '')
    prop_link_btn = ""
    if pd.notna(prop_link) and str(prop_link).startswith('http'):
        prop_link_btn = f"""<div style="margin-top:5px;"><a href="{prop_link}" target="_blank" style="background:#555;color:white;padding:3px 8px;text-decoration:none;border-radius:3px;font-size:10px;">View Property ➔</a></div>"""

    def get_val(key):
        col = resolved_cols.get(key)
        val = row.get(col, '-') if col else '-'
        return val if pd.notna(val) else '-'

    popup = f"""<div style="font-family:sans-serif;width:320px;">
        <h3 style="margin:0 0 5px 0;color:#D32F2F;border-bottom:2px solid #D32F2F;font-size:16px;">{name}</h3>
        <div style="font-size:11px;font-weight:bold;color:#666;margin-bottom:15px;">CVR: {cvr}</div>
        <div style="max-height:250px;overflow-y:auto;padding-right:5px;">
            <div style="margin-bottom:15px;font-size:12px;">
                <div style="margin-bottom:8px;"><b style="color:#333;">Company Address:</b><br><span style="color:#555;">{comp_addr}</span></div>
                <div style="margin-bottom:8px;"><b style="color:#333;">Property Address:</b><br><span style="color:#555;">{prop_addr}</span>{prop_link_btn}</div>
                <div style="margin-top:10px;padding-top:8px;border-top:1px solid #eee;"><b>Email:</b> {email_html}<br><b>Phone:</b> {phone_val}</div>
            </div>
            <div style="font-size:12px;margin-bottom:15px;border-top:1px solid #eee;padding-top:10px;"><b style="color:#333;">Directors:</b><br><span style="color:#555;">{directors}</span></div>
            <div style="font-size:12px;margin-bottom:15px;border-top:1px solid #eee;padding-top:10px;"><b style="color:#333;">Legal Owners:</b><br><div style="color:#555;">{legal_owners}</div></div>
            <div style="margin-bottom:15px;border-top:1px solid #eee;padding-top:10px;"><b style="color:#333;font-size:12px;">Structures on Property:</b><div style="font-size:11px;">{structures_html}</div></div>
            <div style="margin-bottom:10px;border-top:1px solid #eee;padding-top:10px;"><b style="color:#333;font-size:12px;">Financials:</b>
            <table style="width:100%;font-size:11px;margin-top:5px;border-collapse:collapse;">
                <tr style="border-bottom:1px solid #eee;"><td>Pre-tax:</td><td align="right">{get_val('Resultat')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>Equity:</td><td align="right">{get_val('Egenkapital')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>Gross Profit:</td><td align="right">{get_val('Brutto')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>Liabilities:</td><td align="right">{get_val('Gæld')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>Receivables:</td><td align="right">{get_val('Tilgode')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>Liquidity:</td><td align="right">{get_val('Likvid')}</td></tr>
                <tr style="border-bottom:1px solid #eee;"><td>ROA:</td><td align="right">{get_val('Afkast')}</td></tr>
                <tr><td>Solvency:</td><td align="right">{get_val('Solid')}</td></tr>
            </table></div>
            <div style="text-align:center;margin-top:15px;"><a href="{link}" target="_blank" style="background:#D32F2F;color:white;padding:5px 10px;text-decoration:none;border-radius:3px;font-size:12px;">Full Company Profile</a></div>
        </div></div>"""
        
    owners_data.append({
        "lat": row['Latitude'], "lon": row['Longitude'],
        "structs": structure_list, "popup": popup
    })

ordered_structures = []
for c in structure_cols:
    clean = c.replace("Type_", "").replace("_", " ")
    if clean in all_structures_set:
        ordered_structures.append(clean)


# 4. GENERATE MAP HTML

print("Generating Application...")

m = folium.Map(location=[56.2, 10.5], zoom_start=7, tiles=None)

data_js = f"""
<script>
    var turbineData = {json.dumps(turbines_data)};
    var parkData = {json.dumps(parks_data)};
    var ownerData = {json.dumps(owners_data)};
    var structureTypes = {json.dumps(ordered_structures)};
</script>
"""

ui_html = """
<style>
    /* DARK MAP CONTROLS - MATCH THE DATA FILTER HEADER */
    .leaflet-bar,
    .leaflet-control-layers {
        border:0 !important;
        border-radius:8px !important;
        box-shadow:0 0 15px rgba(0,0,0,0.3) !important;
        overflow:hidden;
    }
    .leaflet-bar a,
    .leaflet-control-layers,
    .leaflet-control-layers-expanded {
        background:#333 !important;
        color:#fff !important;
    }
    .leaflet-bar a {
        border-bottom:1px solid #4a4a4a !important;
    }
    .leaflet-bar a:last-child { border-bottom:0 !important; }
    .leaflet-bar a:hover { background:#444 !important; color:#fff !important; }
    .leaflet-bar a.leaflet-disabled { background:#292929 !important; color:#777 !important; }
    .leaflet-control-layers label { color:#fff; }
    .leaflet-control-layers input { accent-color:#777; }
    .leaflet-control-layers-separator { border-top-color:#555; }

    /* PANEL STYLE */
    .filter-panel { 
        position:fixed; bottom:30px; right:30px; width:340px; 
        background:#ccc; /* Gray background to show gap */
        border-radius:8px; box-shadow:0 0 15px rgba(0,0,0,0.3); z-index:9999; 
        font-family:sans-serif; font-size:12px; transition: height 0.3s; overflow:hidden; 
    }
    .filter-header { 
        background:#333; color:white; padding:12px 15px; cursor:pointer; font-weight:bold; 
        display:flex; justify-content:space-between; align-items:center; 
    }
    .filter-content { 
        padding:0; max-height:85vh; overflow-y:auto; display:none; background:#ccc;
    }
    .filter-open .filter-content { display:block; }
    
    /* SECTION CARDS - ACCORDION STYLE */
    .filter-section { 
        background:white; margin-bottom:15px; /* THICK GAP */
        padding:0; 
    }
    .filter-section:last-child { margin-bottom:0; }
    
    .section-header {
        padding:12px 15px; cursor:pointer; background:#fff;
        display:flex; justify-content:space-between; align-items:center;
    }
    .section-header:hover { background:#f9f9f9; }
    .section-title { font-weight:bold; text-transform:uppercase; font-size:12px; letter-spacing:0.5px; margin:0; }
    
    /* ICONS */
    .section-icon { display:inline-block; vertical-align:middle; margin-right:8px; }
    .icon-turbine { width:10px; height:10px; background:#3b5c00; border-radius:50%; }
    
    /* UPDATED ICONS */
    .icon-park { width:12px; height:12px; border:2px solid #568203; border-radius:50%; display:inline-block; }
    /* Blue Park Icon REMOVED as requested */
    
    /* UPDATED OWNER ICON (Red Hollow Ring) */
    .icon-owner { width:12px; height:12px; border:2px solid #D32F2F; border-radius:50%; display:inline-block; }
    
    /* BODY (Closed by default) */
    .section-body { padding:15px; display:none; border-top:1px solid #eee; }
    .section-expanded .section-body { display:block; }
    .section-expanded .arrow { transform: rotate(180deg); }
    
    .struct-list { border:1px solid #ccc; max-height:150px; overflow-y:auto; border-radius:4px; margin-top:5px; }
    .struct-item { padding:5px 8px; border-bottom:1px solid #eee; cursor:pointer; display:flex; align-items:center; }
    .struct-item:hover { background:#f9f9f9; }
    .struct-item.selected { color:#2E7D32; font-weight:bold; background:#E8F5E9; } 
    .struct-item .tick { width:15px; display:inline-block; font-weight:bold; }
    
    /* SLIDERS */
    .dual-slider-container { position:relative; width:100%; height:35px; margin-top:10px; }
    .dual-slider-track { position:absolute; top:15px; left:0; right:0; height:6px; background:#ddd; border-radius:3px; z-index:1; }
    .dual-slider-fill { position:absolute; top:15px; height:6px; background:#555; z-index:1; border-radius:3px; }
    .dual-slider-input { position:absolute; pointer-events:none; -webkit-appearance:none; z-index:2; height:6px; width:100%; background:transparent; top:12px; }
    
    .dual-slider-input::-webkit-slider-thumb { 
        pointer-events:all; width:20px; height:20px; -webkit-appearance:none; 
        background:#333; border:2px solid white; border-radius:50%; cursor:pointer; 
        box-shadow:0 1px 4px rgba(0,0,0,0.4); margin-top:0px; 
    }
    .dual-slider-input::-moz-range-thumb {
        pointer-events:all; width:20px; height:20px; background:#333; 
        border:2px solid white; border-radius:50%; cursor:pointer; 
        box-shadow:0 1px 4px rgba(0,0,0,0.4); 
    }
    
    .input-row { display:flex; justify-content:space-between; gap:10px; margin-bottom:5px; }
    .input-group { flex:1; }
    .input-group label { display:block; font-size:9px; color:#666; margin-bottom:2px; }
    .input-group input { width:90%; padding:4px; border:1px solid #ccc; border-radius:3px; font-size:11px; }
</style>

<div id="filter-panel" class="filter-panel">
    <div class="filter-header" onclick="toggleFilter()">
        <span>Data Filter</span>
        <span id="toggle-icon">▲</span>
    </div>
    <div class="filter-content">
        
        <div class="filter-section">
            <div class="section-header" onclick="toggleSection(this)">
                <div class="section-title" style="color:#3b5c00;">
                    <span class="section-icon icon-turbine"></span>Individual Turbines
                </div>
                <span class="arrow">▼</span>
            </div>
            <div class="section-body">
                <label style="display:flex; align-items:center; margin-bottom:12px; font-weight:bold;">
                    <input type="checkbox" id="chk-turbines" checked onchange="renderLayers()"> Show Turbines
                </label>
                
                <div class="input-row">
                    <div class="input-group">
                        <label>Min (kW)</label>
                        <input type="number" id="num-t-min" min="0" max="10000" step="25" value="0" onchange="syncSlider('t', 'min')">
                    </div>
                    <div class="input-group">
                        <label>Max (kW)</label>
                        <input type="number" id="num-t-max" min="0" max="10000" step="25" value="10000" onchange="syncSlider('t', 'max')">
                    </div>
                </div>
                
                <div class="dual-slider-container">
                    <div class="dual-slider-track"></div>
                    <div class="dual-slider-fill" id="fill-t" style="left:0%; width:100%;"></div>
                    <input type="range" class="dual-slider-input" id="rng-t-min" min="0" max="10000" step="25" value="0" oninput="updateDualSlider('t'); renderLayers()">
                    <input type="range" class="dual-slider-input" id="rng-t-max" min="0" max="10000" step="25" value="10000" oninput="updateDualSlider('t'); renderLayers()">
                </div>
                
                <label style="display:flex; align-items:center; margin-top:20px; background:#f0f8f0; padding:8px; border-radius:4px;">
                    <input type="checkbox" id="chk-turb-orphan" onchange="renderLayers()"> 
                    <span style="margin-left:8px;">Only Single-Turbine Parks</span>
                </label>
            </div>
        </div>

        <div class="filter-section">
            <div class="section-header" onclick="toggleSection(this)">
                <div class="section-title" style="color:#2E7D32;">
                    <span class="section-icon icon-park"></span>Wind Parks
                </div>
                <span class="arrow">▼</span>
            </div>
            <div class="section-body">
                <label style="display:flex; align-items:center; margin-bottom:12px; font-weight:bold;">
                    <input type="checkbox" id="chk-parks" checked onchange="renderLayers()"> Show Parks
                </label>
                
                <div class="input-row">
                    <div class="input-group">
                        <label>Min (kW)</label>
                        <input type="number" id="num-p-min" min="0" max="400000" step="1000" value="0" onchange="syncSlider('p', 'min')">
                    </div>
                    <div class="input-group">
                        <label>Max (kW)</label>
                        <input type="number" id="num-p-max" min="0" max="400000" step="1000" value="400000" onchange="syncSlider('p', 'max')">
                    </div>
                </div>
                
                <div class="dual-slider-container">
                    <div class="dual-slider-track"></div>
                    <div class="dual-slider-fill" id="fill-p" style="left:0%; width:100%;"></div>
                    <input type="range" class="dual-slider-input" id="rng-p-min" min="0" max="400000" step="1000" value="0" oninput="updateDualSlider('p'); renderLayers()">
                    <input type="range" class="dual-slider-input" id="rng-p-max" min="0" max="400000" step="1000" value="400000" oninput="updateDualSlider('p'); renderLayers()">
                </div>
            </div>
        </div>

        <div class="filter-section">
            <div class="section-header" onclick="toggleSection(this)">
                <div class="section-title" style="color:#D32F2F;">
                    <span class="section-icon icon-owner"></span>Owner Information
                </div>
                <span class="arrow">▼</span>
            </div>
            <div class="section-body">
                <label style="display:flex; align-items:center; margin-bottom:12px; font-weight:bold;">
                    <input type="checkbox" id="chk-owners" checked onchange="renderLayers()"> Show Owners
                </label>
                
                <div style="font-size:11px; margin-bottom:5px;">Filter by Structures:</div>
                <div id="struct-list" class="struct-list">
                    </div>
            </div>
        </div>
        
    </div>
</div>
"""

app_js = """
<script>
    var layerTurbines = L.layerGroup();
    var layerParks = L.layerGroup();
    var layerOwners = L.layerGroup();
    var selectedStructures = new Set();
    
    window.updatePopupData = function(selectObj, prodId, cfId) {
        var parts = selectObj.value.split('|');
        document.getElementById(prodId).innerText = parts[0] + ' kWh';
        document.getElementById(cfId).innerText = parts[1];
    };

    function toggleFilter() {
        var p = document.getElementById('filter-panel');
        var icon = document.getElementById('toggle-icon');
        p.classList.toggle('filter-open');
        icon.innerText = p.classList.contains('filter-open') ? '▼' : '▲';
    }
    
    function toggleSection(header) {
        var section = header.parentElement;
        section.classList.toggle('section-expanded');
    }
    
    function syncSlider(prefix, type) {
        var numInput = document.getElementById('num-' + prefix + '-' + type);
        var slider = document.getElementById('rng-' + prefix + '-' + type);
        var val = parseInt(numInput.value);
        
        if(val < parseInt(slider.min)) val = parseInt(slider.min);
        
        // LIMIT OVERRIDE LOGIC
        // If prefix is 't' (turbines) and value > 10000, keep it (it means user wants high value)
        // If prefix is 'p' (parks) and value > 400000, clamp it? No, allow user freedom.
        
        slider.value = val;
        updateDualSlider(prefix);
        renderLayers();
    }
    
    function updateDualSlider(prefix) {
        var minR = document.getElementById('rng-' + prefix + '-min');
        var maxR = document.getElementById('rng-' + prefix + '-max');
        var minNum = document.getElementById('num-' + prefix + '-min');
        var maxNum = document.getElementById('num-' + prefix + '-max');
        
        var minV = parseInt(minR.value);
        var maxV = parseInt(maxR.value);
        
        if (minV > maxV) {
            if (document.activeElement === minR) { minR.value = maxV; minV = maxV; } 
            else { maxR.value = minV; maxV = minV; }
        }
        
        minNum.value = minV;
        maxNum.value = maxV;
        
        var maxAttr = parseInt(minR.max);
        var left = (minV / maxAttr) * 100;
        var width = ((maxV - minV) / maxAttr) * 100;
        var fill = document.getElementById('fill-' + prefix);
        fill.style.left = left + "%";
        fill.style.width = width + "%";
    }

    function toggleStruct(elem, structName) {
        if (selectedStructures.has(structName)) {
            selectedStructures.delete(structName);
            elem.classList.remove('selected');
            elem.querySelector('.tick').innerText = '';
        } else {
            selectedStructures.add(structName);
            elem.classList.add('selected');
            elem.querySelector('.tick').innerText = '☑ ';
        }
        renderLayers();
    }

    function initApp() {
        var dark = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {subdomains: 'abcd', attribution: '&copy; CARTO'});
        var osm = L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {attribution: '&copy; OpenStreetMap'});
        var sat = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {attribution: 'Esri'});

        dark.addTo(map);

        var listContainer = document.getElementById('struct-list');
        structureTypes.forEach(function(s) {
            var div = document.createElement('div');
            div.className = 'struct-item';
            div.innerHTML = '<span class="tick"></span> ' + s;
            div.onclick = function() { toggleStruct(this, s); };
            listContainer.appendChild(div);
        });
        
        var baseMaps = { "Dark Mode": dark, "Light mode": osm, "Satellite": sat };
        L.control.layers(baseMaps, null, {collapsed:false, position:'topright'}).addTo(map);

        layerParks.addTo(map);
        layerOwners.addTo(map);
        layerTurbines.addTo(map);

        renderLayers();
    }

    function renderLayers() {
        // 1. TURBINES
        layerTurbines.clearLayers();
        if (document.getElementById('chk-turbines').checked) {
            var tMin = parseInt(document.getElementById('rng-t-min').value);
            var tMax = parseInt(document.getElementById('rng-t-max').value);
            
            // SMART LOGIC: If slider max is at limit (10,000), treat as infinity
            var filterMax = (tMax === 10000) ? 99999999 : tMax;
            
            var onlySingle = document.getElementById('chk-turb-orphan').checked;
            
            turbineData.forEach(function(d) {
                if (d.cap < tMin || d.cap > filterMax) return;
                if (onlySingle && d.size > 1) return;
                
                L.circleMarker([d.lat, d.lon], {
                    radius: 2, color: d.color, weight: 0, fill: true, fillOpacity: 0.9
                }).bindPopup(d.popup).addTo(layerTurbines);
            });
        }

        // 2. PARKS
        layerParks.clearLayers();
        if (document.getElementById('chk-parks').checked) {
            var pMin = parseInt(document.getElementById('rng-p-min').value);
            var pMax = parseInt(document.getElementById('rng-p-max').value);
            
            parkData.forEach(function(d) {
                if (d.cap < pMin || d.cap > pMax) return;
                L.circleMarker([d.lat, d.lon], {
                    radius: d.radius, color: d.color, weight: 1, fill: true, fillOpacity: 0.2
                }).bindPopup(d.popup).addTo(layerParks);
            });
        }

        // 3. OWNERS
        layerOwners.clearLayers();
        if (document.getElementById('chk-owners').checked) {
            ownerData.forEach(function(d) {
                if (selectedStructures.size > 0) {
                    var matchCount = 0;
                    selectedStructures.forEach(function(s) {
                        if (d.structs.includes(s)) matchCount++;
                    });
                    if (matchCount < selectedStructures.size) return;
                }
                
                L.circleMarker([d.lat, d.lon], {
                    radius: 6, color: "#D32F2F", weight: 2, fill: true, fillColor: "#FFCDD2", fillOpacity: 0.4
                }).bindPopup(d.popup).addTo(layerOwners);
            });
        }
    }

    window.addEventListener('load', function() {
        var checkMap = setInterval(function() {
            for(var key in window) {
                if (window[key] instanceof L.Map) {
                    window.map = window[key];
                    clearInterval(checkMap);
                    initApp();
                    break;
                }
            }
        }, 100);
    });
</script>
"""

m.get_root().html.add_child(Element(data_js))
m.get_root().html.add_child(Element(ui_html))
m.get_root().html.add_child(Element(app_js))

print(f"Saving to: {OUTPUT_FILE}")
m.save(OUTPUT_FILE)
print("Done!")
