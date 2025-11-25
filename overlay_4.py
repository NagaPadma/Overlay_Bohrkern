import pandas as pd
import numpy as np
import folium
import pyproj
from pathlib import Path
from datetime import datetime

print("\n" + "="*70)
print("🗺️  German Boreholes - Multi-Layer Overlay Map")
print("="*70)

# CONFIG - YOUR DATASETS
DATASETS = [
    {
        "file": r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\25-03-19_Dateiverzeichnis_Bohrungen.xlsx",
        "sheet": "Geo_Koordinaten",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "All Points",
        "color": "blue"
    },
    {
        "file": r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\Bohrungen mit SVZ.xlsx",
        "sheet": "SVZ",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "SVZ",
        "color": "red"
    },
    {
        "file": r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\Bohrungen mit SVZ.xlsx",
        "sheet": "Log",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "Log",
        "color": "green"
    },
    {
        "file": r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\Bohrungen mit SVZ.xlsx",
        "sheet": "Bohrkern",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "Bohrkern",
        "color": "purple"
    }
]

OUTPUT_MAP = "German_Boreholes_Map.html"

print("\n📊 DATASETS TO LOAD:\n")
for i, ds in enumerate(DATASETS, 1):
    file_name = Path(ds['file']).name
    print(f"  {i}. {ds['name']}")
    print(f"     File: {file_name}")
    print(f"     Sheet: {ds['sheet']}")
    print(f"     Bohr ID: Column {ds['bohr_id_col']}, Coordinates: {ds['x_col']}, {ds['y_col']}")
    print(f"     Color: {ds['color']}\n")

# ==================== FUNCTION TO PROCESS FILE ====================
def process_excel_file(excel_file, sheet_name, bohr_id_col, x_col, y_col, dataset_name):
    """Read Excel and convert Gauß-Krüger to WGS84"""
    
    print(f"📁 {dataset_name}...", end='', flush=True)
    
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        print(f" ✗ {str(e)[:40]}")
        return None
    
    # Extract Bohr ID column
    try:
        if bohr_id_col in df.columns:
            bohr_column = bohr_id_col
        else:
            bohr_column = df.columns[int(ord(bohr_id_col) - ord('A'))]
        
        df['bohr_id'] = df[bohr_column].astype(str)
    except:
        df['bohr_id'] = 'N/A'
    
    # Extract X, Y
    try:
        if x_col in df.columns:
            x_column = x_col
        else:
            x_column = df.columns[int(ord(x_col) - ord('A'))]
        
        if y_col in df.columns:
            y_column = y_col
        else:
            y_column = df.columns[int(ord(y_col) - ord('A'))]
        
        df['X'] = pd.to_numeric(df[x_column], errors='coerce')
        df['Y'] = pd.to_numeric(df[y_column], errors='coerce')
        
        df_clean = df.dropna(subset=['X', 'Y'])
        
    except Exception as e:
        print(f" ✗ {str(e)[:40]}")
        return None
    
    # Detect zone
    x_mean = df_clean['X'].mean()
    if 2500000 < x_mean < 3500000:
        epsg_code = 31467  # Zone 3 (9°)
        zone = 3
    elif 3500000 < x_mean < 4500000:
        epsg_code = 31468  # Zone 4 (12°)
        zone = 4
    elif 1500000 < x_mean < 2500000:
        epsg_code = 31466  # Zone 2 (6°)
        zone = 2
    else:
        epsg_code = 31467
        zone = 3
    
    # Convert coordinates
    gk_crs = pyproj.CRS.from_epsg(epsg_code)
    wgs84_crs = pyproj.CRS.from_epsg(4326)
    transformer = pyproj.Transformer.from_crs(gk_crs, wgs84_crs, always_xy=True)
    
    lons, lats = transformer.transform(df_clean['X'].values, df_clean['Y'].values)
    
    df_clean['latitude'] = lats
    df_clean['longitude'] = lons
    
    print(f" ✓ {len(df_clean)} points")
    
    return df_clean

# ==================== PROCESS ALL FILES ====================
print("\n" + "="*70)
print("STEP 1: Reading and Converting Datasets")
print("="*70 + "\n")

datasets_loaded = []
for ds in DATASETS:
    df = process_excel_file(
        ds['file'], 
        ds['sheet'], 
        ds['bohr_id_col'],
        ds['x_col'], 
        ds['y_col'], 
        ds['name']
    )
    if df is not None:
        ds['data'] = df
        datasets_loaded.append(ds)
    else:
        print(f"  ⚠️  Skipping {ds['name']}")

if len(datasets_loaded) == 0:
    print("\n✗ No datasets loaded!")
    exit(1)

print(f"\n✓ Successfully loaded {len(datasets_loaded)} datasets")

# ==================== CREATE MAP ====================
print("\n" + "="*70)
print("STEP 2: Creating Overlay Map")
print("="*70)

# Calculate center from all datasets
all_lats = []
all_lons = []
for ds in datasets_loaded:
    all_lats.extend(ds['data']['latitude'].values)
    all_lons.extend(ds['data']['longitude'].values)

all_lats = np.array(all_lats)
all_lons = np.array(all_lons)

center_lat = all_lats.mean()
center_lon = all_lons.mean()

print(f"\n  Map center: {center_lat:.4f}°N, {center_lon:.4f}°E")
print(f"  Total points: {len(all_lats):,}")

m = folium.Map(
    location=[center_lat, center_lon],
    zoom_start=8,
    tiles='OpenStreetMap'
)

# Add satellite layer
folium.TileLayer(
    'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    attr='Esri',
    name='Satellite',
    overlay=False
).add_to(m)

# ==================== ADD ALL DATASETS ====================
print(f"\nAdding datasets to map...\n")

for ds in datasets_loaded:
    print(f"  {ds['name']} ({len(ds['data'])} points, {ds['color']})...", end='', flush=True)
    
    df = ds['data']
    color = ds['color']
    name = ds['name']
    
    fg = folium.FeatureGroup(name=f"{name} ({len(df)} pts)", show=True)
    
    for idx, row in df.iterrows():
        # Get bohr ID
        bohr_id = row.get('bohr_id', 'N/A')
        
        popup_text = f"""
        <b style="font-size: 14px; color: #{['0066CC', 'FF0000', '00AA00', '9933FF'][datasets_loaded.index(ds) % 4]};">{name}</b><br>
        <hr style="margin: 5px 0;">
        <table style="border-collapse: collapse; font-size: 12px;">
        <tr><td><b>Bohr ID:</b></td><td><b style="color: #333;">{bohr_id}</b></td></tr>
        <tr><td><b>Latitude:</b></td><td>{row['latitude']:.6f}°</td></tr>
        <tr><td><b>Longitude:</b></td><td>{row['longitude']:.6f}°</td></tr>
        <tr><td><b>GK X (m):</b></td><td>{row['X']:.0f}</td></tr>
        <tr><td><b>GK Y (m):</b></td><td>{row['Y']:.0f}</td></tr>
        """
        
        for col in df.columns:
            if col not in ['X', 'Y', 'latitude', 'longitude', 'bohr_id', ds['x_col'], ds['y_col'], ds['bohr_id_col']]:
                try:
                    val = row[col]
                    if pd.notna(val):
                        popup_text += f"<tr><td><b>{col}:</b></td><td>{val}</td></tr>"
                except:
                    pass
        
        popup_text += """
        </table>
        <hr style="margin: 5px 0;">
        <i style="font-size: 10px; color: #666;">Gauß-Krüger Zone 3 (9° meridian) → WGS84</i>
        """
        
        folium.Marker(
            location=[row['latitude'], row['longitude']],
            popup=folium.Popup(popup_text, max_width=350),
            icon=folium.Icon(color=color, icon='info-sign'),
            tooltip=f"<b>{bohr_id}</b> • {name}"
        ).add_to(fg)
    
    fg.add_to(m)
    print(" ✓")

# ==================== ADD LAYER CONTROL ====================
print(f"\n  Adding layer control...")
folium.LayerControl(position='topright', collapsed=False).add_to(m)

# ==================== ADD TITLE ====================
title_content = ""
colors_hex = ['0066CC', 'FF0000', '00AA00', '9933FF']
for i, ds in enumerate(datasets_loaded):
    color_hex = colors_hex[i % len(colors_hex)]
    title_content += f"<b style=\"color: #{color_hex}\">● {ds['name']}</b> ({len(ds['data'])} boreholes)<br>"

title_html = f'''
<div style="position: fixed; top: 10px; left: 50px; width: 450px; 
            background-color: white; border:3px solid #333; z-index:9999; 
            font-size:12px; padding: 12px; border-radius: 4px; box-shadow: 0 0 8px rgba(0,0,0,0.2);">
<b style="font-size: 15px;">⛏️ German Boreholes Map</b><br>
<b>Bohrungen Koordinaten Übersicht</b><br>
<hr style="margin: 5px 0;">
{title_content}
<hr style="margin: 5px 0;">
<b>Total Boreholes:</b> {len(all_lats):,}<br>
<b>Datasets:</b> {len(datasets_loaded)}<br>
<b>Region:</b> {center_lat:.4f}°N, {center_lon:.4f}°E<br>
<i style="color: #666; font-size: 11px;">Use layer control (→) to toggle datasets</i>
</div>
'''
m.get_root().html.add_child(folium.Element(title_html))

# ==================== ADD LEGEND ====================
legend_entries = ""
colors_hex = ['0066CC', 'FF0000', '00AA00', '9933FF']
for i, ds in enumerate(datasets_loaded):
    color_hex = colors_hex[i % len(colors_hex)]
    legend_entries += f'<i style="background: #{color_hex}; border-radius: 50%; display: inline-block; height: 12px; width: 12px; margin-right: 8px;"></i> {ds["name"]}<br>'

legend_html = f'''
<div style="position: fixed; bottom: 50px; right: 50px; width: 240px;
            background-color: white; border:3px solid #333; z-index:9999; 
            font-size:12px; padding: 10px; border-radius: 4px; box-shadow: 0 0 8px rgba(0,0,0,0.2);">
<b>Legend ({len(datasets_loaded)} Datasets)</b><br>
<hr style="margin: 5px 0;">
{legend_entries}
<hr style="margin: 5px 0;">
<i style="font-size: 11px; color: #666;">• Hover for Bohr ID<br>• Click for details</i>
</div>
'''
m.get_root().html.add_child(folium.Element(legend_html))

# ==================== SAVE MAP ====================
print(f"\nStep 3: Saving map...")

try:
    m.save(OUTPUT_MAP)
    file_size = Path(OUTPUT_MAP).stat().st_size / (1024*1024)
    print(f"  ✓ Saved: {OUTPUT_MAP}")
    print(f"  File size: {file_size:.2f} MB")
except Exception as e:
    print(f"  ✗ Error saving: {e}")
    exit(1)

# ==================== EXPORT COMBINED CSV ====================
print(f"\nStep 4: Saving combined coordinates...")

dfs_to_combine = []
for ds in datasets_loaded:
    df = ds['data'].copy()
    df['dataset'] = ds['name']
    dfs_to_combine.append(df[['bohr_id', 'X', 'Y', 'latitude', 'longitude', 'dataset']])

df_export = pd.concat(dfs_to_combine, ignore_index=True)

output_csv = OUTPUT_MAP.replace('.html', '_combined.csv')
df_export.to_csv(output_csv, index=False)
print(f"  ✓ Saved: {output_csv}")

# ==================== DONE ====================
print("\n" + "="*70)
print("✓ SUCCESS!")
print("="*70)
print(f"\nOutputs created:")
print(f"  1. {OUTPUT_MAP}")
print(f"  2. {output_csv}")
print(f"\nBoreholes loaded:")
for i, ds in enumerate(datasets_loaded, 1):
    print(f"  {i}. {ds['name']}: {len(ds['data'])} points ({ds['color']})")

print(f"\nTotal: {len(all_lats):,} boreholes")
print(f"\nOpen {OUTPUT_MAP} in your browser")
print("="*70 + "\n")