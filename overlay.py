import pandas as pd
import numpy as np
import folium
import pyproj
from pathlib import Path
from datetime import datetime

print("\n" + "="*70)
print("🗺️  DUAL EXCEL Gauß-Krüger Map Overlay")
print("="*70)

# CONFIG - UPDATE THESE
EXCEL_FILE_1 = r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\25-03-19_Dateiverzeichnis_Bohrungen.xlsx"  # First file
SHEET_NAME_1 = "Geo_Koordinaten"
X_COLUMN_1 = "C"
Y_COLUMN_1 = "D"
LAYER_NAME_1 = "All Points"
COLOR_1 = "blue"  # Color for first dataset

EXCEL_FILE_2 = r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\Bohrungen mit SVZ.xlsx"  # Second file
SHEET_NAME_2 = "SVZ"
X_COLUMN_2 = "C"
Y_COLUMN_2 = "D"
LAYER_NAME_2 = "SVZ"
COLOR_2 = "red"  # Color for second dataset

OUTPUT_MAP = "coordinates_overlay_map.html"

# Available colors: red, blue, green, purple, orange, darkred, darkblue, darkgreen, cadetblue, darkpurple, white, pink, lightblue, lightgreen, gray, black, lightgray

print(f"\n📊 All Points")
print(f"  File: {EXCEL_FILE_1}")
print(f"  Sheet: {SHEET_NAME_1}")
print(f"  Columns: {X_COLUMN_1}, {Y_COLUMN_1}")
print(f"  Color: {COLOR_1}")

print(f"\n📊 SVZ")
print(f"  File: {EXCEL_FILE_2}")
print(f"  Sheet: {SHEET_NAME_2}")
print(f"  Columns: {X_COLUMN_2}, {Y_COLUMN_2}")
print(f"  Color: {COLOR_2}")

# ==================== FUNCTION TO PROCESS FILE ====================
def process_excel_file(excel_file, sheet_name, x_col, y_col, dataset_name):
    """Read Excel and convert Gauß-Krüger to WGS84"""
    
    print(f"\n📁 Reading {dataset_name}...")
    
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        print(f"  ✓ Loaded {len(df)} rows")
    except Exception as e:
        print(f"  ✗ Error reading file: {e}")
        return None
    
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
        
        print(f"  ✓ Found {len(df_clean)} valid coordinates")
        print(f"  X range: {df_clean['X'].min():.0f} to {df_clean['X'].max():.0f}")
        print(f"  Y range: {df_clean['Y'].min():.0f} to {df_clean['Y'].max():.0f}")
        
    except Exception as e:
        print(f"  ✗ Error extracting coordinates: {e}")
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
    
    print(f"  Detected: Gauß-Krüger Zone {zone} (EPSG:{epsg_code})")
    
    # Convert coordinates
    print(f"  Converting {len(df_clean)} coordinates...")
    
    gk_crs = pyproj.CRS.from_epsg(epsg_code)
    wgs84_crs = pyproj.CRS.from_epsg(4326)
    transformer = pyproj.Transformer.from_crs(gk_crs, wgs84_crs, always_xy=True)
    
    lons, lats = transformer.transform(df_clean['X'].values, df_clean['Y'].values)
    
    df_clean['latitude'] = lats
    df_clean['longitude'] = lons
    
    print(f"  ✓ Conversion complete")
    print(f"  Lat range: {lats.min():.4f} to {lats.max():.4f}")
    print(f"  Lon range: {lons.min():.4f} to {lons.max():.4f}")
    
    return df_clean, epsg_code

# ==================== PROCESS BOTH FILES ====================
print("\n" + "="*70)
print("STEP 1: Reading and Converting Datasets")
print("="*70)

df1, epsg1 = process_excel_file(EXCEL_FILE_1, SHEET_NAME_1, X_COLUMN_1, Y_COLUMN_1, LAYER_NAME_1)
if df1 is None:
    exit(1)

df2, epsg2 = process_excel_file(EXCEL_FILE_2, SHEET_NAME_2, X_COLUMN_2, Y_COLUMN_2, LAYER_NAME_2)
if df2 is None:
    exit(1)

# ==================== CREATE MAP ====================
print("\n" + "="*70)
print("STEP 2: Creating Overlay Map")
print("="*70)

# Calculate center from both datasets
all_lats = np.concatenate([df1['latitude'].values, df2['latitude'].values])
all_lons = np.concatenate([df1['longitude'].values, df2['longitude'].values])

center_lat = all_lats.mean()
center_lon = all_lons.mean()

print(f"\n  Map center: {center_lat:.4f}°N, {center_lon:.4f}°E")

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

# ==================== ADD DATASET 1 ====================
print(f"\nAdding {LAYER_NAME_1} ({len(df1)} points, color: {COLOR_1})...")

fg1 = folium.FeatureGroup(name=f"{LAYER_NAME_1} ({len(df1)} pts)", show=True)

for idx, row in df1.iterrows():
    popup_text = f"""
    <b>{LAYER_NAME_1}</b><br>
    <table style="border-collapse: collapse; font-size: 12px;">
    <tr><td><b>Lat:</b></td><td>{row['latitude']:.6f}°</td></tr>
    <tr><td><b>Lon:</b></td><td>{row['longitude']:.6f}°</td></tr>
    <tr><td><b>GK X:</b></td><td>{row['X']:.0f}</td></tr>
    <tr><td><b>GK Y:</b></td><td>{row['Y']:.0f}</td></tr>
    """
    
    for col in df1.columns:
        if col not in ['X', 'Y', 'latitude', 'longitude', X_COLUMN_1, Y_COLUMN_1]:
            try:
                popup_text += f"<tr><td><b>{col}:</b></td><td>{row[col]}</td></tr>"
            except:
                pass
    
    popup_text += "</table>"
    
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=folium.Popup(popup_text, max_width=300),
        icon=folium.Icon(color=COLOR_1, icon='info-sign'),
        tooltip=f"{LAYER_NAME_1}: ({row['latitude']:.4f}, {row['longitude']:.4f})"
    ).add_to(fg1)

fg1.add_to(m)

# ==================== ADD DATASET 2 ====================
print(f"Adding {LAYER_NAME_2} ({len(df2)} points, color: {COLOR_2})...")

fg2 = folium.FeatureGroup(name=f"{LAYER_NAME_2} ({len(df2)} pts)", show=True)

for idx, row in df2.iterrows():
    popup_text = f"""
    <b>{LAYER_NAME_2}</b><br>
    <table style="border-collapse: collapse; font-size: 12px;">
    <tr><td><b>Lat:</b></td><td>{row['latitude']:.6f}°</td></tr>
    <tr><td><b>Lon:</b></td><td>{row['longitude']:.6f}°</td></tr>
    <tr><td><b>GK X:</b></td><td>{row['X']:.0f}</td></tr>
    <tr><td><b>GK Y:</b></td><td>{row['Y']:.0f}</td></tr>
    """
    
    for col in df2.columns:
        if col not in ['X', 'Y', 'latitude', 'longitude', X_COLUMN_2, Y_COLUMN_2]:
            try:
                popup_text += f"<tr><td><b>{col}:</b></td><td>{row[col]}</td></tr>"
            except:
                pass
    
    popup_text += "</table>"
    
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=folium.Popup(popup_text, max_width=300),
        icon=folium.Icon(color=COLOR_2, icon='info-sign'),
        tooltip=f"{LAYER_NAME_2}: ({row['latitude']:.4f}, {row['longitude']:.4f})"
    ).add_to(fg2)

fg2.add_to(m)

# ==================== ADD LAYER CONTROL ====================
print(f"Adding layer control...")
folium.LayerControl(position='topright', collapsed=False).add_to(m)

# ==================== ADD TITLE ====================
title_html = f'''
<div style="position: fixed; top: 10px; left: 50px; width: 420px; 
            background-color: white; border:3px solid #333; z-index:9999; 
            font-size:12px; padding: 12px; border-radius: 4px; box-shadow: 0 0 8px rgba(0,0,0,0.2);">
<b style="font-size: 14px;">🗺️ Overlay Map - Two Datasets</b><br>
<hr style="margin: 5px 0;">
<b style="color: #0066CC;">● {LAYER_NAME_1}</b><br>
   Points: {len(df1)}<br>
<b style="color: #FF0000;">● {LAYER_NAME_2}</b><br>
   Points: {len(df2)}<br>
<hr style="margin: 5px 0;">
<b>Total Points:</b> {len(df1) + len(df2)}<br>
<b>Region:</b> {center_lat:.4f}°N, {center_lon:.4f}°E<br>
<i style="color: #666; font-size: 11px;">Use layer control (→) to toggle datasets</i>
</div>
'''
m.get_root().html.add_child(folium.Element(title_html))

# ==================== ADD LEGEND ====================
legend_html = '''
<div style="position: fixed; bottom: 50px; right: 50px; width: 200px;
            background-color: white; border:3px solid #333; z-index:9999; 
            font-size:12px; padding: 10px; border-radius: 4px; box-shadow: 0 0 8px rgba(0,0,0,0.2);">
<b>Legend</b><br>
<hr style="margin: 5px 0;">
<i style="background: #0066CC; border-radius: 50%; display: inline-block; 
    height: 12px; width: 12px; margin-right: 8px;"></i> Dataset 1<br>
<i style="background: #FF0000; border-radius: 50%; display: inline-block; 
    height: 12px; width: 12px; margin-right: 8px;"></i> Dataset 2<br>
<hr style="margin: 5px 0;">
<i style="font-size: 11px; color: #666;">Click markers for details</i>
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

df1['dataset'] = LAYER_NAME_1
df2['dataset'] = LAYER_NAME_2

df_export = pd.concat([
    df1[['X', 'Y', 'latitude', 'longitude', 'dataset']],
    df2[['X', 'Y', 'latitude', 'longitude', 'dataset']]
], ignore_index=True)

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
print(f"\nDatasets:")
print(f"  • {LAYER_NAME_1}: {len(df1)} points ({COLOR_1})")
print(f"  • {LAYER_NAME_2}: {len(df2)} points ({COLOR_2})")
print(f"  • Total: {len(df1) + len(df2)} points")
print(f"\nOpen {OUTPUT_MAP} in your browser")
print("="*70 + "\n")