import pandas as pd
import numpy as np
import folium
import pyproj
from pathlib import Path
from datetime import datetime

print("\n" + "="*70)
print("📊 EXCEL Gauß-Krüger to Map Visualization")
print("="*70)

# CONFIG - UPDATE THESE
EXCEL_FILE = r"Z:\08_KI-explorer\2022_T_DATA_LIAG\all_temp_Points_vis\25-03-19_Dateiverzeichnis_Bohrungen.xlsx"  # Change to your file path
SHEET_NAME = "Geo_Koordinaten"  # Change if needed
X_COLUMN = "C"  # Column with X (Gauß-Krüger)
Y_COLUMN = "D"  # Column with Y (Gauß-Krüger)
OUTPUT_MAP = "coordinates_map.html"

print(f"\n📁 File: {EXCEL_FILE}")
print(f"📋 Sheet: {SHEET_NAME}")
print(f"📍 Coordinate columns: {X_COLUMN}, {Y_COLUMN}")

# Step 1: Read Excel file
print("\nStep 1: Reading Excel file...")
try:
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    print(f"  ✓ Loaded {len(df)} rows")
    print(f"  Columns: {df.columns.tolist()}")
except Exception as e:
    print(f"  ✗ Error reading file: {e}")
    print(f"\n  Make sure:")
    print(f"    1. File path is correct")
    print(f"    2. Sheet name is correct ('{SHEET_NAME}')")
    print(f"    3. Columns C and D contain coordinates")
    exit(1)

# Step 2: Extract X, Y coordinates
print("\nStep 2: Extracting coordinates...")
try:
    # Use column letters or names
    if X_COLUMN in df.columns:
        x_col = X_COLUMN
    else:
        x_col = df.columns[int(ord(X_COLUMN) - ord('A'))]
    
    if Y_COLUMN in df.columns:
        y_col = Y_COLUMN
    else:
        y_col = df.columns[int(ord(Y_COLUMN) - ord('A'))]
    
    print(f"  X column: {x_col}")
    print(f"  Y column: {y_col}")
    
    # Get coordinates
    df['X'] = pd.to_numeric(df[x_col], errors='coerce')
    df['Y'] = pd.to_numeric(df[y_col], errors='coerce')
    
    # Remove rows with missing coordinates
    df_clean = df.dropna(subset=['X', 'Y'])
    
    print(f"  ✓ Found {len(df_clean)} valid coordinates")
    print(f"  X range: {df_clean['X'].min():.0f} to {df_clean['X'].max():.0f}")
    print(f"  Y range: {df_clean['Y'].min():.0f} to {df_clean['Y'].max():.0f}")
    
except Exception as e:
    print(f"  ✗ Error extracting coordinates: {e}")
    exit(1)

# Step 3: Convert Gauß-Krüger to WGS84
print("\nStep 3: Converting coordinates...")
print("  Detecting Gauß-Krüger zone...")

# Determine which zone based on X coordinate
x_mean = df_clean['X'].mean()
if 2500000 < x_mean < 3500000:
    zone = 3
    epsg_code = 31467  # GK zone 3 (9° central meridian)
    print(f"  → Zone 3 (9° meridian) - EPSG:31467")
elif 3500000 < x_mean < 4500000:
    zone = 4
    epsg_code = 31468  # GK zone 4 (12° central meridian)
    print(f"  → Zone 4 (12° meridian) - EPSG:31468")
elif 1500000 < x_mean < 2500000:
    zone = 2
    epsg_code = 31466  # GK zone 2 (6° central meridian)
    print(f"  → Zone 2 (6° meridian) - EPSG:31466")
else:
    print(f"  ⚠️  Could not determine zone (X mean: {x_mean:.0f})")
    print(f"  Using zone 3 (9°) as default")
    epsg_code = 31467

# Create coordinate reference systems
gk_crs = pyproj.CRS.from_epsg(epsg_code)
wgs84_crs = pyproj.CRS.from_epsg(4326)
transformer = pyproj.Transformer.from_crs(gk_crs, wgs84_crs, always_xy=True)

print(f"  Transforming {len(df_clean)} coordinates...")

# Transform
lons, lats = transformer.transform(df_clean['X'].values, df_clean['Y'].values)

df_clean['latitude'] = lats
df_clean['longitude'] = lons

print(f"  ✓ Conversion complete")
print(f"  Latitude range: {lats.min():.4f} to {lats.max():.4f}")
print(f"  Longitude range: {lons.min():.4f} to {lons.max():.4f}")

# Step 4: Create map
print("\nStep 4: Creating map...")

center_lat = lats.mean()
center_lon = lons.mean()

print(f"  Map center: {center_lat:.4f}°N, {center_lon:.4f}°E")

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

# Step 5: Add markers
print("\nStep 5: Adding markers...")

fg = folium.FeatureGroup(name=f"Coordinates ({len(df_clean)} points)", show=True)

for idx, row in df_clean.iterrows():
    # Create popup with information
    popup_text = f"""
    <b>Coordinate Point</b><br>
    <table style="border-collapse: collapse; font-size: 12px;">
    <tr><td><b>Lat:</b></td><td>{row['latitude']:.6f}°</td></tr>
    <tr><td><b>Lon:</b></td><td>{row['longitude']:.6f}°</td></tr>
    <tr><td><b>GK X:</b></td><td>{row['X']:.0f}</td></tr>
    <tr><td><b>GK Y:</b></td><td>{row['Y']:.0f}</td></tr>
    """
    
    # Add other columns if they exist
    for col in df_clean.columns:
        if col not in ['X', 'Y', 'latitude', 'longitude', x_col, y_col]:
            try:
                popup_text += f"<tr><td><b>{col}:</b></td><td>{row[col]}</td></tr>"
            except:
                pass
    
    popup_text += "</table>"
    
    folium.Marker(
        location=[row['latitude'], row['longitude']],
        popup=folium.Popup(popup_text, max_width=300),
        icon=folium.Icon(color='blue', icon='info-sign'),
        tooltip=f"({row['latitude']:.4f}, {row['longitude']:.4f})"
    ).add_to(fg)

fg.add_to(m)

# Add layer control
folium.LayerControl(position='topright').add_to(m)

# Add title
title_html = f'''
<div style="position: fixed; top: 10px; left: 50px; width: 380px; 
            background-color: white; border:2px solid #333; z-index:9999; 
            font-size:12px; padding: 10px; border-radius: 4px;">
<b style="font-size: 13px;">📍 Gauß-Krüger Coordinates Map</b><br>
<b>File:</b> {Path(EXCEL_FILE).name}<br>
<b>Points:</b> {len(df_clean)}<br>
<b>Source CRS:</b> Gauß-Krüger Zone {zone} (EPSG:{epsg_code})<br>
<b>Target CRS:</b> WGS84 (EPSG:4326)<br>
<b>Region:</b> {center_lat:.4f}°N, {center_lon:.4f}°E
</div>
'''
m.get_root().html.add_child(folium.Element(title_html))

# Step 6: Save map
print(f"\nStep 6: Saving map...")

try:
    m.save(OUTPUT_MAP)
    file_size = Path(OUTPUT_MAP).stat().st_size / (1024*1024)
    print(f"  ✓ Saved: {OUTPUT_MAP}")
    print(f"  File size: {file_size:.2f} MB")
except Exception as e:
    print(f"  ✗ Error saving: {e}")
    exit(1)

# Step 7: Save converted coordinates to CSV
print(f"\nStep 7: Saving converted coordinates...")

output_csv = OUTPUT_MAP.replace('.html', '_coordinates.csv')
df_export = df_clean[['X', 'Y', 'latitude', 'longitude']].copy()
df_export.to_csv(output_csv, index=False)
print(f"  ✓ Saved: {output_csv}")

# Done
print("\n" + "="*70)
print("✓ SUCCESS!")
print("="*70)
print(f"\nOutputs created:")
print(f"  1. {OUTPUT_MAP} - Interactive map")
print(f"  2. {output_csv} - Converted coordinates (CSV)")
print(f"\nOpen {OUTPUT_MAP} in your browser to view the map")
print("="*70 + "\n")