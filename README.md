# 🗺️ German Boreholes Visualization

Interactive mapping tool for visualizing German borehole coordinates on an interactive map with multiple dataset overlay support.

## ✨ Features

- **Multi-dataset overlay** - Display up to 4 datasets with different colors
- **Automatic coordinate conversion** - Gauß-Krüger to WGS84 (auto-detects zone)
- **Bohr ID hover display** - See borehole IDs on mouse hover
- **Interactive popups** - Click markers for detailed information
- **Layer control** - Toggle datasets on/off independently
- **Offline capable** - HTML map works without internet
- **CSV export** - Combined coordinates with dataset labels

## 🚀 Quick Start

### Installation

```bash
pip install pandas numpy folium pyproj openpyxl
```

### Usage

1. **Update your configuration** in `german_boreholes_map.py`:

```python
DATASETS = [
    {
        "file": r"C:\path\to\your\file1.xlsx",
        "sheet": "YourSheetName",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "Dataset Name",
        "color": "blue"
    },
    # Add more datasets...
]
```

2. **Run the script**:

```bash
python german_boreholes_map.py
```

3. **Open the map** - Open `German_Boreholes_Map.html` in your browser

## 📋 Excel File Format

Your Excel files should have:
- **Column B**: Bohr ID (e.g., BohID001, Well_A)
- **Column C**: X coordinate in Gauß-Krüger (e.g., 3268000)
- **Column D**: Y coordinate in Gauß-Krüger (e.g., 5226000)

## 🎨 Available Colors

- Primary: `red`, `blue`, `green`, `purple`, `orange`
- Dark: `darkred`, `darkblue`, `darkgreen`, `darkpurple`
- Light: `lightblue`, `lightgreen`, `lightgray`, `pink`

## 📤 Output Files

1. **German_Boreholes_Map.html** - Interactive map (10-50 MB)
   - Click markers for details
   - Hover to see Bohr ID
   - Toggle layers on/off
   - Switch between map views

2. **German_Boreholes_Map_combined.csv** - All coordinates combined
   - Columns: bohr_id, X, Y, latitude, longitude, dataset
   - Use for GIS or further analysis

## 🔄 Coordinate Conversion

Automatically detects and converts:
- **Zone 2** (6° meridian): X range 1.5M - 2.5M → EPSG:31466
- **Zone 3** (9° meridian): X range 2.5M - 3.5M → EPSG:31467
- **Zone 4** (12° meridian): X range 3.5M - 4.5M → EPSG:31468

All converted to WGS84 (EPSG:4326) for web mapping.

## 📝 Example

```python
DATASETS = [
    {
        "file": r"Z:\data\boreholes1.xlsx",
        "sheet": "Geo_Koordinaten",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "All Points",
        "color": "blue"
    },
    {
        "file": r"Z:\data\boreholes2.xlsx",
        "sheet": "SVZ",
        "bohr_id_col": "B",
        "x_col": "C",
        "y_col": "D",
        "name": "SVZ Data",
        "color": "red"
    },
]
```

## 🛠️ Troubleshooting

| Problem | Solution |
|---------|----------|
| "File not found" | Use raw string: `r"C:\path"` |
| "No X,Y columns" | Check columns C, D exist with data |
| Map won't open | Try Chrome or Firefox |
| Points in wrong location | Verify Gauß-Krüger coordinates (7 digits) |

## 💡 Tips

- Use raw strings for Windows paths: `r"C:\path\to\file"`
- Keep Bohr IDs unique for easy identification
- Test with small files first
- Python 3.7+ required

## 📄 License

MIT License - Free to use and modify

## 👤 Author

Created for geothermal and geophysical research

---

**Ready to visualize your boreholes? Run the script and open the HTML file in your browser!** 🗺️