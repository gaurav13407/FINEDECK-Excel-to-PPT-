# Sales Performance Dashboard - Quick Start

## 📦 Package Contents

- `data/` - CSV files ready for Power BI import
- `dashboard_model.json` - Complete model specification
- `DAX_Measures.txt` - Copy-paste ready DAX measures
- `QUICK_START.md` - This guide

## 🚀 Setup Instructions (5 minutes)

### Step 1: Open Power BI Desktop

1. Download Power BI Desktop (if not installed): https://powerbi.microsoft.com/desktop/
2. Launch Power BI Desktop
3. Click "Get Data" → "Text/CSV"

### Step 2: Import Data

1. Navigate to the `data/` folder in this package
2. Select all CSV files and click "Open"
3. Click "Load" to import the data
4. Wait for data to load (should take a few seconds)

### Step 3: Add DAX Measures

1. Go to "Modeling" tab in the ribbon
2. Click "New Measure"
3. Open `DAX_Measures.txt` from this package
4. Copy each measure one by one and paste into Power BI
5. Press Enter after each measure

### Step 4: Create Visuals

1. **Sales Trend**
   - Type: Line Chart
   - Fields: Date (X), Salesamount (Y), Region (Legend)

2. **Regional Map**
   - Type: Map
   - Fields: Region (Location), Salesamount (Size)

3. **Top Products**
   - Type: Bar Chart
   - Fields: Product (Y), Salesamount (X)

4. **Total Sales KPI**
   - Type: Card
   - Fields: TotalSalesamount

5. **Target Achievement**
   - Type: Gauge
   - Fields: TotalSalesamount (Value), TotalSalestarget (Target)


### Step 5: Apply Theme & Publish

1. Go to "View" → "Themes" → Choose a theme
2. Add slicers for filtering (Date, Region, Category, etc.)
3. Save your dashboard: File → Save
4. Publish to Power BI Service: Home → Publish

## 💡 Pro Tips

- Use Ctrl+Click to select multiple visuals at once
- Right-click visuals for formatting options
- Enable "Show Data Label" for better readability
- Use bookmarks to save different dashboard views
- Set up automatic data refresh in Power BI Service

## 🆘 Need Help?

- Power BI Documentation: https://docs.microsoft.com/power-bi/
- DAX Reference: https://dax.guide/
- Community Forum: https://community.powerbi.com/

---

**Created by FinDeck Power BI Dashboard Generator**  
**Date: C:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)**
