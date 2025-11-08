# FinDeck Template Selection - Testing Instructions

## Quick Start

### Option 1: Use the Startup Script (RECOMMENDED)
1. Double-click `START_APP.bat` in the project root
2. It will automatically:
   - Start Backend on http://localhost:8000
   - Start Frontend on http://localhost:8001
   - Open browser to http://localhost:8001/mainpage.html
3. Press any key in the batch window to stop both servers

### Option 2: Manual Start

**Terminal 1 - Backend:**
```bash
cd "c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend"
python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

**Terminal 2 - Frontend:**
```bash
cd "c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\ui"
python -m http.server 8001
```

**Browser:**
Open: http://localhost:8001/mainpage.html

---

## Testing Template Selection

### Step 1: Upload File
1. Click "Browse Files" or drag & drop an Excel file
2. File list will appear

### Step 2: Select Template
You should see 10 template checkboxes:
- ✅ Corporate Blue (selected by default)
- ⚪ Modern Gradient
- ⚪ Minimal White
- ⚪ Financial Pro
- ⚪ Executive Suite
- ⚪ Tech Blue
- ⚪ Creative Studio
- ⚪ Luxury Gold
- ⚪ Startup Pitch
- ⚪ Professional Gray

### Step 3: Check Browser Console
Press F12 to open DevTools, go to Console tab

You should see these logs when selecting a template:
```
🎨 ========= TEMPLATE SELECTION =========
🎨 Template selected: modern_gradient
🎨 Previous template was: corporate_blue
🎨 New template is now: modern_gradient
🎨 ====================================
✅ Convert button enabled with template: modern_gradient
```

### Step 4: Generate PowerPoint
Click "Generate PowerPoint" button

Check console for these logs:
```
🎨 SELECTED TEMPLATE VALUE: modern_gradient
📋 ALL PROPERTIES: {selectedTemplate: "modern_gradient", availableTemplates: Array(10)}
🎨 CONVERSION OPTIONS: {
  "fileName": "sample.xlsx",
  "template_name": "modern_gradient",  ← Should match selected!
  "presentation_title": "sample",
  "useTieredConversion": true,
  "tier": "ai_pro"
}
📤 Adding template_name to FormData: modern_gradient
📋 FormData contents:
  - file : [object File]
  - template_name : modern_gradient  ← Should be here!
  - presentation_title : sample
```

---

## Troubleshooting

### CORS Error
**Error:** `Access to fetch at 'http://localhost:8000/...' has been blocked by CORS`

**Solution:** 
- Make sure backend is running on http://localhost:8000
- Don't open file:// URLs - use http://localhost:8001 instead

### Template Not Changing
Check console logs:
1. Is `selectTemplate()` being called when you click a checkbox?
2. Does `this.selectedTemplate` show the correct value?
3. Does FormData contain the correct `template_name`?

### Backend Not Receiving Template
Check backend logs for the template_name parameter in the request

---

## What Should Happen

1. ✅ Select "Modern Gradient" checkbox
2. ✅ Console shows: `Template selected: modern_gradient`
3. ✅ Click "Generate PowerPoint"
4. ✅ Console shows template_name in FormData
5. ✅ Backend receives template_name
6. ✅ PowerPoint created with Modern Gradient template

---

## Expected Console Output Flow

```
# On page load:
📡 Loading templates...
🎨 ========= TEMPLATE SELECTION =========
🎨 Template selected: corporate_blue
🎨 Previous template was: undefined
🎨 New template is now: corporate_blue
🎨 ====================================
✅ Convert button enabled with template: corporate_blue

# When selecting different template:
🎨 ========= TEMPLATE SELECTION =========
🎨 Template selected: luxury_gold
🎨 Previous template was: corporate_blue
🎨 New template is now: luxury_gold
🎨 ====================================
✅ Convert button enabled with template: luxury_gold

# When clicking Generate PowerPoint:
🔄 Starting conversion process...
🎨 SELECTED TEMPLATE VALUE: luxury_gold
📋 ALL PROPERTIES: {selectedTemplate: "luxury_gold", ...}
🎨 CONVERSION OPTIONS: {...template_name: "luxury_gold"...}
📤 Adding template_name to FormData: luxury_gold
📋 FormData contents:
  - file : [object File]
  - template_name : luxury_gold
  - presentation_title : ...
```

If you see different template names in these logs, the selection is working!
