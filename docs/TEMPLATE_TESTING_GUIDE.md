# 🎨 Template Testing Guide

## Quick Test Steps

### 1. Start Backend
```bash
cd src\backend
python -m uvicorn app.main:app --reload
```

Watch for logs like:
```
🎨 Received template_name from frontend: royal_purple
🎨 Template name: Royal Purple
🎨 Primary color (navy): (74, 20, 140)  ← This should change!
```

### 2. Open Frontend
http://localhost:8001/mainpage.html

### 3. Test Template Changes

| Step | Action | Expected Result |
|------|--------|----------------|
| 1 | Upload MSFT_Financial_Data.xlsx | File uploaded ✅ |
| 2 | Select **Royal Purple** | Radio button checked |
| 3 | Click "Generate PowerPoint" | Conversion starts |
| 4 | Check browser console | `template_name: royal_purple` |
| 5 | Check backend terminal | `Primary color: (74, 20, 140)` |
| 6 | Download PPT | File downloaded |
| 7 | Open PPT slide 1 | **PURPLE background** (not blue!) |
| 8 | Check slide 2 headings | **PURPLE text** (not blue!) |

### 4. Compare Templates

Create 3 PPTs with different templates and compare:

**Corporate Blue** (default):
- Title background: Navy blue (25, 42, 86)
- Looks: Traditional corporate

**Royal Purple**:
- Title background: Deep purple (74, 20, 140)  
- Looks: Luxurious, creative

**Elegant Gray**:
- Title background: Dark gray (69, 90, 100)
- Looks: Professional, minimalist

### 5. Visual Checklist

Open the downloaded PPT and verify:
- [ ] Slide 1 (Title): Background color matches template
- [ ] Slide 2 (Executive Summary): Heading color matches template
- [ ] Slide 3 (Key Metrics): Card colors match template
- [ ] Slide 4+: Chart colors from template's chart_colors array

---

## If Colors Still Don't Change

### Check 1: Backend Receiving Template?
Look for this in backend terminal:
```
🎨 Received template_name from frontend: royal_purple
```
- If missing or shows `corporate_blue`, frontend not sending correctly
- If shows correct template, continue to Check 2

### Check 2: Template File Loading?
Look for this in backend terminal:
```
✅ Using template: royal_purple
```
- If missing, template file not found
- Check `src/templates/built_in/royal_purple.json` exists

### Check 3: Template Colors Loading?
Look for this in backend terminal:
```
🎨 Primary color (navy): (74, 20, 140)
```
- If shows `(25, 42, 86)`, template colors not loading
- Check `_load_template_colors()` method

### Check 4: Python Cache?
Sometimes Python caches old code:
```bash
# Delete cache
cd src\converter
del /s /q __pycache__

# Restart backend
cd ..\backend
python -m uvicorn app.main:app --reload
```

---

## Success Criteria ✅

You'll know it's working when:
1. Browser console shows: `template_name: royal_purple` or your selected template
2. Backend shows: `Primary color: (74, 20, 140)` (NOT `(25, 42, 86)`)
3. Downloaded PPT has PURPLE title slide (NOT blue)
4. All headings throughout PPT are purple
5. Charts use purple color palette

---

## Color Reference

### Corporate Blue (Default)
```python
primary: (25, 42, 86)     # Navy blue
accent: (52, 152, 219)    # Light blue
```

### Royal Purple
```python
primary: (74, 20, 140)    # Deep purple
accent: (171, 71, 188)    # Light purple
```

### Elegant Gray
```python
primary: (69, 90, 100)    # Dark gray
accent: (255, 87, 34)     # Orange
```

### Forest Green
```python
primary: (6, 95, 70)      # Dark green
accent: (102, 187, 106)   # Light green
```

If you see `(25, 42, 86)` in ANY template, something is wrong!

---

**Ready to test?** Upload a file and select "Royal Purple" - the PPT should be PURPLE! 🟣
