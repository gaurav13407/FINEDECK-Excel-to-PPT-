# Enhanced Finance Visual Standards - Implementation Guide

## 🎯 Overview

This refactoring transforms the Excel→PowerPoint automation to produce **consulting/investor deck quality presentations** automatically. All slides and charts now follow a single, professional finance visual standard inspired by McKinsey, Goldman Sachs, and BlackRock presentations.

## 📘 Global Design Theme

### Font Standards
```python
Font Family: "Segoe UI" (primary), fallback "Lato" → "Arial"
Chart Title: 28pt Bold, #004F9E (Professional Blue)
Axis Labels: 10-11pt, #6C757D (Gray)
Data Labels: 10pt, Black
Legend: 9pt, #6C757D (Gray)
Number Color: Black
Background: White (no chart backgrounds)
```

### Color Palette
```python
Primary:   #004F9E  (Professional Blue) - Titles, main series
Secondary: #16A085  (Teal) - Positive values, growth
Accent:    #E15759  (Red) - Negative values, decline
Neutral:   #6C757D  (Gray) - Labels, legends
Light Grid: #E9ECEF (Light Gray) - Gridlines
```

### Layout Standards
```python
Top Margin: 36px (exactly)
Side Padding: 24px
Chart Width: 70-80% of slide width
Chart Max Height: 60% of slide height
Aspect Ratio: 16:9
Y-Axis Ticks: 4-6 major ticks (no more!)
X-Axis Rotation: 35° only if >8 categories
Overlap Threshold: 30%
```

## 💵 Number Formatting Rules

### Human-Readable Units
```python
≥1e12 → "$X.XT"   (e.g., "$2.5T")
≥1e9  → "$X.XXB"  (e.g., "$865.4B")
≥1e6  → "$X.XM"   (e.g., "$42.7M")
≥1e3  → "$X.XK"   (e.g., "$5.2K")
else  → "$X"      (e.g., "$1,234")
```

**NO scientific notation** (3E+12) anywhere in presentations!

### Percentages
```python
Format: X.X% (one decimal)
Example: 12.4%, 5.0%, -3.7%
```

### NaN/None Values
```python
Replace with: "–" (dash)
Never show: NaN, None, null, NA
```

## 🧩 Chart Formatting Rules

### 1️⃣ Column / Bar Charts

```python
# Limit to Top 5 + "Other"
if categories > 8:
    show_top_5_plus_other()

# Sort bars descending by value
df.sort_values(ascending=False)

# Bar gap width
bar_gap_width = 150%

# Gradient
gradient_angle = 90°  # Vertical

# Data labels (only top 3)
show_data_labels(limit=3)
format = "$X.XT"  # No full numbers
```

**Example Output:**
```
[████████████] Apple   $2.8T   ← Data label
[██████████]  Microsoft $2.6T
[████████]    Google    $1.7T
[██████]      Amazon    (no label)
[████]        NVIDIA    (no label)
[██]          Other     (no label)
```

### 2️⃣ Line Charts (Trend / Time Series)

```python
# Smooth lines
series.smooth = True

# Reduced marker size
series.marker.size = 5

# End-of-line label
show_last_value_label = True

# Legend
legend.position = BOTTOM
legend.font.size = 9pt

# Axis scale
consistent_scale_across_similar_charts = True

# Shaded band (±2% if variation detected)
add_confidence_band(percentage=2)
```

**Annotations:**
```
Peak: $2.8T on 2024-06
Last: $2.5T (+12.4% vs previous)
```

### 3️⃣ Dual Axis Charts

```python
# Create secondary Y-axis if scales differ >10×
if max_value_1 / max_value_2 > 10:
    create_secondary_axis()
    label_both_sides_clearly()
```

### 4️⃣ Donut / Pie Charts (Sector Distribution)

```python
# Use donut style (not pie)
chart_type = DONUT

# Show top 5 + "Other"
if sectors > 8:
    show_top_5_plus_other()

# Percentages inside slices
data_labels.show_percentage = True
data_labels.show_value = False

# Legend on right
legend.position = RIGHT
```

### 5️⃣ Waterfall Charts (Financial Flow)

```python
# Highlight start/end totals
start_bar.color = ACCENT
end_bar.color = ACCENT

# Data labels on ALL bars
show_all_data_labels = True
format = "$X.XB"
```

## 🧠 Insight Annotations (Automatic)

### Time Series Annotations

```python
def annotate_time_series(chart, data):
    """
    Automatically adds:
    - Peak value and date
    - Last value with % change vs previous
    """
    peak = data.max()
    peak_date = data.idxmax()
    last = data.iloc[-1]
    prev = data.iloc[-2]
    change_pct = ((last - prev) / prev) * 100
    
    # Add gray textbox
    annotation = f"Peak: {format_number(peak)} on {peak_date}\n"
    annotation += f"Last: {format_number(last)} ({change_pct:+.1f}% vs previous)"
```

**Example:**
```
╔════════════════════════╗
║ Peak: $2.8T on 2024-06 ║  (gray, 9pt)
║ Last: $2.5T (+12.4%    ║  (green if positive)
║         vs previous)   ║
╚════════════════════════╝
```

### Comparison Chart Captions

```python
def add_comparison_caption(chart, data):
    """
    Adds caption below chart:
    'Sorted descending | Top 5 metrics shown | Values in B/T'
    """
    caption_parts = []
    caption_parts.append("Sorted descending")
    
    if len(data) <= 6:  # Top 5 + Other
        caption_parts.append("Top 5 metrics shown")
    
    # Detect unit
    max_val = data.max()
    if max_val >= 1e12:
        unit = "Values in T"
    elif max_val >= 1e9:
        unit = "Values in B"
    # ...
    
    caption_parts.append(unit)
    caption = " | ".join(caption_parts)
```

**Example:**
```
────────────────────────────────────────────
Sorted descending | Top 5 metrics shown | Values in T
────────────────────────────────────────────
```

## ⚙ Cleanup & Safety

### Data Cleaning
```python
# Replace NaN/None
df = df.fillna("–")

# Remove empty rows/columns
df = df.dropna(how='all', axis=0)
df = df.dropna(how='all', axis=1)

# Fill remaining NaN with 0 for charts
df_chart = df.fillna(0)
```

### Axis Configuration
```python
# Y-Axis
major_ticks = 4-6  # No more than 6!
minor_ticks = 0    # None
gridlines = HORIZONTAL_ONLY
gridline_width = 0.25pt
gridline_color = #E9ECEF

# X-Axis
gridlines = NONE
rotation = 35° if categories > 8 or overlap > 30%
```

### Category Filtering
```python
# Auto Top-N enforcement
if len(categories) > 8:
    df_top = df.nlargest(5, value_column)
    other_sum = df.iloc[5:][value_column].sum()
    df_chart = pd.concat([
        df_top,
        pd.DataFrame({'Category': ['Other'], 'Value': [other_sum]})
    ])
    # Sort descending
    df_chart = df_chart.sort_values(by='Value', ascending=False)
```

### Layout Validation
```python
# Auto-adjust chart dimensions
chart_width = slide_width * 0.75  # 70-80%
chart_height = min(chart_height, slide_height * 0.60)  # Max 60%

# Top margin
chart_top = 36px  # From slide top

# Side padding
chart_left = 24px
chart_right = slide_width - 24px
```

## 🧱 Code Structure

### Reusable Helper Functions

```python
def format_number(value, prefix="$", decimals=1):
    """
    Format numbers with K/M/B/T suffixes
    Always prefix currency with "$"
    NO scientific notation
    """
    if value >= 1e12:
        return f"{prefix}{value/1e12:.{decimals}f}T"
    elif value >= 1e9:
        return f"{prefix}{value/1e9:.{decimals}f}B"
    # ...
    return f"{prefix}{value:,.0f}"


def clean_nan_values(series):
    """Replace NaN or None with '–' (dash)"""
    return series.fillna("–")


def apply_finance_theme(chart, chart_type=None):
    """
    Apply complete finance theme to chart
    
    - Segoe UI fonts (28pt title, 11pt axes, 9pt legend)
    - Finance color palette
    - Bottom legend (horizontal, 9pt gray)
    - 4-6 Y-ticks, no minor gridlines
    - Light horizontal gridlines only (0.25pt)
    - Chart-specific styling (smooth lines, bar gaps, etc.)
    """
    # Remove backgrounds (white)
    # Configure legend (bottom, 9pt, gray)
    # Configure axes (4-6 ticks, light gridlines)
    # Apply color palette
    # Chart-specific enhancements


def annotate_chart(chart, data, slide, position, size, chart_type):
    """
    Add automatic insight annotations
    
    LINE/AREA: Peak value + date, Last value + % change
    BAR/COLUMN: Caption with sorting info and units
    """
    if chart_type in ['LINE', 'AREA']:
        # Add peak/last annotations
        add_peak_annotation()
        add_last_value_with_change()
    
    elif chart_type in ['BAR', 'COLUMN']:
        # Add caption below chart
        add_caption_with_units()


def finalize_presentation(prs):
    """
    MASTER FINALIZER - Apply all standards globally
    
    Called automatically before saving
    Idempotent (safe to run multiple times)
    
    Applies:
    - Human-readable number formatting ($K/M/B/T)
    - Font unification (Segoe UI 28pt/11pt/9pt)
    - Gridline normalization (horizontal only, light)
    - Padding normalization (36px top)
    - Color palette consistency
    - Brand accent line under titles (2px, #004F9E)
    - Chart height validation (≤60%)
    """
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_chart:
                apply_finance_theme(shape.chart)
                normalize_chart_layout(shape.chart)
                validate_chart_dimensions(shape)
            
            if shape.has_text_frame:
                unify_fonts(shape.text_frame)
                apply_title_formatting(shape.text_frame)
```

## 🚀 Usage Examples

### Example 1: Basic Chart Creation

```python
from src.converter.advanced_finance_charts import AdvancedFinanceChartBuilder
from src.converter.finance_chart_formatter import finalize_presentation

# Create data
df = pd.DataFrame({
    'Company': ['Apple', 'Microsoft', 'Google'],
    'Market Cap': [2.8e12, 2.6e12, 1.7e12]
})

# Create chart
chart_builder = AdvancedFinanceChartBuilder(
    slide=slide,
    position=(1, 2),
    size=(8, 4.5)
)

chart_builder.create_chart(
    df=df,
    chart_type='COLUMN',
    title='Top Tech Companies'
)

# Finalize (applies all standards)
finalize_presentation(prs)

# Save
prs.save('output.pptx')
```

**Result:**
- Chart with Top 3 companies
- Market Cap shown as $2.8T, $2.6T, $1.7T
- Segoe UI 28pt Bold title with underline
- 11pt gray axis labels
- Bottom legend (9pt gray)
- 4-6 Y-ticks with light gridlines
- Caption: "Sorted descending | Values in T"

### Example 2: Time Series with Annotations

```python
df = pd.DataFrame({
    'Date': ['2024-01', '2024-02', '2024-03', '2024-04'],
    'Revenue': [2.1e9, 2.3e9, 2.8e9, 2.5e9]
})

chart_builder.create_chart(
    df=df,
    chart_type='LINE',
    title='Monthly Revenue Trend'
)

# Annotations automatically added:
# "Peak: $2.8B on 2024-03"
# "Last: $2.5B (-10.7% vs previous)"
```

### Example 3: Auto Top-N Filtering

```python
df = pd.DataFrame({
    'Product': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],
    'Sales': [100, 90, 80, 70, 60, 50, 40, 30, 20, 10]
})

chart_builder.create_chart(
    df=df,
    chart_type='BAR',
    title='Product Performance'
)

# Automatically filters to Top 5 + "Other"
# Result: 6 bars (A, B, C, D, E, Other)
# Sorted descending
# Caption: "Sorted descending | Top 5 metrics shown | Values in units"
```

## 📊 Before vs After

### Before Enhancements
```
❌ Scientific notation: 3E+12
❌ Full numbers: 4,000,000,000
❌ Inconsistent fonts: Arial 32pt, Calibri 10pt
❌ Random colors
❌ 15+ Y-axis ticks
❌ Vertical and horizontal gridlines
❌ Right-side legend
❌ All 13 categories shown
❌ Overlapping labels
❌ No annotations
```

### After Enhancements
```
✅ Human-readable: $3.0T
✅ Formatted numbers: $4.0B
✅ Consistent fonts: Segoe UI 28pt Bold (titles), 11pt (axes), 9pt (legends)
✅ Finance palette: #004F9E, #16A085, #E15759
✅ 4-6 Y-axis ticks (exact)
✅ Horizontal gridlines only (0.25pt, light gray)
✅ Bottom-center legend (9pt gray)
✅ Top 5 + "Other" (auto-filtered)
✅ 35° rotated labels (no overlap)
✅ Auto annotations: "Peak: $X on DATE", "Last: $Y (+Z%)"
✅ Brand accent underline on titles (2px, #004F9E)
✅ Chart height ≤60% of slide
```

## 🧪 Testing

Run comprehensive test suite:
```bash
python test_enhanced_finance_standards.py
```

**Tests:**
1. Number formatting ($ prefix, K/M/B/T suffixes)
2. Percentage formatting (1 decimal)
3. NaN cleaning (replace with "–")
4. Global design constants validation
5. Full chart creation integration
6. finalize_presentation() master normalizer
7. Validation checks (fonts, colors, layout)

**Expected Output:**
```
✅ 2.53e+12 → $2.5T
✅ 865.4e+09 → $865.4B
✅ Title font: Segoe UI 28pt Bold
✅ Axis font: Segoe UI 11pt Gray
✅ Legend: 9pt Gray, bottom-center
✅ Y-ticks: 4-6 (not >6)
✅ Top-N: 5 + Other
✅ Chart width: 70-80% of slide
✅ Chart height: ≤60% of slide
```

## 🎉 Production Ready

All features are now production-ready and automatically applied to every Excel upload:

1. **Automatic:** No manual formatting needed
2. **Consistent:** Single finance visual standard across all slides
3. **Professional:** Consulting/investor deck quality
4. **Idempotent:** Safe to run finalize_presentation() multiple times
5. **Client-Ready:** No post-processing required

**Every Excel upload now produces McKinsey/Goldman-quality presentations!** 🚀
