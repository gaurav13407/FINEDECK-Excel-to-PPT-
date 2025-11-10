# 📊 Advanced Finance Chart System

## Overview

FinDeck now features a **completely new advanced chart system** with **12+ professional financial chart types** and intelligent auto-detection. The system automatically selects the best chart type based on your data structure and column names.

---

## 🎨 Supported Chart Types

### 1️⃣ **Performance & Growth Tracking**

#### LINE Chart
- **Use Case**: Stock prices, sales growth, profit margins over time
- **Why**: Shows clear trends and changes
- **Auto-Detection**: Looks for date/time columns (date, quarter, month, year, period)
- **Example Data**: Quarterly revenue, monthly sales, daily stock prices

#### AREA Chart  
- **Use Case**: Net income, cumulative profit, expenses
- **Why**: Visually stronger for growth over time, emphasizes magnitude
- **Auto-Detection**: Keywords like "cumulative", "total", "net income", "profit margin"
- **Example Data**: Cumulative cash flow, net profit trends

#### COLUMN Chart
- **Use Case**: Monthly revenue, YoY comparisons, performance rankings
- **Why**: Ideal for side-by-side comparison
- **Auto-Detection**: Comparison keywords ("vs", "yoy", "mom", "performance", "ranking")
- **Example Data**: Regional sales comparison, department performance

#### BAR Chart
- **Use Case**: Many categories or long category labels
- **Why**: Better readability for long labels and many items
- **Auto-Detection**: Many rows (>10) or comparison keywords
- **Example Data**: Top 20 products by revenue, department rankings

---

### 2️⃣ **Portfolio & Asset Analysis**

#### PIE Chart
- **Use Case**: Asset allocation, revenue by sector (simple distributions)
- **Why**: Good for showing parts of a whole
- **Auto-Detection**: Keywords like "allocation", "portfolio", "sector", "distribution"
- **Example Data**: Asset class breakdown, market share

#### DONUT Chart (Recommended over Pie)
- **Use Case**: Portfolio distribution, sector allocation
- **Why**: More professional appearance, allows center label
- **Auto-Detection**: Same as pie, preferred for professional presentations
- **Example Data**: Investment portfolio allocation, expense categories

#### STACKED COLUMN Chart
- **Use Case**: Portfolio composition over time, segment performance
- **Why**: Tracks changes in allocations over time periods
- **Auto-Detection**: Multiple value columns + time dimension
- **Example Data**: Quarterly portfolio composition, monthly expense breakdown

---

### 3️⃣ **Profit, Expense & Cash Flow**

#### WATERFALL Chart
- **Use Case**: P&L flow (Revenue → Net Income), cash flow bridges
- **Why**: Perfect for showing incremental changes and flow
- **Auto-Detection**: P&L keywords ("revenue", "cogs", "expense", "profit", "ebitda")
- **Example Data**: Income statement breakdown, cash flow analysis

#### STACKED BAR Chart
- **Use Case**: Expense composition per quarter, cost breakdown
- **Why**: Shows components of totals across categories
- **Auto-Detection**: Multiple expense/cost columns
- **Example Data**: Quarterly expense composition, departmental costs

---

### 4️⃣ **Financial Market & Stock Analysis**

#### CANDLESTICK Chart
- **Use Case**: Stock price movement, trading analysis
- **Why**: Shows Open, High, Low, Close (OHLC) clearly
- **Auto-Detection**: Requires columns: Open, High, Low, Close
- **Example Data**: Daily stock prices, forex rates

#### SCATTER Plot
- **Use Case**: Risk-return relationship, correlation analysis
- **Why**: Perfect for showing relationships between two variables
- **Auto-Detection**: Keywords like "risk", "return", "correlation", "vs"
- **Example Data**: Portfolio risk vs return, price vs volume analysis

#### BUBBLE Chart
- **Use Case**: 3D analysis (market cap, risk, return)
- **Why**: Visualizes three dimensions simultaneously
- **Auto-Detection**: 3+ numeric columns with keywords like "market cap", "volume", "size"
- **Example Data**: Stock comparison (cap/return/volatility)

---

## 🧠 Intelligent Auto-Detection

The system uses a **priority-based detection algorithm**:

### Detection Priority Order:

1. **CANDLESTICK** - Highest priority if OHLC columns found
2. **WATERFALL** - P&L keywords detected
3. **BUBBLE** - 3+ numeric dimensions
4. **SCATTER** - Risk/return/correlation keywords
5. **DONUT/PIE** - Allocation/distribution keywords
6. **STACKED_COLUMN** - Composition + time columns
7. **AREA** - Cumulative growth keywords
8. **LINE** - Time-series data
9. **COLUMN/BAR** - Default comparison charts

### Detection Examples:

```python
# Example 1: Stock Data
df = pd.DataFrame({
    'Date': ['2025-01-01', '2025-01-02'],
    'Open': [150.2, 152.5],
    'High': [153.8, 154.2],
    'Low': [149.5, 151.2],
    'Close': [152.5, 151.8]
})
# → AUTO-DETECTS: CANDLESTICK

# Example 2: Portfolio Allocation
df = pd.DataFrame({
    'Sector': ['Technology', 'Healthcare', 'Finance'],
    'Allocation': [45, 30, 25]
})
# → AUTO-DETECTS: DONUT

# Example 3: Revenue Growth
df = pd.DataFrame({
    'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024'],
    'Revenue': [125000, 142000, 158000],
    'Profit': [35000, 42000, 48000]
})
# → AUTO-DETECTS: LINE

# Example 4: P&L Bridge
df = pd.DataFrame({
    'Category': ['Revenue', 'COGS', 'Gross Profit', 'OpEx', 'Net Income'],
    'Amount': [1000000, -450000, 550000, -280000, 270000]
})
# → AUTO-DETECTS: WATERFALL
```

---

## 🎨 Professional Styling Features

### Color Schemes
- **Profit/Positive**: Green (#228B22)
- **Loss/Negative**: Red (#DC143C)
- **Neutral**: Steel Blue (#4682B4)
- **Primary (Brand)**: Gold (#E59D02)
- **Secondary**: Gray (#646464)

### Automatic Features
- ✅ **Smart Column Selection**: Skips date columns from numeric series
- ✅ **Data Labels**: Auto-added for small datasets (≤10 items)
- ✅ **Percentage Labels**: Automatic on pie/donut charts
- ✅ **Profit/Loss Colors**: Green/red based on positive/negative values
- ✅ **Trendlines**: Added where applicable
- ✅ **Professional Fonts**: 18pt titles, 10pt legends
- ✅ **Legend Positioning**: Optimized per chart type

---

## 💻 Usage

### Basic Usage (Auto-Detection)

```python
from src.converter.advanced_finance_charts import AdvancedFinanceChartBuilder
import pandas as pd

# Create your data
df = pd.DataFrame({
    'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'Revenue': [100000, 120000, 145000, 180000],
    'Profit': [25000, 32000, 41000, 55000]
})

# Create chart (auto-detects LINE chart)
builder = AdvancedFinanceChartBuilder(slide, position=(1, 2), size=(8, 4.5))
success = builder.create_chart(df, title="Quarterly Performance")
```

### Explicit Chart Type

```python
# Force specific chart type
builder.create_chart(df, chart_type='AREA', title="Cumulative Growth")
builder.create_chart(df, chart_type='COLUMN', title="YoY Comparison")
builder.create_chart(df, chart_type='WATERFALL', title="P&L Bridge")
```

### Integration with Backend

The new system is already integrated with your backend:

```python
# In enhanced_professional_builder.py
self.chart_builder = AdvancedFinanceChartBuilder(None, self.template_colors)

# Charts are created automatically in slides:
# - Executive Summary
# - Data Insights  
# - Sector Distribution
# - Trend Analysis
```

---

## 📈 Chart Type Selection Guide

### When to Use Each Chart:

| Your Goal | Best Chart Type | Alternative |
|-----------|----------------|-------------|
| Show trend over time | LINE | AREA |
| Compare categories | COLUMN | BAR |
| Show parts of whole | DONUT | PIE |
| Track composition changes | STACKED_COLUMN | STACKED_BAR |
| Show P&L flow | WATERFALL | COLUMN |
| Display stock prices | CANDLESTICK | LINE |
| Analyze correlation | SCATTER | - |
| Compare 3 dimensions | BUBBLE | SCATTER |
| Emphasize growth | AREA | LINE |
| Many categories | BAR | COLUMN |

---

## 🚀 Performance Optimizations

- **Data Limiting**: Charts use first 20-30 rows for clarity
- **Series Limiting**: Max 5 series for line charts, 8 for stacked
- **Smart Skipping**: Automatically skips ID, date, timestamp columns from numeric series
- **Fallback System**: Graceful degradation if chart type fails

---

## 🎯 Best Practices

### ✅ DO:
- Use descriptive column names (e.g., "Quarterly Revenue" not "Col1")
- Include date columns for time-series data
- Keep data sets focused (10-30 rows for most charts)
- Use consistent units in your data
- Provide meaningful chart titles

### ❌ DON'T:
- Mix different scales in same chart (use separate charts)
- Use >10 categories in pie/donut charts
- Include too many series (>5 lines gets cluttered)
- Use scatter/bubble for categorical data
- Forget to label axes properly

---

## 🔧 Technical Details

### Dependencies
- `python-pptx` - PowerPoint generation
- `pandas` - Data manipulation
- `numpy` - Numerical operations

### File Location
- Main Chart Builder: `src/converter/advanced_finance_charts.py`
- Integration: `src/converter/enhanced_professional_builder.py`
- Test Suite: `test_advanced_charts.py`

### Output
- Professional PowerPoint charts
- Automatic positioning and sizing
- Template color integration
- Print-ready quality

---

## 📊 Example Output

Run the test suite to see all chart types:

```bash
python test_advanced_charts.py
```

This creates `examples/demo_PPT/advanced_finance_charts_showcase.pptx` with:
- ✅ 12 slides, each demonstrating a different chart type
- ✅ Sample financial data
- ✅ Professional styling
- ✅ Real-world use cases

---

## 🎉 Summary

The new Advanced Finance Chart System gives you:

1. **12+ Professional Chart Types** - Cover all financial visualization needs
2. **Intelligent Auto-Detection** - No manual configuration needed
3. **Smart Column Selection** - Automatically picks the right data
4. **Professional Styling** - Publication-ready charts
5. **Seamless Integration** - Works with existing backend
6. **Fallback Safety** - Graceful degradation if issues occur

**Result**: Beautiful, professional financial presentations in seconds! 🚀
