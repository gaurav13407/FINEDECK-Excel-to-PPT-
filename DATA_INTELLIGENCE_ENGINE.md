# FinDeck Data Intelligence Engine

## 🎯 Overview

The **Data Intelligence Engine** is FinDeck's core analysis component that automatically analyzes **any** Excel file and produces structured, accurate insights for PowerPoint generation.

**Key Features:**
- ✅ **100% Data-Driven** - Zero hallucination, only real data
- ✅ **Universal Compatibility** - Works with ANY dataset structure
- ✅ **Automatic Classification** - Detects column types automatically
- ✅ **Hierarchy Detection** - Identifies Sales, Profit, Geography patterns
- ✅ **PPT-Ready Output** - Structured JSON for slide generation

---

## 🚀 How It Works

### STEP 1: Automatic Column Classification

The engine analyzes each column and classifies it as:

| Type | Description | Examples |
|------|-------------|----------|
| **Identifier** | Never summed - unique IDs | Order ID, Customer ID, Postal Code |
| **Numeric** | Quantitative measurements | Sales, Quantity, Price, Revenue |
| **Categorical** | Limited unique values | Category, Region, Product Type |
| **Date** | Temporal data | Order Date, Ship Date, Timestamp |
| **Boolean** | True/false values | Is Active, Has Discount |
| **Text** | Free-form text | Description, Notes, Comments |

**Identifier Detection Rules:**
```python
# These patterns are NEVER summed:
- *_id, id_*
- order_id, customer_id, product_id
- row_id, transaction_id, invoice_id
- postal_code, zip_code
- phone, email, ssn, account_number
```

---

### STEP 2: Metric Extraction

#### For Numeric Columns:
```json
{
  "sum": 1500000.00,
  "mean": 15000.00,
  "median": 14500.00,
  "min": 100.00,
  "max": 50000.00,
  "std": 5000.00,
  "top_10_values": [50000, 48000, 45000, ...]
}
```

#### For Categorical Columns:
```json
{
  "unique_count": 15,
  "total_count": 1000,
  "top_10_categories": [
    {"category": "Electronics", "count": 450, "percentage": 45.0},
    {"category": "Furniture", "count": 350, "percentage": 35.0},
    ...
  ]
}
```

#### For Date Columns:
```json
{
  "earliest": "2024-01-01",
  "latest": "2024-12-31",
  "monthly_distribution": [
    {"month": "2024-01", "count": 85},
    {"month": "2024-02", "count": 92},
    ...
  ],
  "trend": "increasing"
}
```

---

### STEP 3: Hierarchy Analysis

The engine automatically detects and analyzes:

#### 📊 Sales/Revenue
```python
# Detected patterns:
sales, revenue, amount, total_price, value, gmv, turnover

# Analysis includes:
- Total Sales
- Average Sale
- Sales by Category (if categorical column exists)
- Sales by Region (if geo column exists)
- Monthly Sales Trend (if date column exists)
```

#### 💰 Profit
```python
# Detected patterns:
profit, margin, net_income, earnings, ebitda

# Analysis includes:
- Total Profit
- Average Profit
- Profit Margin % (if sales column exists)
- Profit by Category
- Profit Trend
```

#### 🌍 Geography
```python
# Detected patterns:
city, state, country, region, postal_code, location

# Analysis includes:
- Top locations by count
- Top locations by sales
- Geographic distribution
```

---

### STEP 4: PPT-Ready Output

The engine returns a structured JSON perfect for PowerPoint generation:

```json
{
  "executive_summary": {
    "total_records": 1000,
    "total_columns": 12,
    "numeric_columns": 5,
    "categorical_columns": 4,
    "date_columns": 2,
    "identifier_columns": 1
  },
  "key_metrics": [
    {
      "metric_name": "Sales",
      "total": 1500000.00,
      "average": 1500.00,
      "min": 100.00,
      "max": 50000.00
    }
  ],
  "top_categories": [
    {
      "category_type": "Product Category",
      "unique_count": 15,
      "top_values": [...]
    }
  ],
  "hierarchy_analysis": {
    "sales": {...},
    "profit": {...},
    "geography": {...}
  },
  "trend_analysis": {...},
  "recommendations": [
    "Total sales of 1,500,000.00 with average transaction of 1,500.00",
    "Profit margin is 25.5%",
    "Positive Order Date trend detected"
  ]
}
```

---

## 📋 Strict Rules (Zero Hallucination)

### ❌ NEVER:
- Sum identifier columns (Order ID, Customer ID, Postal Code)
- Generate placeholder values
- Invent numbers not in the data
- Repeat the same metric unless it's actually in the data

### ✅ ALWAYS:
- Check if column is numeric before summing
- Generate trends based on actual date columns
- Ensure distribution percentages sum to 100%
- Base all values on real data only

---

## 🔧 Usage Examples

### Example 1: Sales Dataset
```python
from services.data_intelligence import DataIntelligenceEngine

engine = DataIntelligenceEngine()
results = engine.analyze_file('sales_data.xlsx')

# Results include:
# - Total Sales, Average Sale, Sales by Category
# - Total Profit, Profit Margin
# - Monthly Sales Trend
# - Top Locations by Sales
```

### Example 2: Marketing Dataset
```python
results = engine.analyze_file('marketing_data.xlsx')

# Results include:
# - Impressions, Clicks, Conversions metrics
# - Campaign performance by channel
# - ROI calculations (if revenue exists)
# - Trend analysis
```

### Example 3: Unknown Dataset
```python
results = engine.analyze_file('mystery_data.xlsx')

# Engine automatically:
# - Classifies all columns
# - Extracts appropriate metrics
# - Detects any hierarchy patterns
# - Returns structured insights
```

---

## 🌐 API Integration

### Endpoint 1: Analyze Excel File
```http
POST /api/v1/analyze-excel
Content-Type: multipart/form-data

file: [Excel file]
```

**Response:**
```json
{
  "success": true,
  "data": {
    "executive_summary": {...},
    "key_metrics": [...],
    "hierarchy_analysis": {...},
    "recommendations": [...]
  }
}
```

### Endpoint 2: Intelligent Conversion
```http
POST /api/v1/convert-with-intelligence
Content-Type: multipart/form-data

file: [Excel file]
template_id: [optional]
```

**Response:**
```json
{
  "success": true,
  "analysis": {...},
  "ppt_slides_generated": 15,
  "message": "Intelligent conversion completed"
}
```

---

## 🧪 Testing

Run the test suite:

```bash
cd src/backend/app/services
python test_data_intelligence.py
```

This will:
1. Create sample datasets (Sales, Marketing, Unknown)
2. Analyze each one
3. Generate JSON output files
4. Demonstrate 100% accuracy with zero hallucination

---

## 📊 Supported Dataset Types

The engine works with **any** dataset, including:

- ✅ **Sales Data** - Orders, transactions, revenue
- ✅ **Marketing Data** - Campaigns, impressions, conversions
- ✅ **Finance Data** - P&L, balance sheets, expenses
- ✅ **HR Data** - Employees, salaries, performance
- ✅ **Health Data** - Patients, treatments, outcomes
- ✅ **E-commerce Data** - Products, customers, orders
- ✅ **Supply Chain Data** - Inventory, shipments, logistics
- ✅ **Any CSV/Excel** - Unknown schema, automatically analyzed

---

## 🎯 Key Advantages

1. **Zero Configuration** - No need to tell the engine what your data is
2. **Universal** - Works with any industry, any dataset
3. **Accurate** - 100% based on actual data, no guessing
4. **Fast** - Analyzes thousands of rows in seconds
5. **PPT-Ready** - Output is structured for slide generation
6. **Business-Friendly** - Generates insights executives understand

---

## 🔮 Future Enhancements

- [ ] Advanced trend detection (seasonality, anomalies)
- [ ] Correlation analysis between columns
- [ ] Predictive insights (forecasting)
- [ ] Custom metric definitions
- [ ] Multi-sheet Excel support
- [ ] Real-time streaming analysis
- [ ] Power BI integration for dashboards

---

## 📝 License

Proprietary - FinDeck Excel to PPT Conversion Tool

---

## 💡 Support

For questions or issues:
- Email: support@findeck.live
- Documentation: Coming soon
- API Docs: `/api/docs` (Swagger UI)

---

**Built with intelligence. Powered by data. Zero hallucination.** 🎯
