# 🤖 AI-Powered Insights Implementation Guide
## Using Groq API for Ultra-Fast AI Analysis

---

## 🎯 Why Groq API?

### **Advantages:**
✅ **10x Faster** than OpenAI (2,000+ tokens/sec vs 200 tokens/sec)
✅ **70% Cheaper** than GPT-4 ($0.27/M tokens vs $0.90/M tokens)
✅ **Free Tier** - 14,400 requests/day, perfect for MVP
✅ **Multiple Models** - Llama 3, Mixtral, Gemma
✅ **Low Latency** - Real-time responses (<1 second)
✅ **Simple API** - OpenAI-compatible

### **Perfect For:**
- Real-time chart analysis
- Bulk presentation processing
- Cost-effective scaling
- MVP testing without spending

---

## 📊 AI Features to Implement

### **Phase 1: Core AI Insights (Week 1-2)**

#### 1. **Smart Slide Titles** 🎯
**Input:** Chart data + chart type
**Output:** Engaging, data-driven title

```python
# Example
Input: Revenue data showing 23% increase
Output: "Revenue Surges 23% in Q4, Beating Expectations"

Input: Top 5 products by sales
Output: "Product A Dominates with 42% Market Share"
```

#### 2. **Auto-Generated Bullet Points** 📝
**Input:** DataFrame with numbers
**Output:** 3-5 key insights

```python
# Example
Input: Quarterly sales data
Output:
• Revenue grew 23% YoY, driven by strong Q4 performance
• Product line A accounts for 67% of total revenue
• Q4 sales exceeded forecast by $2.3M (18%)
• Customer acquisition cost decreased 15%
• All regions showed positive growth, APAC leading at 34%
```

#### 3. **Trend Detection** 📈
**Input:** Time series data
**Output:** Pattern analysis

```python
# Detects:
- Growth/Decline trends
- Seasonality patterns  
- Anomalies/outliers
- Inflection points
- Forecasts (simple)

# Example
"Revenue shows strong upward trend with 12% monthly growth rate.
Notable spike in December (+45%) indicates seasonal pattern.
Forecast: $1.2M next quarter assuming current growth continues."
```

#### 4. **Executive Summary** 🎯
**Input:** All slides data
**Output:** One-slide summary

```python
# Auto-generates first slide with:
- Top 3 key findings
- Main recommendation
- Critical metrics
- Risk/opportunity highlights
```

#### 5. **Smart Recommendations** 💡
**Input:** Data patterns
**Output:** Actionable advice

```python
# Examples:
"Focus marketing spend on Q4 when conversion is 2x higher"
"Product C shows declining trend - investigate or phase out"
"APAC region underperforming - consider additional resources"
```

---

## 🛠️ Technical Implementation

### **Step 1: Install Dependencies**

```bash
pip install groq
pip install pandas numpy
```

### **Step 2: Get Groq API Key**

1. Visit: https://console.groq.com
2. Sign up (free)
3. Get API key
4. Free tier: 14,400 requests/day (enough for 1,000+ presentations)

### **Step 3: Create AI Service**

```python
# src/services/ai_service.py

import os
from groq import Groq
import pandas as pd
import json
from typing import Dict, List, Optional, Tuple

class AIInsightsService:
    """
    AI-powered insights generation using Groq API
    Ultra-fast analysis for presentations
    """
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("GROQ_API_KEY")
        self.client = Groq(api_key=self.api_key)
        
        # Model selection (fastest to most capable)
        self.fast_model = "llama-3.1-8b-instant"  # 1000+ tokens/sec
        self.balanced_model = "mixtral-8x7b-32768"  # 500+ tokens/sec
        self.capable_model = "llama-3.1-70b-versatile"  # 300+ tokens/sec
    
    def generate_slide_title(self, df: pd.DataFrame, chart_type: str, 
                            column_names: List[str]) -> str:
        """Generate engaging title for a slide"""
        
        # Prepare data summary
        data_summary = self._summarize_dataframe(df, max_rows=5)
        
        prompt = f"""You are an expert at creating engaging presentation titles.

Data Type: {chart_type} chart
Columns: {', '.join(column_names)}
Data Sample:
{data_summary}

Create a clear, data-driven title (max 10 words) that highlights the key insight.
Focus on the most important number or trend.

Examples:
- "Revenue Grew 23% YoY, Reaching $12M"
- "Product A Leads with 42% Market Share"
- "Q4 Sales Exceeded Target by $2.3M"

Title:"""

        response = self.client.chat.completions.create(
            model=self.fast_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=50
        )
        
        title = response.choices[0].message.content.strip()
        return title.strip('"\'')
    
    def generate_bullet_points(self, df: pd.DataFrame, chart_type: str,
                               num_points: int = 5) -> List[str]:
        """Generate key insights as bullet points"""
        
        # Calculate statistics
        stats = self._calculate_statistics(df)
        data_summary = self._summarize_dataframe(df, max_rows=10)
        
        prompt = f"""You are a data analyst creating insights for an executive presentation.

Chart Type: {chart_type}
Data Summary:
{data_summary}

Statistics:
{json.dumps(stats, indent=2)}

Generate {num_points} concise, data-driven bullet points (each max 15 words).
Focus on trends, comparisons, and actionable insights.
Use specific numbers and percentages.

Format: Start each point with •

Example:
• Revenue increased 23% YoY, exceeding forecast by $2.3M
• Top 3 products account for 67% of total revenue
• Customer acquisition cost decreased 15% in Q4

Bullet Points:"""

        response = self.client.chat.completions.create(
            model=self.balanced_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
            max_tokens=300
        )
        
        content = response.choices[0].message.content.strip()
        
        # Parse bullet points
        bullets = [line.strip('• -').strip() 
                  for line in content.split('\n') 
                  if line.strip() and ('•' in line or line.strip().startswith('-'))]
        
        return bullets[:num_points]
    
    def detect_trends(self, df: pd.DataFrame, time_column: str, 
                     value_columns: List[str]) -> Dict[str, str]:
        """Detect trends in time series data"""
        
        trends = {}
        
        for col in value_columns:
            if col not in df.columns or df[col].dtype not in ['int64', 'float64']:
                continue
            
            # Calculate trend
            values = df[col].dropna()
            if len(values) < 2:
                continue
            
            # Simple trend calculation
            first_half = values[:len(values)//2].mean()
            second_half = values[len(values)//2:].mean()
            
            if pd.isna(first_half) or pd.isna(second_half) or first_half == 0:
                continue
            
            change = ((second_half - first_half) / first_half) * 100
            
            # Get AI analysis
            data_summary = f"{col}: {values.to_list()}"
            
            prompt = f"""Analyze this trend data:

Metric: {col}
Values: {data_summary}
Change: {change:.1f}%

Provide a one-sentence insight (max 20 words) about the trend.
Include the specific percentage change.

Example: "Shows strong upward trend with 23% growth, accelerating in recent periods"

Insight:"""

            response = self.client.chat.completions.create(
                model=self.fast_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5,
                max_tokens=100
            )
            
            trends[col] = response.choices[0].message.content.strip()
        
        return trends
    
    def generate_executive_summary(self, all_slides_data: List[Dict]) -> Dict[str, any]:
        """Generate executive summary from all slides"""
        
        # Combine all insights
        all_insights = []
        for slide in all_slides_data:
            all_insights.append({
                'title': slide.get('title', ''),
                'data': slide.get('summary', '')
            })
        
        combined = "\n".join([f"{i['title']}: {i['data']}" 
                             for i in all_insights[:10]])
        
        prompt = f"""You are creating an executive summary for a data presentation.

Presentation Data:
{combined}

Create a comprehensive executive summary with:
1. Top 3 key findings (specific numbers)
2. Main recommendation
3. Critical metrics to watch
4. Risks or opportunities

Format as JSON:
{
    "key_findings": ["finding 1", "finding 2", "finding 3"],
    "recommendation": "main recommendation",
    "critical_metrics": ["metric 1", "metric 2"],
    "risks_opportunities": "brief risk/opportunity statement"
}

Response:"""

        response = self.client.chat.completions.create(
            model=self.capable_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
            max_tokens=500
        )
        
        try:
            summary = json.loads(response.choices[0].message.content)
        except:
            # Fallback if JSON parsing fails
            content = response.choices[0].message.content
            summary = {
                "key_findings": [content],
                "recommendation": "See detailed slides for analysis",
                "critical_metrics": [],
                "risks_opportunities": ""
            }
        
        return summary
    
    def generate_recommendations(self, df: pd.DataFrame, 
                                chart_type: str) -> List[str]:
        """Generate actionable recommendations"""
        
        stats = self._calculate_statistics(df)
        data_summary = self._summarize_dataframe(df, max_rows=8)
        
        prompt = f"""You are a business consultant analyzing data.

Chart Type: {chart_type}
Data:
{data_summary}

Statistics:
{json.dumps(stats, indent=2)}

Generate 3 specific, actionable recommendations (each max 15 words).
Focus on business impact and next steps.

Format: Start each with •

Example:
• Increase marketing spend in Q4 when conversion is 2x higher
• Investigate Product C's declining trend - consider phase-out
• Expand APAC team by 3 members to capture 34% growth opportunity

Recommendations:"""

        response = self.client.chat.completions.create(
            model=self.balanced_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.7,
            max_tokens=250
        )
        
        content = response.choices[0].message.content.strip()
        
        recommendations = [line.strip('• -').strip() 
                          for line in content.split('\n') 
                          if line.strip() and ('•' in line or line.strip().startswith('-'))]
        
        return recommendations[:3]
    
    def detect_anomalies(self, df: pd.DataFrame, column: str) -> List[Dict]:
        """Detect unusual data points"""
        
        if column not in df.columns:
            return []
        
        values = df[column].dropna()
        if len(values) < 5:
            return []
        
        # Calculate mean and std
        mean = values.mean()
        std = values.std()
        
        # Find outliers (values > 2 standard deviations)
        anomalies = []
        for idx, value in values.items():
            if abs(value - mean) > 2 * std:
                anomalies.append({
                    'index': int(idx),
                    'value': float(value),
                    'deviation': float((value - mean) / std)
                })
        
        if not anomalies:
            return []
        
        # Get AI explanation
        anomaly_str = ", ".join([f"Row {a['index']}: {a['value']:.2f}" 
                                for a in anomalies[:3]])
        
        prompt = f"""Analyze these data anomalies:

Column: {column}
Average: {mean:.2f}
Std Dev: {std:.2f}
Anomalies: {anomaly_str}

Provide a brief explanation (max 20 words) of what might cause these outliers.

Explanation:"""

        response = self.client.chat.completions.create(
            model=self.fast_model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.6,
            max_tokens=100
        )
        
        explanation = response.choices[0].message.content.strip()
        
        return [{
            **anomaly,
            'explanation': explanation
        } for anomaly in anomalies[:5]]
    
    def _summarize_dataframe(self, df: pd.DataFrame, max_rows: int = 5) -> str:
        """Create text summary of dataframe"""
        
        summary_parts = []
        summary_parts.append(f"Rows: {len(df)}, Columns: {len(df.columns)}")
        
        # Column types
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        if numeric_cols:
            summary_parts.append(f"Numeric columns: {', '.join(numeric_cols[:5])}")
        
        # Sample data
        sample = df.head(max_rows).to_string(index=False)
        summary_parts.append(f"\nSample data:\n{sample}")
        
        return "\n".join(summary_parts)
    
    def _calculate_statistics(self, df: pd.DataFrame) -> Dict:
        """Calculate basic statistics"""
        
        stats = {}
        numeric_cols = df.select_dtypes(include=['number']).columns
        
        for col in numeric_cols:
            values = df[col].dropna()
            if len(values) == 0:
                continue
            
            stats[col] = {
                'mean': float(values.mean()),
                'min': float(values.min()),
                'max': float(values.max()),
                'std': float(values.std()) if len(values) > 1 else 0
            }
            
            # Calculate growth if enough data
            if len(values) >= 2:
                first = values.iloc[0]
                last = values.iloc[-1]
                if first != 0:
                    growth = ((last - first) / first) * 100
                    stats[col]['growth'] = float(growth)
        
        return stats


# Helper function for easy use
def create_ai_service() -> AIInsightsService:
    """Create AI service instance"""
    return AIInsightsService()
```

---

## 📝 Usage Examples

### **Example 1: Generate Insights for Single Slide**

```python
from services.ai_service import create_ai_service
import pandas as pd

# Your data
df = pd.DataFrame({
    'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'Revenue': [100000, 120000, 145000, 178000],
    'Profit': [20000, 25000, 32000, 42000]
})

# Create AI service
ai = create_ai_service()

# Generate title
title = ai.generate_slide_title(df, 'line', ['Quarter', 'Revenue'])
print(f"Title: {title}")
# Output: "Revenue Surged 78% to $178K in Q4"

# Generate bullet points
bullets = ai.generate_bullet_points(df, 'line', num_points=5)
for bullet in bullets:
    print(f"• {bullet}")
# Output:
# • Revenue grew 78% from Q1 to Q4, reaching $178K
# • Profit margin improved from 20% to 24%
# • Q4 revenue exceeded Q3 by 23%, showing acceleration
# • All quarters showed positive growth
# • Average quarterly growth rate: 21%

# Detect trends
trends = ai.detect_trends(df, 'Quarter', ['Revenue', 'Profit'])
print(trends)
# Output: {'Revenue': 'Strong upward trend with 21% average quarterly growth',
#          'Profit': 'Accelerating growth with improving margins'}

# Generate recommendations
recommendations = ai.generate_recommendations(df, 'line')
for rec in recommendations:
    print(f"→ {rec}")
# Output:
# → Maintain current growth strategy - results consistently exceed targets
# → Invest in Q4 initiatives as they show highest ROI
# → Forecast $215K revenue for Q1 next year assuming current trend
```

### **Example 2: Full Presentation with AI**

```python
from converter.excel_reader import excel_reader
from converter.ppt_writer import create_presentation
from converter.chart_detector import detect_chart_type
from services.ai_service import create_ai_service

def create_ai_powered_presentation(excel_file, output_file):
    """Create presentation with full AI insights"""
    
    # Read Excel
    df = excel_reader(excel_file)
    
    # Initialize AI
    ai = create_ai_service()
    
    # Create presentation
    prs = create_presentation("Quarterly Report", "AI-Generated Insights")
    
    # Detect chart
    chart_type, config = detect_chart_type(df)
    
    # Generate AI insights
    title = ai.generate_slide_title(df, chart_type, list(df.columns))
    bullets = ai.generate_bullet_points(df, chart_type)
    trends = ai.detect_trends(df, config.get('x_col'), config.get('y_cols', []))
    recommendations = ai.generate_recommendations(df, chart_type)
    
    # Add slide with chart + AI insights
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_box.text = title
    
    # Add chart (left side)
    # ... create chart code ...
    
    # Add AI insights (right side)
    insights_box = slide.shapes.add_textbox(Inches(6), Inches(1.5), Inches(3.5), Inches(5))
    text_frame = insights_box.text_frame
    
    # Add bullet points
    for bullet in bullets:
        p = text_frame.add_paragraph()
        p.text = f"• {bullet}"
        p.level = 0
    
    # Add trends
    if trends:
        text_frame.add_paragraph().text = "\nKey Trends:"
        for col, trend in trends.items():
            p = text_frame.add_paragraph()
            p.text = f"  {col}: {trend}"
    
    # Add recommendations
    if recommendations:
        text_frame.add_paragraph().text = "\nRecommendations:"
        for rec in recommendations:
            p = text_frame.add_paragraph()
            p.text = f"→ {rec}"
    
    # Save
    prs.save(output_file)
    print(f"✅ AI-powered presentation created: {output_file}")
```

---

## 🔌 Backend Integration

### **Add AI Endpoint**

```python
# src/backend/app/api/v1/endpoints/ai_insights.py

from fastapi import APIRouter, Depends, HTTPException
from typing import List, Dict
import pandas as pd

from services.ai_service import create_ai_service
from api.deps import get_current_active_user
from models.user import UserInDB

router = APIRouter()

@router.post("/ai/analyze")
async def analyze_data(
    file_id: str,
    features: List[str] = ["title", "bullets", "trends", "recommendations"],
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Generate AI insights for uploaded Excel file
    
    Features:
    - title: Generate slide title
    - bullets: Key insights bullet points
    - trends: Trend detection
    - recommendations: Actionable recommendations
    - anomalies: Unusual data points
    - summary: Executive summary
    """
    
    try:
        # Get file
        file_doc = await get_file_by_id(file_id, str(current_user.id))
        if not file_doc:
            raise HTTPException(status_code=404, detail="File not found")
        
        # Read Excel
        df = pd.read_excel(file_doc.storage_path)
        
        # Initialize AI
        ai = create_ai_service()
        
        # Generate requested insights
        insights = {}
        
        if "title" in features:
            chart_type = "line"  # Default, or detect from data
            insights["title"] = ai.generate_slide_title(df, chart_type, list(df.columns))
        
        if "bullets" in features:
            insights["bullet_points"] = ai.generate_bullet_points(df, "line")
        
        if "trends" in features:
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            if len(numeric_cols) > 0:
                insights["trends"] = ai.detect_trends(df, df.columns[0], numeric_cols[:3])
        
        if "recommendations" in features:
            insights["recommendations"] = ai.generate_recommendations(df, "line")
        
        if "anomalies" in features:
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            if len(numeric_cols) > 0:
                insights["anomalies"] = ai.detect_anomalies(df, numeric_cols[0])
        
        return {
            "file_id": file_id,
            "insights": insights,
            "status": "success"
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/ai/generate-presentation")
async def generate_ai_presentation(
    file_id: str,
    template: str = "corporate_blue",
    include_ai: bool = True,
    current_user: UserInDB = Depends(get_current_active_user)
):
    """Generate full presentation with AI insights"""
    
    # ... implementation ...
    pass
```

### **Register AI Router**

```python
# src/backend/app/api/v1/api.py

from api.v1.endpoints import ai_insights

api_router.include_router(
    ai_insights.router,
    prefix="/ai",
    tags=["ai-insights"]
)
```

---

## 💰 Pricing & Costs

### **Groq API Pricing:**
- **Free Tier:** 14,400 requests/day
- **Paid:** $0.27 per 1M tokens (input), $0.27 per 1M tokens (output)

### **Cost Per Presentation:**
- Title: ~100 tokens = $0.00003
- Bullets (5): ~500 tokens = $0.00015
- Trends (3): ~300 tokens = $0.00009
- Recommendations (3): ~300 tokens = $0.00009
- **Total: ~$0.00036 per presentation**

### **Your Costs:**
- 1,000 presentations/month = $0.36
- 10,000 presentations/month = $3.60
- 100,000 presentations/month = $36

### **Your Pricing:**
- AI Insights add-on: $20/month
- Profit per user: $19.64/month (99.8% margin!)

---

## 📊 Feature Rollout Plan

### **Week 1:**
✅ Setup Groq API
✅ Implement slide title generation
✅ Implement bullet points
✅ Test with sample data

### **Week 2:**
✅ Add trend detection
✅ Add recommendations
✅ Add anomaly detection
✅ Create API endpoints

### **Week 3:**
✅ Frontend integration
✅ Add toggle for AI features
✅ Add loading states
✅ User testing

### **Week 4:**
✅ Executive summary
✅ Performance optimization
✅ Error handling
✅ Production deployment

---

## 🚀 Next Steps

1. **Get Groq API Key** (5 minutes)
2. **Create AI service file** (copy code above)
3. **Test with sample data** (30 minutes)
4. **Add API endpoints** (2 hours)
5. **Frontend integration** (4 hours)
6. **Launch AI features!** 🎉

**YOU'LL HAVE WORKING AI IN 1 WEEK!** 💪
