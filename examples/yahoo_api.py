import yfinance as yf
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
from pathlib import Path

class FinancialExcelGenerator:
    def __init__(self, output_folder="example/Company_Data"):
        """Initialize the generator with output folder path"""
        self.output_folder = output_folder
        # Create folder if it doesn't exist
        Path(self.output_folder).mkdir(parents=True, exist_ok=True)
        
    def fetch_company_data(self, ticker):
        """Fetch all financial data for a company"""
        print(f"📊 Fetching data for {ticker}...")
        
        try:
            stock = yf.Ticker(ticker)
            
            # Get various data
            info = stock.info
            hist = stock.history(period="1mo")
            
            # Remove timezone from datetime index
            if not hist.empty and hist.index.tz is not None:
                hist.index = hist.index.tz_localize(None)
            
            financials = stock.financials
            balance_sheet = stock.balance_sheet
            cashflow = stock.cashflow
            
            return {
                'info': info,
                'history': hist,
                'financials': financials,
                'balance_sheet': balance_sheet,
                'cashflow': cashflow,
                'ticker': ticker
            }
        except Exception as e:
            print(f"❌ Error fetching data for {ticker}: {str(e)}")
            return None
    
    def sanitize_value(self, value):
        """Convert any value to Excel-safe format"""
        # Handle None/NaN
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return 'N/A'
        
        # Handle pandas Timestamp
        if isinstance(value, pd.Timestamp):
            if value.tzinfo is not None:
                value = value.tz_localize(None)
            return value.strftime('%Y-%m-%d %H:%M:%S')
        
        # Handle datetime
        if isinstance(value, datetime):
            if value.tzinfo is not None:
                value = value.replace(tzinfo=None)
            return value.strftime('%Y-%m-%d %H:%M:%S')
        
        # Handle numeric values
        if isinstance(value, (int, float)):
            if pd.isna(value):
                return 'N/A'
            return value
        
        # Handle strings
        if isinstance(value, str):
            # Limit very long strings
            if len(value) > 32000:  # Excel cell limit
                return value[:32000]
            return value
        
        # Default: convert to string
        try:
            return str(value)
        except:
            return 'N/A'
    
    def write_rows_safe(self, ws, data_rows):
        """Write rows to worksheet with sanitized values"""
        for row in data_rows:
            safe_row = [self.sanitize_value(cell) for cell in row]
            ws.append(safe_row)
    
    def set_column_width(self, ws, col_num, width):
        """Safely set column width"""
        try:
            col_letter = get_column_letter(col_num)
            ws.column_dimensions[col_letter].width = width
        except Exception as e:
            print(f"    Warning: Could not set width for column {col_num}: {e}")
    
    def create_overview_sheet(self, wb, data):
        """Create Company Overview sheet"""
        ws = wb.create_sheet("Overview")
        info = data['info']
        
        # Header styling
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        
        # Add data
        overview_data = [
            ["Company Overview", ""],
            ["Symbol", data['ticker']],
            ["Company Name", info.get('longName', 'N/A')],
            ["Sector", info.get('sector', 'N/A')],
            ["Industry", info.get('industry', 'N/A')],
            ["Website", info.get('website', 'N/A')],
            ["Country", info.get('country', 'N/A')],
            ["Employees", info.get('fullTimeEmployees', 'N/A')],
            ["", ""],
            ["Market Data", ""],
            ["Current Price", f"${info.get('currentPrice', 'N/A')}"],
            ["Previous Close", f"${info.get('previousClose', 'N/A')}"],
            ["Market Cap", info.get('marketCap', 'N/A')],
            ["Enterprise Value", info.get('enterpriseValue', 'N/A')],
            ["52 Week High", f"${info.get('fiftyTwoWeekHigh', 'N/A')}"],
            ["52 Week Low", f"${info.get('fiftyTwoWeekLow', 'N/A')}"],
            ["Volume", info.get('volume', 'N/A')],
            ["Average Volume", info.get('averageVolume', 'N/A')],
            ["", ""],
            ["Company Description", ""],
            [info.get('longBusinessSummary', 'N/A'), ""],
        ]
        
        self.write_rows_safe(ws, overview_data)
        
        # Style the sheet
        self.set_column_width(ws, 1, 30)
        self.set_column_width(ws, 2, 50)
        ws['A1'].fill = header_fill
        ws['A1'].font = header_font
        
        return ws
    
    def create_price_history_sheet(self, wb, data):
        """Create Price History sheet"""
        ws = wb.create_sheet("Price History")
        hist = data['history'].copy()
        
        if hist.empty:
            ws.append(["No price history available"])
            return ws
        
        # Reset index to make Date a column
        hist_reset = hist.reset_index()
        
        # Sanitize all data
        ws.append(["Stock Price History - Last 30 Days"])
        ws.append([])
        
        # Write header
        header = [str(col) for col in hist_reset.columns]
        ws.append(header)
        
        # Write data rows
        for idx, row in hist_reset.iterrows():
            safe_row = [self.sanitize_value(val) for val in row.values]
            ws.append(safe_row)
        
        # Set column widths
        for col_num in range(1, len(header) + 1):
            self.set_column_width(ws, col_num, 15)
        
        return ws
    
    def create_key_statistics_sheet(self, wb, data):
        """Create Key Statistics sheet"""
        ws = wb.create_sheet("Key Statistics")
        info = data['info']
        
        stats_data = [
            ["Key Statistics", ""],
            ["", ""],
            ["Valuation Metrics", ""],
            ["P/E Ratio (TTM)", info.get('trailingPE', 'N/A')],
            ["Forward P/E", info.get('forwardPE', 'N/A')],
            ["PEG Ratio", info.get('pegRatio', 'N/A')],
            ["Price/Sales (TTM)", info.get('priceToSalesTrailing12Months', 'N/A')],
            ["Price/Book", info.get('priceToBook', 'N/A')],
            ["Enterprise Value/Revenue", info.get('enterpriseToRevenue', 'N/A')],
            ["Enterprise Value/EBITDA", info.get('enterpriseToEbitda', 'N/A')],
            ["", ""],
            ["Profitability", ""],
            ["Profit Margin", f"{info.get('profitMargins', 0) * 100:.2f}%" if info.get('profitMargins') else 'N/A'],
            ["Operating Margin", f"{info.get('operatingMargins', 0) * 100:.2f}%" if info.get('operatingMargins') else 'N/A'],
            ["Return on Assets", f"{info.get('returnOnAssets', 0) * 100:.2f}%" if info.get('returnOnAssets') else 'N/A'],
            ["Return on Equity", f"{info.get('returnOnEquity', 0) * 100:.2f}%" if info.get('returnOnEquity') else 'N/A'],
            ["Revenue Growth", f"{info.get('revenueGrowth', 0) * 100:.2f}%" if info.get('revenueGrowth') else 'N/A'],
            ["Earnings Growth", f"{info.get('earningsGrowth', 0) * 100:.2f}%" if info.get('earningsGrowth') else 'N/A'],
            ["", ""],
            ["Financial Health", ""],
            ["Total Cash", info.get('totalCash', 'N/A')],
            ["Total Debt", info.get('totalDebt', 'N/A')],
            ["Debt to Equity", info.get('debtToEquity', 'N/A')],
            ["Current Ratio", info.get('currentRatio', 'N/A')],
            ["Quick Ratio", info.get('quickRatio', 'N/A')],
            ["Free Cash Flow", info.get('freeCashflow', 'N/A')],
            ["", ""],
            ["Trading Information", ""],
            ["Beta", info.get('beta', 'N/A')],
            ["50-Day Moving Average", f"${info.get('fiftyDayAverage', 'N/A')}"],
            ["200-Day Moving Average", f"${info.get('twoHundredDayAverage', 'N/A')}"],
            ["Shares Outstanding", info.get('sharesOutstanding', 'N/A')],
            ["Float Shares", info.get('floatShares', 'N/A')],
            ["Shares Short", info.get('sharesShort', 'N/A')],
            ["Short Ratio", info.get('shortRatio', 'N/A')],
            ["Short % of Float", f"{info.get('shortPercentOfFloat', 0) * 100:.2f}%" if info.get('shortPercentOfFloat') else 'N/A'],
        ]
        
        self.write_rows_safe(ws, stats_data)
        
        self.set_column_width(ws, 1, 35)
        self.set_column_width(ws, 2, 20)
        
        return ws
    
    def create_dividends_sheet(self, wb, data):
        """Create Dividends & Splits sheet"""
        ws = wb.create_sheet("Dividends & Splits")
        info = data['info']
        
        dividend_data = [
            ["Dividend Information", ""],
            ["Dividend Rate", info.get('dividendRate', 'N/A')],
            ["Dividend Yield", f"{info.get('dividendYield', 0) * 100:.2f}%" if info.get('dividendYield') else 'N/A'],
            ["Ex-Dividend Date", info.get('exDividendDate', 'N/A')],
            ["Payout Ratio", f"{info.get('payoutRatio', 0) * 100:.2f}%" if info.get('payoutRatio') else 'N/A'],
            ["5 Year Avg Dividend Yield", info.get('fiveYearAvgDividendYield', 'N/A')],
            ["", ""],
            ["Stock Split Information", ""],
            ["Last Split Factor", info.get('lastSplitFactor', 'N/A')],
            ["Last Split Date", info.get('lastSplitDate', 'N/A')],
        ]
        
        self.write_rows_safe(ws, dividend_data)
        
        self.set_column_width(ws, 1, 30)
        self.set_column_width(ws, 2, 20)
        
        return ws
    
    def create_financials_sheet(self, wb, data, sheet_name, df):
        """Create financial statement sheet (Income/Balance/Cashflow)"""
        ws = wb.create_sheet(sheet_name)
        
        if df is None or df.empty:
            ws.append([f"No {sheet_name} data available"])
            return ws
        
        try:
            # Transpose so dates are columns
            df_transposed = df.T.copy()
            
            # Sanitize column names (dates) - convert ALL to strings
            sanitized_columns = []
            for col in df_transposed.columns:
                if isinstance(col, pd.Timestamp):
                    if col.tzinfo is not None:
                        col = col.tz_localize(None)
                    sanitized_columns.append(col.strftime('%Y-%m-%d'))
                elif isinstance(col, datetime):
                    if col.tzinfo is not None:
                        col = col.replace(tzinfo=None)
                    sanitized_columns.append(col.strftime('%Y-%m-%d'))
                else:
                    sanitized_columns.append(str(col))
            
            ws.append([f"{sheet_name} Statement"])
            ws.append([])
            
            # Write header with sanitized column names
            header = ['Metric'] + sanitized_columns
            ws.append(header)
            
            # Write data rows with sanitized values
            for idx, row in df_transposed.iterrows():
                row_data = [str(idx)] + [self.sanitize_value(val) for val in row.values]
                ws.append(row_data)
            
            # Set column widths safely
            for col_num in range(1, len(header) + 1):
                self.set_column_width(ws, col_num, 20)
                
        except Exception as e:
            print(f"    Warning: Issue with {sheet_name} sheet: {str(e)}")
            ws.append([f"Error processing {sheet_name} data"])
        
        return ws
    
    def generate_excel(self, ticker):
        """Generate complete Excel file for a company"""
        data = self.fetch_company_data(ticker)
        
        if data is None:
            return False
        
        try:
            # Create workbook
            wb = Workbook()
            # Remove default sheet
            wb.remove(wb.active)
            
            # Create all sheets
            print(f"  ├─ Creating Overview sheet...")
            self.create_overview_sheet(wb, data)
            
            print(f"  ├─ Creating Price History sheet...")
            self.create_price_history_sheet(wb, data)
            
            print(f"  ├─ Creating Key Statistics sheet...")
            self.create_key_statistics_sheet(wb, data)
            
            print(f"  ├─ Creating Dividends & Splits sheet...")
            self.create_dividends_sheet(wb, data)
            
            print(f"  ├─ Creating Income Statement sheet...")
            self.create_financials_sheet(wb, data, "Income Statement", data['financials'])
            
            print(f"  ├─ Creating Balance Sheet...")
            self.create_financials_sheet(wb, data, "Balance Sheet", data['balance_sheet'])
            
            print(f"  ├─ Creating Cash Flow sheet...")
            self.create_financials_sheet(wb, data, "Cash Flow", data['cashflow'])
            
            # Save file
            filename = f"{ticker}_Financial_Data.xlsx"
            filepath = os.path.join(self.output_folder, filename)
            
            print(f"  ├─ Saving file...")
            wb.save(filepath)
            
            print(f"  ✓ Saved: {filepath}")
            return True
            
        except Exception as e:
            print(f"  ❌ Error creating Excel file: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def generate_multiple(self, tickers):
        """Generate Excel files for multiple companies"""
        print(f"\n🚀 Starting Excel generation for {len(tickers)} companies\n")
        print("=" * 60)
        
        success_count = 0
        failed = []
        
        for i, ticker in enumerate(tickers, 1):
            print(f"\n[{i}/{len(tickers)}] Processing {ticker}...")
            if self.generate_excel(ticker):
                success_count += 1
            else:
                failed.append(ticker)
            print()
        
        print("=" * 60)
        print(f"\n✅ Successfully generated {success_count}/{len(tickers)} Excel files")
        
        if failed:
            print(f"❌ Failed: {', '.join(failed)}")
        
        print(f"\n📁 Files saved in: {os.path.abspath(self.output_folder)}")


def main():
    # Example usage
    companies = ['AAPL', 'MSFT', 'GOOGL', 'TSLA', 'AMZN']
    
    # You can also add custom companies
    print("Financial Data Excel Generator")
    print("=" * 60)
    
    use_default = input("\nUse default companies (AAPL, MSFT, GOOGL, TSLA, AMZN)? (y/n): ").lower()
    
    if use_default != 'y':
        custom_input = input("\nEnter company tickers separated by commas (e.g., NVDA,META,JPM): ").strip()
        if custom_input:
            companies = [ticker.strip().upper() for ticker in custom_input.split(',') if ticker.strip()]
        else:
            print("No tickers entered. Using default companies.")
    
    if not companies:
        print("Error: No companies to process!")
        return
    
    print(f"\n📋 Companies to process: {', '.join(companies)}")
    
    # Generate Excel files
    generator = FinancialExcelGenerator(output_folder="example/Company_Data")
    generator.generate_multiple(companies)
    
    print("\n✨ Done! Check the 'example/Company_Data' folder for your Excel files.\n")


if __name__ == "__main__":
    main()