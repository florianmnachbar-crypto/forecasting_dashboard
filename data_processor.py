"""
Data Processor for Amazon Haul EU5 Forecasting Dashboard
Handles Excel file parsing and data transformation
v2.0.0 - Adds manual forecast support
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re


class DataProcessor:
    """Processes Excel input files for the forecasting dashboard"""
    
    METRICS = ['Net Ordered Units', 'Transits', 'Transit Conversion', 'UPO']
    # Map alternative metric names
    METRIC_ALIASES = {
        'CVR': 'Transit Conversion',
        'Conversion': 'Transit Conversion',
    }
    MARKETPLACES = ['UK', 'DE', 'FR', 'IT', 'ES', 'EU5']
    
    def __init__(self):
        self.data = {}  # Actuals data
        self.manual_forecast = {}  # Manual forecast data
        self.weeks = []
        self.forecast_weeks = []
        self.raw_df = None
        self.has_manual_forecast = False
        
    def parse_week_column(self, col_name):
        """Convert week column name to datetime"""
        if not isinstance(col_name, str):
            col_name = str(col_name)
        
        col_name = col_name.strip()
        
        # Match patterns like "Wk19 2025", "Wk 1 2026", "Wk19 2025"
        match = re.match(r'Wk\s*(\d+)\s+(\d{4})', col_name)
        if match:
            week_num = int(match.group(1))
            year = int(match.group(2))
            # Convert ISO week to date (Monday of that week)
            try:
                date = datetime.strptime(f'{year}-W{week_num:02d}-1', '%G-W%V-%u')
                return date
            except ValueError:
                # Handle edge cases
                return datetime(year, 1, 1) + timedelta(weeks=week_num-1)
        return None
    
    def find_cell_value(self, df, search_value):
        """Find the row and column index of a value in the dataframe"""
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    if str(cell_value).strip() == search_value:
                        return row_idx, col_idx
        return None, None
    
    def load_excel(self, file_path):
        """Load and parse the Excel file"""
        try:
            # Check available sheets
            xl = pd.ExcelFile(file_path)
            sheet_names = xl.sheet_names
            print(f"Available sheets: {sheet_names}")
            
            # Determine actuals sheet name
            actuals_sheet = None
            if 'Actuals' in sheet_names:
                actuals_sheet = 'Actuals'
            elif 'Sheet1' in sheet_names:
                actuals_sheet = 'Sheet1'
            else:
                actuals_sheet = sheet_names[0]
            
            # Load actuals
            df = pd.read_excel(file_path, sheet_name=actuals_sheet, header=None)
            self.raw_df = df
            
            print(f"Actuals sheet '{actuals_sheet}': {df.shape[0]} rows, {df.shape[1]} columns")
            
            # Find all metric sections for actuals
            self.data = {}
            
            for metric in self.METRICS:
                metric_data = self._parse_metric_section(df, metric)
                if metric_data:
                    self.data[metric] = metric_data
                    print(f"Parsed {metric}: {list(metric_data.keys())}")
            
            if not self.data:
                return False, "No data sections found in the Excel file"
            
            # Recalculate EU5 totals for actuals
            self.calculate_eu5_totals()
            
            # Load manual forecast if available
            if 'Forecast' in sheet_names:
                self._load_manual_forecast(file_path, 'Forecast')
            
            return True, "File loaded successfully"
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"Error loading file: {str(e)}"
    
    def _load_manual_forecast(self, file_path, sheet_name):
        """Load manual forecast from a separate sheet"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            print(f"Forecast sheet '{sheet_name}': {df.shape[0]} rows, {df.shape[1]} columns")
            
            self.manual_forecast = {}
            
            # Search for metrics, including aliases
            search_metrics = self.METRICS + list(self.METRIC_ALIASES.keys())
            
            for search_metric in search_metrics:
                # Map alias to standard name
                standard_metric = self.METRIC_ALIASES.get(search_metric, search_metric)
                
                # Skip if already parsed
                if standard_metric in self.manual_forecast:
                    continue
                
                metric_data = self._parse_metric_section(df, search_metric, is_forecast=True)
                if metric_data:
                    self.manual_forecast[standard_metric] = metric_data
                    print(f"Parsed manual forecast {standard_metric}: {list(metric_data.keys())}")
            
            if self.manual_forecast:
                self.has_manual_forecast = True
                # Recalculate EU5 for manual forecast
                self.calculate_eu5_totals(is_forecast=True)
                print("Manual forecast loaded successfully")
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Warning: Could not load manual forecast: {str(e)}")
    
    def _parse_metric_section(self, df, metric_name, is_forecast=False):
        """Parse a single metric section from the dataframe"""
        # Find the metric header row
        metric_row, metric_col = self.find_cell_value(df, metric_name)
        
        if metric_row is None:
            # print(f"Could not find metric: {metric_name}")
            return None
        
        print(f"Found {metric_name} at row {metric_row}, col {metric_col}")
        
        # Find the MP header row (should be 1 row after metric name)
        mp_row = None
        header_col = None
        week_start_col = None
        
        # Look for 'MP' in the rows after the metric name
        for search_row in range(metric_row + 1, min(metric_row + 3, len(df))):
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[search_row, col_idx]
                if pd.notna(cell_value) and str(cell_value).strip() == 'MP':
                    mp_row = search_row
                    header_col = col_idx
                    week_start_col = col_idx + 1
                    break
            if mp_row is not None:
                break
        
        if mp_row is None:
            print(f"Could not find MP header for {metric_name}")
            return None
        
        # Parse week columns from the MP row
        weeks = []
        for col_idx in range(week_start_col, len(df.columns)):
            col_value = df.iloc[mp_row, col_idx]
            if pd.notna(col_value):
                week_str = str(col_value).strip()
                if self.parse_week_column(week_str):
                    weeks.append(week_str)
                else:
                    break  # Stop at first non-week column
            else:
                # Allow empty columns up to a point
                if col_idx - week_start_col > 100:
                    break
        
        if not weeks:
            print(f"No week columns found for {metric_name}")
            return None
        
        # Store weeks for this data type
        if is_forecast:
            if not self.forecast_weeks:
                self.forecast_weeks = weeks
        else:
            if not self.weeks:
                self.weeks = weeks
        
        print(f"Found {len(weeks)} weeks")
        
        # Parse marketplace data rows
        metric_data = {}
        
        # Look for marketplace rows after the MP header
        for row_idx in range(mp_row + 1, min(mp_row + 10, len(df))):
            mp_value = df.iloc[row_idx, header_col]
            
            if pd.isna(mp_value):
                continue
            
            mp_name = str(mp_value).strip()
            
            if mp_name in self.MARKETPLACES:
                # Extract values for this marketplace
                values = []
                for col_idx in range(week_start_col, week_start_col + len(weeks)):
                    if col_idx < len(df.columns):
                        val = df.iloc[row_idx, col_idx]
                        # Handle invalid values
                        if pd.isna(val) or str(val) == '[object Object]' or str(val) == 'nan':
                            values.append(np.nan)
                        else:
                            try:
                                values.append(float(val))
                            except (ValueError, TypeError):
                                values.append(np.nan)
                    else:
                        values.append(np.nan)
                
                # Count valid values
                valid_count = sum(1 for v in values if not np.isnan(v))
                print(f"  {mp_name}: {valid_count} valid data points")
                
                metric_data[mp_name] = values
            elif mp_name in self.METRICS or mp_name in self.METRIC_ALIASES:
                # We've reached the next metric section
                break
        
        return metric_data if metric_data else None
    
    def get_dataframe(self, metric, marketplace, is_forecast=False):
        """Get a pandas DataFrame for a specific metric and marketplace"""
        data_source = self.manual_forecast if is_forecast else self.data
        weeks_source = self.forecast_weeks if is_forecast else self.weeks
        
        if metric not in data_source or marketplace not in data_source[metric]:
            return None
        
        values = data_source[metric][marketplace]
        
        # Use only as many weeks as we have values
        week_count = min(len(values), len(weeks_source))
        dates = [self.parse_week_column(weeks_source[i]) for i in range(week_count)]
        
        df = pd.DataFrame({
            'ds': dates[:len(values)],
            'y': values[:len(dates)],
            'week': weeks_source[:len(values)]
        })
        
        # Remove rows with NaN values for cleaner data
        df = df.dropna(subset=['y'])
        df = df[df['ds'].notna()]
        
        return df
    
    def format_week_label(self, date):
        """Format date as week label (Wk## YYYY) - Sunday to Saturday weeks"""
        if date is None:
            return None
        # Get ISO week but adjust for Sunday start
        # Add 1 day to shift Sunday to be the start of the week
        adjusted = date + timedelta(days=1)
        week_num = adjusted.isocalendar()[1]
        year = adjusted.isocalendar()[0]
        return f"Wk{week_num:02d} {year}"
    
    def get_all_data(self):
        """Get all data in a structured format for the frontend"""
        result = {}
        
        for metric in self.METRICS:
            if metric not in self.data:
                continue
                
            result[metric] = {}
            for mp in self.MARKETPLACES:
                if mp not in self.data[metric]:
                    continue
                    
                df = self.get_dataframe(metric, mp)
                if df is not None and not df.empty:
                    result[metric][mp] = {
                        'dates': df['ds'].dt.strftime('%Y-%m-%d').tolist(),
                        'values': df['y'].tolist(),
                        'weeks': [self.format_week_label(d) for d in df['ds']],
                        'week_labels': df['week'].tolist()  # Original labels from Excel
                    }
        
        return result
    
    def get_manual_forecast_data(self):
        """Get manual forecast data in a structured format for the frontend"""
        if not self.has_manual_forecast:
            return None
        
        result = {}
        
        for metric in self.METRICS:
            if metric not in self.manual_forecast:
                continue
                
            result[metric] = {}
            for mp in self.MARKETPLACES:
                if mp not in self.manual_forecast[metric]:
                    continue
                    
                df = self.get_dataframe(metric, mp, is_forecast=True)
                if df is not None and not df.empty:
                    result[metric][mp] = {
                        'dates': df['ds'].dt.strftime('%Y-%m-%d').tolist(),
                        'values': df['y'].tolist(),
                        'weeks': [self.format_week_label(d) for d in df['ds']],
                        'week_labels': df['week'].tolist()
                    }
        
        return result
    
    def calculate_forecast_accuracy(self, metric, marketplace):
        """Calculate accuracy metrics for manual forecast vs actuals
        
        Returns MAPE, WMAPE, Bias, and Accuracy for overlapping periods
        """
        if not self.has_manual_forecast:
            return None
        
        # Get both datasets
        actuals_df = self.get_dataframe(metric, marketplace, is_forecast=False)
        forecast_df = self.get_dataframe(metric, marketplace, is_forecast=True)
        
        if actuals_df is None or forecast_df is None:
            return None
        
        if actuals_df.empty or forecast_df.empty:
            return None
        
        # Merge on date to find overlapping periods
        merged = pd.merge(
            actuals_df[['ds', 'y']].rename(columns={'y': 'actual'}),
            forecast_df[['ds', 'y']].rename(columns={'y': 'forecast'}),
            on='ds',
            how='inner'
        )
        
        if merged.empty or len(merged) < 1:
            return None
        
        # Calculate metrics
        merged['abs_error'] = abs(merged['actual'] - merged['forecast'])
        merged['abs_pct_error'] = merged['abs_error'] / merged['actual'].replace(0, np.nan) * 100
        merged['error'] = merged['forecast'] - merged['actual']
        
        # Filter out infinite values
        valid_data = merged.dropna()
        
        if valid_data.empty:
            return None
        
        # MAPE - Mean Absolute Percentage Error
        mape = valid_data['abs_pct_error'].mean()
        
        # WMAPE - Weighted MAPE (weighted by actual values)
        total_actual = valid_data['actual'].sum()
        wmape = (valid_data['abs_error'].sum() / total_actual * 100) if total_actual > 0 else np.nan
        
        # Bias - Average percentage error (positive = over-forecasting)
        bias = (valid_data['error'].sum() / total_actual * 100) if total_actual > 0 else np.nan
        
        # Accuracy
        accuracy = max(0, 100 - wmape) if not np.isnan(wmape) else np.nan
        
        return {
            'mape': round(mape, 2) if not np.isnan(mape) else None,
            'wmape': round(wmape, 2) if not np.isnan(wmape) else None,
            'bias': round(bias, 2) if not np.isnan(bias) else None,
            'accuracy': round(accuracy, 2) if not np.isnan(accuracy) else None,
            'overlap_weeks': len(valid_data),
            'total_actual': round(total_actual, 2),
            'total_forecast': round(valid_data['forecast'].sum(), 2),
            'periods': [d.strftime('%Y-%m-%d') for d in valid_data['ds']]
        }
    
    def get_all_accuracy_metrics(self):
        """Get forecast accuracy for all metrics and marketplaces"""
        if not self.has_manual_forecast:
            return None
        
        result = {}
        
        for metric in self.METRICS:
            result[metric] = {}
            for mp in self.MARKETPLACES:
                accuracy = self.calculate_forecast_accuracy(metric, mp)
                if accuracy:
                    result[metric][mp] = accuracy
        
        return result
    
    def get_summary_statistics(self, metric, marketplace):
        """Calculate summary statistics for a metric/marketplace combination"""
        df = self.get_dataframe(metric, marketplace)
        if df is None or df.empty:
            return None
        
        values = df['y'].dropna()
        if len(values) == 0:
            return None
        
        # Calculate statistics
        stats = {
            'total': round(float(values.sum()), 2),
            'average': round(float(values.mean()), 2),
            'min': round(float(values.min()), 2),
            'max': round(float(values.max()), 2),
            'count': int(len(values)),
            'last_4_week_avg': round(float(values.tail(4).mean()), 2) if len(values) >= 4 else round(float(values.mean()), 2),
            'std_dev': round(float(values.std()), 2) if len(values) > 1 else 0
        }
        
        return stats
    
    def get_latest_week_overview(self):
        """Get the latest week's data comparing actuals vs forecasts for all marketplaces and metrics
        
        Returns a dictionary with:
        - latest_week: The week label
        - latest_date: The date
        - data: Dictionary keyed by marketplace, containing all 4 metrics with:
          - actual: Actual value
          - manual_forecast: Manual forecast value (if available)
          - manual_dev: Deviation from manual forecast (%)
        """
        if not self.data:
            return None
        
        # Find the latest week across all metrics
        latest_date = None
        latest_week = None
        
        for metric in self.METRICS:
            if metric not in self.data:
                continue
            for mp in self.MARKETPLACES:
                df = self.get_dataframe(metric, mp, is_forecast=False)
                if df is not None and not df.empty:
                    last_date = df['ds'].max()
                    if latest_date is None or last_date > latest_date:
                        latest_date = last_date
                        latest_week = df[df['ds'] == last_date]['week'].iloc[0]
        
        if latest_date is None:
            return None
        
        # Build the overview data
        result = {
            'latest_week': latest_week,
            'latest_date': latest_date.strftime('%Y-%m-%d'),
            'data': {}
        }
        
        for mp in self.MARKETPLACES:
            result['data'][mp] = {}
            
            for metric in self.METRICS:
                metric_data = {
                    'actual': None,
                    'manual_forecast': None,
                    'manual_dev': None,
                    'manual_dev_pct': None
                }
                
                # Get actual value for the latest week
                df_actual = self.get_dataframe(metric, mp, is_forecast=False)
                if df_actual is not None and not df_actual.empty:
                    latest_actual = df_actual[df_actual['ds'] == latest_date]
                    if not latest_actual.empty:
                        metric_data['actual'] = float(latest_actual['y'].iloc[0])
                
                # Get manual forecast value for the latest week
                if self.has_manual_forecast:
                    df_forecast = self.get_dataframe(metric, mp, is_forecast=True)
                    if df_forecast is not None and not df_forecast.empty:
                        latest_forecast = df_forecast[df_forecast['ds'] == latest_date]
                        if not latest_forecast.empty:
                            metric_data['manual_forecast'] = float(latest_forecast['y'].iloc[0])
                
                # Calculate deviation
                if metric_data['actual'] is not None and metric_data['manual_forecast'] is not None:
                    if metric_data['manual_forecast'] != 0:
                        dev = metric_data['actual'] - metric_data['manual_forecast']
                        dev_pct = (dev / metric_data['manual_forecast']) * 100
                        metric_data['manual_dev'] = round(dev, 4)
                        metric_data['manual_dev_pct'] = round(dev_pct, 1)
                
                result['data'][mp][metric] = metric_data
        
        return result
    
    def calculate_eu5_totals(self, is_forecast=False):
        """Recalculate EU5 totals from individual marketplace data"""
        individual_mps = ['UK', 'DE', 'FR', 'IT', 'ES']
        data_source = self.manual_forecast if is_forecast else self.data
        
        for metric in self.METRICS:
            if metric not in data_source:
                continue
            
            # Get the length from existing data
            max_len = 0
            for mp in individual_mps:
                if mp in data_source[metric]:
                    max_len = max(max_len, len(data_source[metric][mp]))
            
            if max_len == 0:
                continue
            
            # Initialize EU5 with NaN
            eu5_values = [np.nan] * max_len
            valid_counts = [0] * max_len
            
            for mp in individual_mps:
                if mp not in data_source[metric]:
                    continue
                    
                mp_values = data_source[metric][mp]
                for i, val in enumerate(mp_values):
                    if i < max_len and not np.isnan(val):
                        if np.isnan(eu5_values[i]):
                            eu5_values[i] = 0
                        
                        if metric == 'Transit Conversion' or metric == 'UPO':
                            # For rates, we'll average them
                            eu5_values[i] += val
                            valid_counts[i] += 1
                        else:
                            # For counts, we sum them
                            eu5_values[i] += val
                            valid_counts[i] = 1  # Mark as valid
            
            # For rates, calculate average
            if metric == 'Transit Conversion' or metric == 'UPO':
                eu5_values = [v / c if c > 0 else np.nan for v, c in zip(eu5_values, valid_counts)]
            
            data_source[metric]['EU5'] = eu5_values
            valid_count = sum(1 for v in eu5_values if not np.isnan(v))
            source_name = "forecast" if is_forecast else "actuals"
            print(f"  EU5 ({metric}) [{source_name}]: {valid_count} valid data points (calculated)")


def test_processor():
    """Test the data processor with a sample file"""
    processor = DataProcessor()
    success, message = processor.load_excel('inputs_forecasting.xlsx')
    print(f"\nLoad result: {success}, {message}")
    
    if success:
        print("\n=== Actuals Data Summary ===")
        all_data = processor.get_all_data()
        for metric in all_data:
            print(f"\n{metric}:")
            for mp in all_data[metric]:
                data = all_data[metric][mp]
                print(f"  {mp}: {len(data['values'])} data points")
                if data['values']:
                    print(f"    First: {data['dates'][0]} = {data['values'][0]}")
                    print(f"    Last: {data['dates'][-1]} = {data['values'][-1]}")
        
        if processor.has_manual_forecast:
            print("\n=== Manual Forecast Summary ===")
            forecast_data = processor.get_manual_forecast_data()
            for metric in forecast_data:
                print(f"\n{metric}:")
                for mp in forecast_data[metric]:
                    data = forecast_data[metric][mp]
                    print(f"  {mp}: {len(data['values'])} data points")
                    if data['values']:
                        print(f"    First: {data['dates'][0]} = {data['values'][0]}")
                        print(f"    Last: {data['dates'][-1]} = {data['values'][-1]}")
            
            print("\n=== Forecast Accuracy ===")
            accuracy = processor.get_all_accuracy_metrics()
            for metric in accuracy:
                print(f"\n{metric}:")
                for mp in accuracy[metric]:
                    acc = accuracy[metric][mp]
                    print(f"  {mp}: WMAPE={acc['wmape']}%, Accuracy={acc['accuracy']}%, Bias={acc['bias']}%, Overlap={acc['overlap_weeks']} weeks")


if __name__ == '__main__':
    test_processor()
