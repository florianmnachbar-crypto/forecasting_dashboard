"""
Data Processor for Amazon Haul EU5 Forecasting Dashboard
Handles Excel file parsing and data transformation
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re


class DataProcessor:
    """Processes Excel input files for the forecasting dashboard"""
    
    METRICS = ['Net Ordered Units', 'Transits', 'Transit Conversion', 'UPO']
    MARKETPLACES = ['UK', 'DE', 'FR', 'IT', 'ES', 'EU5']
    
    def __init__(self):
        self.data = {}
        self.weeks = []
        self.raw_df = None
        
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
            # Read the Excel file without headers
            df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
            self.raw_df = df
            
            print(f"Excel loaded: {df.shape[0]} rows, {df.shape[1]} columns")
            
            # Find all metric sections
            self.data = {}
            
            for metric in self.METRICS:
                metric_data = self._parse_metric_section(df, metric)
                if metric_data:
                    self.data[metric] = metric_data
                    print(f"Parsed {metric}: {list(metric_data.keys())}")
            
            if not self.data:
                return False, "No data sections found in the Excel file"
            
            # Recalculate EU5 totals if needed
            self.calculate_eu5_totals()
            
            return True, "File loaded successfully"
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"Error loading file: {str(e)}"
    
    def _parse_metric_section(self, df, metric_name):
        """Parse a single metric section from the dataframe"""
        # Find the metric header row
        metric_row, metric_col = self.find_cell_value(df, metric_name)
        
        if metric_row is None:
            print(f"Could not find metric: {metric_name}")
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
        
        print(f"Found MP header at row {mp_row}, col {header_col}")
        
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
        
        # Store weeks for this metric (use the first one found)
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
            elif mp_name in self.METRICS:
                # We've reached the next metric section
                break
        
        return metric_data if metric_data else None
    
    def get_dataframe(self, metric, marketplace):
        """Get a pandas DataFrame for a specific metric and marketplace"""
        if metric not in self.data or marketplace not in self.data[metric]:
            return None
        
        values = self.data[metric][marketplace]
        
        # Use only as many weeks as we have values
        week_count = min(len(values), len(self.weeks))
        dates = [self.parse_week_column(self.weeks[i]) for i in range(week_count)]
        
        df = pd.DataFrame({
            'ds': dates[:len(values)],
            'y': values[:len(dates)],
            'week': self.weeks[:len(values)]
        })
        
        # Remove rows with NaN values for cleaner forecasting
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
    
    def calculate_eu5_totals(self):
        """Recalculate EU5 totals from individual marketplace data"""
        individual_mps = ['UK', 'DE', 'FR', 'IT', 'ES']
        
        for metric in self.METRICS:
            if metric not in self.data:
                continue
            
            # Get the length from existing data
            max_len = 0
            for mp in individual_mps:
                if mp in self.data[metric]:
                    max_len = max(max_len, len(self.data[metric][mp]))
            
            if max_len == 0:
                continue
            
            # Initialize EU5 with NaN
            eu5_values = [np.nan] * max_len
            valid_counts = [0] * max_len
            
            for mp in individual_mps:
                if mp not in self.data[metric]:
                    continue
                    
                mp_values = self.data[metric][mp]
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
            
            self.data[metric]['EU5'] = eu5_values
            valid_count = sum(1 for v in eu5_values if not np.isnan(v))
            print(f"  EU5 ({metric}): {valid_count} valid data points (calculated)")


def test_processor():
    """Test the data processor with a sample file"""
    processor = DataProcessor()
    success, message = processor.load_excel('inputs_forecasting.xlsx')
    print(f"\nLoad result: {success}, {message}")
    
    if success:
        print("\n=== Data Summary ===")
        all_data = processor.get_all_data()
        for metric in all_data:
            print(f"\n{metric}:")
            for mp in all_data[metric]:
                data = all_data[metric][mp]
                print(f"  {mp}: {len(data['values'])} data points")
                if data['values']:
                    print(f"    First: {data['dates'][0]} = {data['values'][0]}")
                    print(f"    Last: {data['dates'][-1]} = {data['values'][-1]}")


if __name__ == '__main__':
    test_processor()
