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
    
    # Promo score bands for analysis
    PROMO_BANDS = [
        (0, 1, 'No/Low Promo'),
        (1.1, 2, 'Light Promo'),
        (2.1, 3, 'Medium Promo'),
        (3.1, 5, 'Strong Promo')
    ]
    
    def __init__(self):
        self.data = {}  # Actuals data
        self.manual_forecast = {}  # Manual forecast data
        self.promo_scores = {}  # Promo scores by marketplace and week
        self.promo_descriptions = {}  # Campaign descriptions
        self.weeks = []
        self.forecast_weeks = []
        self.raw_df = None
        self.has_manual_forecast = False
        self.has_promo_scores = False
        
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
            
            # Load promo scores if available
            if 'Promo Scores' in sheet_names:
                self.load_promo_scores(file_path)
            
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
                
                # Recalculate Net Ordered Units from components to fix [object Object] issues
                self._recalculate_net_ordered_units(is_forecast=True)
                
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
    
    def calculate_forecast_accuracy(self, metric, marketplace, timeframe='total'):
        """Calculate accuracy metrics for manual forecast vs actuals
        
        Args:
            metric: The metric name
            marketplace: The marketplace code
            timeframe: 'total' (all overlap), 't4w' (last 4 weeks), or 'cw' (current week only)
        
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
        
        # Sort by date to ensure correct ordering for timeframe filtering
        merged = merged.sort_values('ds')
        
        # Apply timeframe filter
        if timeframe == 'cw':
            # Current week only - take the most recent overlapping week
            merged = merged.tail(1)
        elif timeframe == 't4w':
            # Trailing 4 weeks - take the last 4 overlapping weeks
            merged = merged.tail(4)
        # else: 'total' - use all overlapping data
        
        if merged.empty:
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
            'periods': [d.strftime('%Y-%m-%d') for d in valid_data['ds']],
            'timeframe': timeframe
        }
    
    def get_all_accuracy_metrics(self, timeframe='total'):
        """Get forecast accuracy for all metrics and marketplaces
        
        Args:
            timeframe: 'total' (all overlap), 't4w' (last 4 weeks), or 'cw' (current week only)
        """
        if not self.has_manual_forecast:
            return None
        
        result = {}
        
        for metric in self.METRICS:
            result[metric] = {}
            for mp in self.MARKETPLACES:
                accuracy = self.calculate_forecast_accuracy(metric, mp, timeframe=timeframe)
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
    
    def _recalculate_net_ordered_units(self, is_forecast=False):
        """Recalculate Net Ordered Units from component metrics (Transits × CVR × UPO)
        
        This fixes issues where Excel formulas produce [object Object] instead of values
        """
        data_source = self.manual_forecast if is_forecast else self.data
        source_name = "forecast" if is_forecast else "actuals"
        
        # Check if we have all required component metrics
        required_metrics = ['Transits', 'Transit Conversion', 'UPO']
        for metric in required_metrics:
            if metric not in data_source:
                print(f"  Cannot recalculate Net Ordered Units: missing {metric}")
                return
        
        # Initialize Net Ordered Units if not present
        if 'Net Ordered Units' not in data_source:
            data_source['Net Ordered Units'] = {}
        
        # Get the list of marketplaces from component metrics
        all_mps = set()
        for metric in required_metrics:
            all_mps.update(data_source[metric].keys())
        
        individual_mps = ['UK', 'DE', 'FR', 'IT', 'ES']
        
        for mp in individual_mps:
            if mp not in all_mps:
                continue
            
            # Check if all components exist for this marketplace
            has_all = all(mp in data_source[m] for m in required_metrics)
            if not has_all:
                continue
            
            transits = data_source['Transits'][mp]
            cvr = data_source['Transit Conversion'][mp]
            upo = data_source['UPO'][mp]
            
            # Determine the length
            max_len = max(len(transits), len(cvr), len(upo))
            
            # Calculate Net Ordered Units for each week
            calculated_values = []
            recalculated_count = 0
            
            for i in range(max_len):
                # Get component values (with bounds checking)
                t = transits[i] if i < len(transits) else np.nan
                c = cvr[i] if i < len(cvr) else np.nan
                u = upo[i] if i < len(upo) else np.nan
                
                # Get existing NOU value if available
                existing_nou = np.nan
                if mp in data_source.get('Net Ordered Units', {}) and i < len(data_source['Net Ordered Units'].get(mp, [])):
                    existing_nou = data_source['Net Ordered Units'][mp][i]
                
                # If existing value is valid, keep it; otherwise calculate
                if not np.isnan(existing_nou):
                    calculated_values.append(existing_nou)
                elif not np.isnan(t) and not np.isnan(c) and not np.isnan(u):
                    # Calculate: Net Ordered Units = Transits × CVR × UPO
                    nou = t * c * u
                    calculated_values.append(nou)
                    recalculated_count += 1
                else:
                    calculated_values.append(np.nan)
            
            # Store the calculated values
            data_source['Net Ordered Units'][mp] = calculated_values
            
            valid_count = sum(1 for v in calculated_values if not np.isnan(v))
            print(f"  {mp} Net Ordered Units [{source_name}]: {valid_count} valid ({recalculated_count} recalculated from components)")
    
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
    
    def load_promo_scores(self, file_path):
        """Load promo scores from the 'Promo Scores' sheet"""
        try:
            xl = pd.ExcelFile(file_path)
            if 'Promo Scores' not in xl.sheet_names:
                print("No 'Promo Scores' sheet found")
                return False
            
            df = pd.read_excel(file_path, sheet_name='Promo Scores', header=None)
            print(f"Promo Scores sheet: {df.shape[0]} rows, {df.shape[1]} columns")
            
            self.promo_scores = {}
            self.promo_descriptions = {}
            
            # Find the header row (row containing "MP")
            header_row = None
            mp_col = None
            
            for row_idx in range(min(5, len(df))):
                for col_idx in range(min(5, len(df.columns))):
                    cell = df.iloc[row_idx, col_idx]
                    if pd.notna(cell) and str(cell).strip() == 'MP':
                        header_row = row_idx
                        mp_col = col_idx
                        break
                if header_row is not None:
                    break
            
            if header_row is None:
                print("Could not find 'MP' header in Promo Scores sheet")
                return False
            
            print(f"Found MP header at row {header_row}, col {mp_col}")
            
            # Parse week headers from the header row (columns after MP)
            promo_weeks = []
            week_col_map = {}  # Maps column index to normalized week label
            
            for col_idx in range(mp_col + 1, len(df.columns)):
                header = df.iloc[header_row, col_idx]
                if pd.isna(header):
                    continue
                
                header_str = str(header).strip()
                normalized_week = self._normalize_promo_week(header_str)
                if normalized_week:
                    promo_weeks.append(normalized_week)
                    week_col_map[col_idx] = normalized_week
            
            print(f"Found {len(promo_weeks)} week columns in promo scores")
            
            if not week_col_map:
                print("No valid week columns found")
                return False
            
            # Parse marketplace rows (rows after header)
            individual_mps = ['UK', 'DE', 'FR', 'IT', 'ES', 'EU5']
            
            for row_idx in range(header_row + 1, min(header_row + 10, len(df))):
                mp_value = df.iloc[row_idx, mp_col]
                if pd.isna(mp_value):
                    continue
                
                mp_name = str(mp_value).strip()
                
                if mp_name in individual_mps:
                    self.promo_scores[mp_name] = {}
                    
                    for col_idx, week_label in week_col_map.items():
                        val = df.iloc[row_idx, col_idx]
                        if pd.notna(val):
                            try:
                                score = float(val)
                                self.promo_scores[mp_name][week_label] = score
                            except (ValueError, TypeError):
                                # Not a number, might be description row
                                pass
                    
                    score_count = len(self.promo_scores[mp_name])
                    print(f"  {mp_name}: {score_count} promo scores")
                elif mp_name == 'WK':
                    # This marks the start of description rows
                    break
            
            # Find and parse description rows (look for second set of MP rows after "WK" marker)
            wk_row = None
            for row_idx in range(header_row + 1, min(header_row + 15, len(df))):
                mp_value = df.iloc[row_idx, mp_col]
                if pd.notna(mp_value) and str(mp_value).strip() == 'WK':
                    wk_row = row_idx
                    break
            
            if wk_row is not None:
                for row_idx in range(wk_row + 1, min(wk_row + 10, len(df))):
                    mp_value = df.iloc[row_idx, mp_col]
                    if pd.isna(mp_value):
                        continue
                    
                    mp_name = str(mp_value).strip()
                    
                    if mp_name in individual_mps:
                        self.promo_descriptions[mp_name] = {}
                        
                        for col_idx, week_label in week_col_map.items():
                            val = df.iloc[row_idx, col_idx]
                            if pd.notna(val):
                                desc = str(val).strip()
                                if desc and desc != '0' and not desc.replace('.', '').replace('-', '').isdigit():
                                    self.promo_descriptions[mp_name][week_label] = desc
            
            if self.promo_scores:
                self.has_promo_scores = True
                # Recalculate EU5 promo scores to only include weeks with full coverage
                self._calculate_eu5_promo_scores()
                print("Promo scores loaded successfully")
                return True
            
            return False
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Warning: Could not load promo scores: {str(e)}")
            return False
    
    def _normalize_promo_week(self, week_str, default_year=2025):
        """Normalize week string from promo scores to standard format (Wk## YYYY)
        
        Handles various formats:
        - "Wk19" or "Wk 19" -> "Wk19 2025" (uses default_year)
        - "2026 wk1" or "2026 wk 1" -> "Wk01 2026"
        - "Wk19 2025" or "Wk 1 2026" -> "Wk19 2025" / "Wk01 2026"
        """
        if not week_str:
            return None
        
        week_str = str(week_str).strip()
        
        # Pattern 1: "Wk19" or "Wk 19" (no year - use default_year)
        match = re.match(r'^Wk\s*(\d+)$', week_str, re.IGNORECASE)
        if match:
            week_num = int(match.group(1))
            return f"Wk{week_num:02d} {default_year}"
        
        # Pattern 2: "2026 wk1" or "2026 wk 1" or "2026wk1" (year prefix)
        match = re.match(r'^(\d{4})\s*wk\s*(\d+)$', week_str, re.IGNORECASE)
        if match:
            year = int(match.group(1))
            week_num = int(match.group(2))
            return f"Wk{week_num:02d} {year}"
        
        # Pattern 3: "Wk19 2025" or "Wk 1 2026" (already has year)
        match = re.match(r'^Wk\s*(\d+)\s+(\d{4})$', week_str, re.IGNORECASE)
        if match:
            week_num = int(match.group(1))
            year = int(match.group(2))
            return f"Wk{week_num:02d} {year}"
        
        return None
    
    def _calculate_eu5_promo_scores(self):
        """Recalculate EU5 promo scores to only include weeks where all 5 MPs have data"""
        individual_mps = ['UK', 'DE', 'FR', 'IT', 'ES']
        
        # Get all weeks that appear in any individual MP promo scores
        all_weeks = set()
        for mp in individual_mps:
            if mp in self.promo_scores:
                all_weeks.update(self.promo_scores[mp].keys())
        
        # For EU5, only include weeks where ALL 5 MPs have promo scores
        eu5_scores = {}
        for week in all_weeks:
            scores = []
            for mp in individual_mps:
                if mp in self.promo_scores and week in self.promo_scores[mp]:
                    scores.append(self.promo_scores[mp][week])
            
            # Only include if all 5 MPs have data for this week
            if len(scores) == 5:
                # Use average of all MP scores for EU5
                eu5_scores[week] = round(sum(scores) / len(scores), 2)
        
        self.promo_scores['EU5'] = eu5_scores
        print(f"  EU5 promo scores recalculated: {len(eu5_scores)} weeks (full coverage only)")
    
    def get_promo_scores_data(self):
        """Get promo scores in a structured format for the frontend"""
        if not self.has_promo_scores:
            return None
        
        result = {
            'scores': self.promo_scores,
            'descriptions': self.promo_descriptions,
            'bands': [
                {'min': b[0], 'max': b[1], 'label': b[2]} 
                for b in self.PROMO_BANDS
            ]
        }
        
        return result
    
    def get_promo_score_for_week(self, marketplace, week_label):
        """Get the promo score for a specific marketplace and week
        
        Handles matching between different week formats:
        - Actuals: "Wk19 2025", "Wk01 2026"
        - Promo scores: "Wk19 2025", "Wk01 2026" (after normalization)
        """
        if not self.has_promo_scores:
            return None
        
        if marketplace not in self.promo_scores:
            return None
        
        mp_scores = self.promo_scores[marketplace]
        
        # Try exact match first
        if week_label in mp_scores:
            return mp_scores[week_label]
        
        # Normalize the input week label and try again
        normalized_input = self._normalize_promo_week(week_label)
        if normalized_input and normalized_input in mp_scores:
            return mp_scores[normalized_input]
        
        # Try matching by extracting week number and year from the input
        # Parse "Wk19 2025" or "Wk 1 2026" format
        match = re.match(r'^Wk\s*(\d+)\s+(\d{4})$', str(week_label).strip(), re.IGNORECASE)
        if match:
            week_num = int(match.group(1))
            year = int(match.group(2))
            
            # Try different normalized formats
            possible_keys = [
                f"Wk{week_num:02d} {year}",
                f"Wk{week_num} {year}",
            ]
            
            for key in possible_keys:
                if key in mp_scores:
                    return mp_scores[key]
        
        return None
    
    def get_promo_band(self, score):
        """Get the promo band label for a given score"""
        if score is None:
            return None
        
        for min_val, max_val, label in self.PROMO_BANDS:
            if min_val <= score <= max_val:
                return label
        
        return 'Unknown'
    
    def calculate_promo_uplift_analysis(self, metric='Net Ordered Units'):
        """Calculate performance uplift coefficients for each promo band
        
        Returns uplift factors relative to baseline (no/low promo weeks)
        """
        if not self.has_promo_scores or not self.data:
            return None
        
        # Collect data points by promo band for each marketplace
        result = {}
        
        for mp in ['UK', 'DE', 'FR', 'IT', 'ES', 'EU5']:
            df = self.get_dataframe(metric, mp, is_forecast=False)
            if df is None or df.empty:
                continue
            
            # Group data by promo band
            band_data = {b[2]: [] for b in self.PROMO_BANDS}
            
            for _, row in df.iterrows():
                week_label = self.format_week_label(row['ds'])
                score = self.get_promo_score_for_week(mp, week_label)
                
                if score is not None:
                    band = self.get_promo_band(score)
                    if band in band_data:
                        band_data[band].append({
                            'week': week_label,
                            'value': row['y'],
                            'score': score
                        })
            
            # Calculate statistics for each band
            band_stats = {}
            baseline_avg = None
            
            for band_label, data_points in band_data.items():
                if not data_points:
                    continue
                
                values = [d['value'] for d in data_points]
                avg = np.mean(values)
                
                band_stats[band_label] = {
                    'count': len(data_points),
                    'total': sum(values),
                    'average': round(avg, 2),
                    'min': round(min(values), 2),
                    'max': round(max(values), 2),
                    'weeks': [d['week'] for d in data_points]
                }
                
                # Use "No/Low Promo" as baseline
                if band_label == 'No/Low Promo':
                    baseline_avg = avg
            
            # Calculate uplift factors relative to baseline
            if baseline_avg and baseline_avg > 0:
                for band_label in band_stats:
                    uplift = band_stats[band_label]['average'] / baseline_avg
                    band_stats[band_label]['uplift_factor'] = round(uplift, 2)
                    band_stats[band_label]['uplift_pct'] = round((uplift - 1) * 100, 1)
            
            result[mp] = {
                'bands': band_stats,
                'baseline_avg': round(baseline_avg, 2) if baseline_avg else None,
                'total_weeks_analyzed': sum(len(b) for b in band_data.values())
            }
        
        return result
    
    def get_all_promo_analysis(self):
        """Get promo analysis for all metrics"""
        if not self.has_promo_scores:
            return None
        
        result = {}
        
        for metric in self.METRICS:
            analysis = self.calculate_promo_uplift_analysis(metric)
            if analysis:
                result[metric] = analysis
        
        return result
    
    def get_forecast_with_promo_uplift(self, metric, marketplace):
        """Apply promo uplift factors to forecast values for FUTURE weeks only
        
        Uplift is calculated as: baseline_avg × uplift_factor
        (NOT manual_forecast × uplift_factor)
        
        Returns forecast data with uplift applied ONLY for weeks after the last actuals date
        Historic weeks (where actuals exist) keep their original manual forecast values
        """
        if not self.has_manual_forecast or not self.has_promo_scores:
            return None
        
        # Get forecast data
        forecast_df = self.get_dataframe(metric, marketplace, is_forecast=True)
        if forecast_df is None or forecast_df.empty:
            return None
        
        # Get actuals data to find the last actuals date
        actuals_df = self.get_dataframe(metric, marketplace, is_forecast=False)
        last_actuals_date = None
        if actuals_df is not None and not actuals_df.empty:
            last_actuals_date = actuals_df['ds'].max()
        
        # Get promo analysis for uplift factors and baseline
        analysis = self.calculate_promo_uplift_analysis(metric)
        if not analysis or marketplace not in analysis:
            return None
        
        mp_analysis = analysis[marketplace]
        bands = mp_analysis.get('bands', {})
        baseline_avg = mp_analysis.get('baseline_avg')
        
        # Build uplift factor map from bands
        uplift_map = {}
        for band_label, stats in bands.items():
            if 'uplift_factor' in stats:
                uplift_map[band_label] = stats['uplift_factor']
        
        # If no uplift factors or baseline available, return None
        if not uplift_map or baseline_avg is None:
            return None
        
        # Apply uplift to each forecast value
        uplifted_values = []
        uplift_details = []
        
        for _, row in forecast_df.iterrows():
            week_date = row['ds']
            week_label = self.format_week_label(week_date)
            original_value = row['y']
            
            # Only apply uplift to FUTURE weeks (after last actuals date)
            is_future_week = last_actuals_date is None or week_date > last_actuals_date
            
            if is_future_week:
                # Get promo score for this week
                score = self.get_promo_score_for_week(marketplace, week_label)
                
                if score is not None:
                    band = self.get_promo_band(score)
                    uplift_factor = uplift_map.get(band, 1.0)
                    # Use baseline × uplift_factor (NOT manual_forecast × uplift_factor)
                    uplifted_value = baseline_avg * uplift_factor
                    
                    uplifted_values.append(uplifted_value)
                    uplift_details.append({
                        'week': week_label,
                        'original': original_value,
                        'uplifted': round(uplifted_value, 2),
                        'score': score,
                        'band': band,
                        'uplift_factor': uplift_factor,
                        'baseline_used': baseline_avg,
                        'is_future': True
                    })
                else:
                    # No promo score for future week - use baseline (No/Low Promo assumption)
                    uplifted_values.append(baseline_avg)
                    uplift_details.append({
                        'week': week_label,
                        'original': original_value,
                        'uplifted': baseline_avg,
                        'score': None,
                        'band': 'No Data (baseline)',
                        'uplift_factor': 1.0,
                        'baseline_used': baseline_avg,
                        'is_future': True
                    })
            else:
                # Historic week - keep original manual forecast value (no uplift)
                uplifted_values.append(original_value)
                uplift_details.append({
                    'week': week_label,
                    'original': original_value,
                    'uplifted': original_value,
                    'score': None,
                    'band': 'Historic (no uplift)',
                    'uplift_factor': 1.0,
                    'baseline_used': None,
                    'is_future': False
                })
        
        return {
            'dates': forecast_df['ds'].dt.strftime('%Y-%m-%d').tolist(),
            'weeks': [self.format_week_label(d) for d in forecast_df['ds']],
            'original_values': forecast_df['y'].tolist(),
            'uplifted_values': uplifted_values,
            'details': uplift_details,
            'uplift_factors': uplift_map,
            'baseline_avg': baseline_avg,
            'last_actuals_date': last_actuals_date.strftime('%Y-%m-%d') if last_actuals_date else None
        }
    
    def get_all_forecast_with_uplift(self):
        """Get uplifted forecast data for all metrics and marketplaces"""
        if not self.has_manual_forecast or not self.has_promo_scores:
            return None
        
        result = {}
        
        for metric in self.METRICS:
            result[metric] = {}
            for mp in self.MARKETPLACES:
                uplift_data = self.get_forecast_with_promo_uplift(metric, mp)
                if uplift_data:
                    result[metric][mp] = uplift_data
        
        return result


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
