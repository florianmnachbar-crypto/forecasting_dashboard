"""
Forecasting Module for Amazon Haul EU5 Dashboard
Implements SARIMAX and Prophet models with configurable seasonality
"""

import pandas as pd
import numpy as np
from datetime import timedelta
import warnings
warnings.filterwarnings('ignore')


class Forecaster:
    """Handles forecasting using SARIMAX or Prophet models"""
    
    def __init__(self, forecast_horizon=12):
        self.forecast_horizon = forecast_horizon
        self.model = None
        self.fitted = False
        
    def prepare_data(self, df):
        """Prepare data for forecasting"""
        if df is None or df.empty:
            return None
        
        # Ensure we have the required columns
        if 'ds' not in df.columns or 'y' not in df.columns:
            return None
        
        # Sort by date and reset index
        df = df.sort_values('ds').reset_index(drop=True)
        
        # Remove any NaN values
        df = df.dropna(subset=['y'])
        
        return df
    
    def forecast_sarimax(self, df, use_seasonality=True, exog=None, future_exog=None):
        """
        Generate forecast using SARIMAX model
        
        Parameters:
        - df: DataFrame with 'ds' (dates) and 'y' (values) columns
        - use_seasonality: If True, uses seasonal component (SARIMAX), else ARIMAX
        - exog: DataFrame with exogenous variables for historical data (e.g., promo_score)
        - future_exog: DataFrame with exogenous variables for forecast period
        
        Returns:
        - Dictionary with forecast dates, values, and confidence intervals
        """
        from statsmodels.tsa.statespace.sarimax import SARIMAX
        
        df = self.prepare_data(df)
        if df is None or len(df) < 4:
            return None
        
        try:
            # Set frequency to weekly
            y = df.set_index('ds')['y']
            y.index = pd.DatetimeIndex(y.index).to_period('W').to_timestamp()
            
            # Handle duplicates by taking mean
            y = y.groupby(y.index).mean()
            y = y.asfreq('W-MON', method='ffill')
            
            if len(y) < 4:
                return None
            
            # Process exogenous variables if provided
            exog_aligned = None
            has_exog = False
            
            if exog is not None and len(exog) > 0:
                # Align exog with y index
                try:
                    exog_df = exog.copy()
                    if 'ds' in exog_df.columns:
                        exog_df = exog_df.set_index('ds')
                        exog_df.index = pd.DatetimeIndex(exog_df.index).to_period('W').to_timestamp()
                    
                    # Get only columns that are numeric (exclude date columns)
                    numeric_cols = exog_df.select_dtypes(include=[np.number]).columns.tolist()
                    if numeric_cols:
                        exog_aligned = exog_df[numeric_cols].reindex(y.index)
                        # Fill NaN with column mean
                        exog_aligned = exog_aligned.fillna(exog_aligned.mean())
                        # Drop any rows that are still NaN
                        valid_mask = ~exog_aligned.isna().any(axis=1) & ~y.isna()
                        y = y[valid_mask]
                        exog_aligned = exog_aligned[valid_mask]
                        has_exog = len(exog_aligned) == len(y) and len(numeric_cols) > 0
                except Exception as e:
                    print(f"Exog alignment warning: {e}")
                    has_exog = False
            
            # Define SARIMAX parameters
            if use_seasonality and len(y) >= 8:
                # Seasonal ARIMA with weekly seasonality (period=4 for monthly pattern)
                order = (1, 1, 1)
                seasonal_order = (1, 0, 1, 4)  # Seasonal component
            else:
                # Non-seasonal ARIMA
                order = (1, 1, 1)
                seasonal_order = (0, 0, 0, 0)
            
            # Fit the model
            model = SARIMAX(
                y,
                exog=exog_aligned if has_exog else None,
                order=order,
                seasonal_order=seasonal_order,
                enforce_stationarity=False,
                enforce_invertibility=False
            )
            
            fitted_model = model.fit(disp=False, maxiter=200)
            
            # Generate forecast
            # Prepare future exogenous variables if needed
            forecast_exog = None
            if has_exog and future_exog is not None:
                try:
                    if 'ds' in future_exog.columns:
                        future_exog_df = future_exog.drop(columns=['ds'], errors='ignore')
                    else:
                        future_exog_df = future_exog.copy()
                    
                    # Ensure same columns as training exog
                    numeric_cols = exog_aligned.columns.tolist()
                    forecast_exog = future_exog_df[numeric_cols].iloc[:self.forecast_horizon]
                    
                    # Fill NaN with mean from training
                    for col in numeric_cols:
                        if forecast_exog[col].isna().any():
                            forecast_exog[col] = forecast_exog[col].fillna(exog_aligned[col].mean())
                except Exception as e:
                    print(f"Future exog warning: {e}")
                    forecast_exog = None
            
            forecast_result = fitted_model.get_forecast(
                steps=self.forecast_horizon,
                exog=forecast_exog if has_exog and forecast_exog is not None else None
            )
            forecast_mean = forecast_result.predicted_mean
            conf_int = forecast_result.conf_int(alpha=0.15)  # 85% confidence interval
            
            # Generate future dates
            last_date = y.index[-1]
            future_dates = [last_date + timedelta(weeks=i+1) for i in range(self.forecast_horizon)]
            
            # Ensure non-negative values for count metrics
            forecast_values = forecast_mean.values
            lower_bound = conf_int.iloc[:, 0].values
            upper_bound = conf_int.iloc[:, 1].values
            
            return {
                'dates': [d.strftime('%Y-%m-%d') for d in future_dates],
                'values': [max(0, float(v)) for v in forecast_values],
                'lower_bound': [max(0, float(v)) for v in lower_bound],
                'upper_bound': [max(0, float(v)) for v in upper_bound],
                'model': 'SARIMAX' if use_seasonality else 'ARIMAX',
                'model_info': {
                    'order': order,
                    'seasonal_order': seasonal_order if use_seasonality else None,
                    'aic': round(fitted_model.aic, 2) if hasattr(fitted_model, 'aic') else None,
                    'has_exog': has_exog,
                    'exog_columns': list(exog_aligned.columns) if has_exog else None
                }
            }
            
        except Exception as e:
            print(f"SARIMAX error: {str(e)}")
            return self._fallback_forecast(df)
    
    def forecast_prophet(self, df, use_seasonality=True):
        """
        Generate forecast using Facebook Prophet model
        
        Parameters:
        - df: DataFrame with 'ds' (dates) and 'y' (values) columns
        - use_seasonality: If True, enables weekly/yearly seasonality
        
        Returns:
        - Dictionary with forecast dates, values, and confidence intervals
        """
        try:
            from prophet import Prophet
        except ImportError:
            print("Prophet not installed. Using fallback forecast.")
            return self._fallback_forecast(df)
        
        df = self.prepare_data(df)
        if df is None or len(df) < 2:
            return None
        
        try:
            # Prepare data for Prophet (requires 'ds' and 'y' columns)
            prophet_df = df[['ds', 'y']].copy()
            prophet_df['ds'] = pd.to_datetime(prophet_df['ds'])
            
            # Initialize Prophet model
            model = Prophet(
                yearly_seasonality=use_seasonality,
                weekly_seasonality=False,  # Data is already weekly
                daily_seasonality=False,
                interval_width=0.95,
                changepoint_prior_scale=0.05
            )
            
            # Add custom seasonality if enabled
            if use_seasonality and len(prophet_df) >= 8:
                model.add_seasonality(
                    name='monthly',
                    period=30.5,
                    fourier_order=3
                )
            
            # Fit the model
            model.fit(prophet_df)
            
            # Create future dataframe
            future = model.make_future_dataframe(periods=self.forecast_horizon, freq='W')
            
            # Generate forecast
            forecast = model.predict(future)
            
            # Extract only the forecast period (not the historical fit)
            forecast_period = forecast.tail(self.forecast_horizon)
            
            return {
                'dates': forecast_period['ds'].dt.strftime('%Y-%m-%d').tolist(),
                'values': [max(0, float(v)) for v in forecast_period['yhat'].values],
                'lower_bound': [max(0, float(v)) for v in forecast_period['yhat_lower'].values],
                'upper_bound': [max(0, float(v)) for v in forecast_period['yhat_upper'].values],
                'model': 'Prophet',
                'model_info': {
                    'seasonality': use_seasonality,
                    'changepoints': len(model.changepoints) if hasattr(model, 'changepoints') else 0
                }
            }
            
        except Exception as e:
            print(f"Prophet error: {str(e)}")
            return self._fallback_forecast(df)
    
    def _fallback_forecast(self, df):
        """Simple fallback forecast using moving average when models fail"""
        df = self.prepare_data(df)
        if df is None or df.empty:
            return None
        
        try:
            # Use simple moving average
            recent_values = df['y'].tail(4)
            avg_value = recent_values.mean()
            std_value = recent_values.std() if len(recent_values) > 1 else avg_value * 0.1
            
            last_date = df['ds'].max()
            future_dates = [last_date + timedelta(weeks=i+1) for i in range(self.forecast_horizon)]
            
            # Add slight trend based on recent data
            if len(df) >= 2:
                trend = (df['y'].iloc[-1] - df['y'].iloc[-2]) / df['y'].iloc[-2] if df['y'].iloc[-2] != 0 else 0
                trend = max(-0.1, min(0.1, trend))  # Limit trend
            else:
                trend = 0
            
            forecast_values = []
            for i in range(self.forecast_horizon):
                val = avg_value * (1 + trend * (i + 1) * 0.5)
                forecast_values.append(max(0, val))
            
            return {
                'dates': [d.strftime('%Y-%m-%d') for d in future_dates],
                'values': forecast_values,
                'lower_bound': [max(0, v - 2 * std_value) for v in forecast_values],
                'upper_bound': [v + 2 * std_value for v in forecast_values],
                'model': 'Moving Average (Fallback)',
                'model_info': {
                    'method': 'simple_moving_average',
                    'window': 4
                }
            }
        except Exception as e:
            print(f"Fallback forecast error: {str(e)}")
            return None
    
    def generate_forecast(self, df, model_type='sarimax', use_seasonality=True):
        """
        Main method to generate forecast
        
        Parameters:
        - df: DataFrame with historical data
        - model_type: 'sarimax' or 'prophet'
        - use_seasonality: Enable/disable seasonality component
        
        Returns:
        - Forecast dictionary or None if failed
        """
        if model_type.lower() == 'prophet':
            return self.forecast_prophet(df, use_seasonality)
        else:
            return self.forecast_sarimax(df, use_seasonality)


def test_forecaster():
    """Test the forecaster with sample data"""
    # Create sample data
    dates = pd.date_range(start='2025-01-01', periods=20, freq='W')
    values = [100 + i * 5 + np.random.normal(0, 10) for i in range(20)]
    
    df = pd.DataFrame({
        'ds': dates,
        'y': values
    })
    
    forecaster = Forecaster(forecast_horizon=12)
    
    # Test SARIMAX
    print("Testing SARIMAX with seasonality:")
    result = forecaster.generate_forecast(df, model_type='sarimax', use_seasonality=True)
    if result:
        print(f"  Model: {result['model']}")
        print(f"  Forecast dates: {result['dates'][:3]}...")
        print(f"  Forecast values: {[round(v, 2) for v in result['values'][:3]]}...")
    
    print("\nTesting SARIMAX without seasonality:")
    result = forecaster.generate_forecast(df, model_type='sarimax', use_seasonality=False)
    if result:
        print(f"  Model: {result['model']}")
        print(f"  Forecast values: {[round(v, 2) for v in result['values'][:3]]}...")
    
    # Test Prophet
    print("\nTesting Prophet:")
    result = forecaster.generate_forecast(df, model_type='prophet', use_seasonality=True)
    if result:
        print(f"  Model: {result['model']}")
        print(f"  Forecast values: {[round(v, 2) for v in result['values'][:3]]}...")


if __name__ == '__main__':
    test_forecaster()
