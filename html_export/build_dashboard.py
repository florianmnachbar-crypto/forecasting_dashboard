"""
Build Dashboard Script - Full 1:1 Replica
Generates a self-contained HTML dashboard matching the localhost version exactly.

Usage (from html_export/ folder):
    python build_dashboard.py --input ../inputs_forecasting.xlsx --output dashboard_report.html

Usage (from root folder):
    python html_export/build_dashboard.py --input inputs_forecasting.xlsx --output html_export/dashboard_report.html
"""

import os
import sys
import json
import argparse
import webbrowser
from datetime import datetime

# Add parent directory to path for imports when running from html_export/
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from data_processor import DataProcessor
from forecaster import Forecaster

BUILD_VERSION = "2.0.0"


def generate_all_forecasts(data_processor, forecast_horizon=12):
    """Generate forecasts for all metrics and marketplaces"""
    forecasts = {}
    forecaster = Forecaster(forecast_horizon=forecast_horizon)
    
    driver_metrics = ['Transits', 'Transit Conversion', 'UPO']
    MAX_TRANSIT_CONVERSION = 0.10
    UPO_CAP_MULTIPLIER = 2.0
    TRANSITS_CAP_MULTIPLIER = 3.0
    
    def get_historical_max(metric, marketplace):
        try:
            df = data_processor.get_dataframe(metric, marketplace)
            if df is not None and not df.empty:
                return df['y'].max()
        except Exception:
            pass
        return None
    
    eu5_transits_max = get_historical_max('Transits', 'EU5')
    
    for metric in driver_metrics:
        forecasts[metric] = {}
        for mp in DataProcessor.MARKETPLACES:
            df = data_processor.get_dataframe(metric, mp)
            
            if df is not None and not df.empty and len(df) >= 4:
                try:
                    forecast = forecaster.forecast_sarimax(df, use_seasonality=True)
                    
                    if forecast:
                        if metric == 'Transit Conversion':
                            forecast['values'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['values']]
                            forecast['lower_bound'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['lower_bound']]
                            forecast['upper_bound'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['upper_bound']]
                        elif metric == 'Transits':
                            mp_max = get_historical_max('Transits', mp)
                            if mp_max and eu5_transits_max:
                                cap = min(eu5_transits_max, mp_max * TRANSITS_CAP_MULTIPLIER)
                                forecast['values'] = [min(v, cap) for v in forecast['values']]
                                forecast['lower_bound'] = [min(v, cap) for v in forecast['lower_bound']]
                                forecast['upper_bound'] = [min(v, cap) for v in forecast['upper_bound']]
                        elif metric == 'UPO':
                            mp_max = get_historical_max('UPO', mp)
                            if mp_max:
                                cap = mp_max * UPO_CAP_MULTIPLIER
                                forecast['values'] = [min(v, cap) for v in forecast['values']]
                                forecast['lower_bound'] = [min(v, cap) for v in forecast['lower_bound']]
                                forecast['upper_bound'] = [min(v, cap) for v in forecast['upper_bound']]
                        
                        forecasts[metric][mp] = forecast
                except Exception as e:
                    print(f"  Warning: Could not forecast {metric} for {mp}: {e}")
    
    # Calculate Net Ordered Units
    forecasts['Net Ordered Units'] = {}
    for mp in DataProcessor.MARKETPLACES:
        if mp in forecasts.get('Transits', {}) and mp in forecasts.get('Transit Conversion', {}) and mp in forecasts.get('UPO', {}):
            t_fc = forecasts['Transits'][mp]
            c_fc = forecasts['Transit Conversion'][mp]
            u_fc = forecasts['UPO'][mp]
            
            nou_values = [max(0, t_fc['values'][i] * c_fc['values'][i] * u_fc['values'][i]) for i in range(len(t_fc['values']))]
            nou_lower = [max(0, t_fc['lower_bound'][i] * c_fc['lower_bound'][i] * u_fc['lower_bound'][i]) for i in range(len(t_fc['values']))]
            nou_upper = [max(0, t_fc['upper_bound'][i] * c_fc['upper_bound'][i] * u_fc['upper_bound'][i]) for i in range(len(t_fc['values']))]
            
            forecasts['Net Ordered Units'][mp] = {
                'dates': t_fc['dates'],
                'values': nou_values,
                'lower_bound': nou_lower,
                'upper_bound': nou_upper,
                'model': 'Calculated (T×C×U)'
            }
    
    return forecasts


def generate_statistics(data_processor):
    """Generate summary statistics with T4W and CW values"""
    stats = {}
    for metric in DataProcessor.METRICS:
        stats[metric] = {}
        for mp in DataProcessor.MARKETPLACES:
            stat = data_processor.get_summary_statistics(metric, mp)
            if stat:
                # Get the dataframe for additional calculations
                df = data_processor.get_dataframe(metric, mp)
                if df is not None and not df.empty:
                    values = df['y'].dropna()
                    
                    # T4W (Trailing 4 Weeks) statistics
                    t4w_values = values.tail(4)
                    stat['t4w_total'] = round(float(t4w_values.sum()), 2) if len(t4w_values) > 0 else 0
                    stat['t4w_avg'] = round(float(t4w_values.mean()), 2) if len(t4w_values) > 0 else 0
                    stat['t4w_min'] = round(float(t4w_values.min()), 2) if len(t4w_values) > 0 else 0
                    stat['t4w_max'] = round(float(t4w_values.max()), 2) if len(t4w_values) > 0 else 0
                    
                    # CW (Current Week) - latest value
                    stat['cw_value'] = round(float(values.iloc[-1]), 2) if len(values) > 0 else 0
                
                stats[metric][mp] = stat
    return stats


def generate_accuracy_metrics(data_processor):
    """Generate accuracy metrics for all timeframes (total, t4w, cw)"""
    if not data_processor.has_manual_forecast:
        return None
    try:
        return {
            'total': data_processor.get_all_accuracy_metrics(timeframe='total'),
            't4w': data_processor.get_all_accuracy_metrics(timeframe='t4w'),
            'cw': data_processor.get_all_accuracy_metrics(timeframe='cw')
        }
    except Exception:
        return None


def read_css_file():
    """Read the original CSS file from parent directory"""
    # Try current directory first, then parent directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    css_path = os.path.join(script_dir, 'static', 'css', 'style.css')
    if os.path.exists(css_path):
        with open(css_path, 'r', encoding='utf-8') as f:
            return f.read()
    # Try parent directory (when running from html_export/)
    parent_dir = os.path.dirname(script_dir)
    css_path = os.path.join(parent_dir, 'static', 'css', 'style.css')
    if os.path.exists(css_path):
        with open(css_path, 'r', encoding='utf-8') as f:
            return f.read()
    return ""


def build_html(data, forecasts, statistics, accuracy, latest_week, promo_scores, has_manual_forecast, generated_at, input_file):
    """Build the complete HTML dashboard"""
    
    data_json = json.dumps(data, default=str)
    forecasts_json = json.dumps(forecasts, default=str)
    statistics_json = json.dumps(statistics, default=str)
    accuracy_json = json.dumps(accuracy, default=str) if accuracy else 'null'
    latest_week_json = json.dumps(latest_week, default=str) if latest_week else 'null'
    promo_json = json.dumps(promo_scores, default=str) if promo_scores else 'null'
    
    css = read_css_file()
    
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Amazon Haul EU5 Forecasting Dashboard - Static Report</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
{css}
    </style>
</head>
<body>
    <div class="bg-animation"></div>
    
    <header class="header">
        <div class="header-content">
            <div class="logo">
                <div class="logo-icon">📊</div>
                <div>
                    <h1>Amazon Haul EU5</h1>
                    <span>Static Report | Generated: {generated_at}</span>
                </div>
            </div>
            <div class="header-actions">
                <div class="theme-toggle" id="themeToggle" title="Toggle Light/Dark Mode">
                    <i class="fas fa-moon active" id="darkIcon"></i>
                    <div class="theme-switch"></div>
                    <i class="fas fa-sun" id="lightIcon"></i>
                </div>
                <div class="file-status loaded">
                    <i class="fas fa-file-excel"></i>
                    <span>{input_file}</span>
                </div>
            </div>
        </div>
    </header>

    <div class="main-container">
        <aside class="sidebar" id="sidebar">
            <div class="sidebar-section">
                <h3><i class="fas fa-sliders-h"></i> Controls</h3>
                <div class="form-group">
                    <label>Metric</label>
                    <select id="metricSelect">
                        <option value="Net Ordered Units">Net Ordered Units</option>
                        <option value="Transits">Transits</option>
                        <option value="Transit Conversion">Transit Conversion</option>
                        <option value="UPO">Units Per Order (UPO)</option>
                    </select>
                </div>
                <div class="toggle-group" id="manualForecastToggleGroup" style="display: {'block' if has_manual_forecast else 'none'};">
                    <span class="toggle-label">Manual Forecast</span>
                    <label class="toggle-switch">
                        <input type="checkbox" id="manualForecastToggle" checked>
                        <span class="toggle-slider"></span>
                    </label>
                </div>
            </div>
            
            <div class="sidebar-section">
                <h3><i class="fas fa-globe-europe"></i> Marketplaces</h3>
                <div class="mp-list">
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="EU5" checked><div class="mp-flag eu5">EU5</div><span>EU5 (All)</span></label>
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="UK" checked><div class="mp-flag uk">UK</div><span>United Kingdom</span></label>
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="DE" checked><div class="mp-flag de">DE</div><span>Germany</span></label>
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="FR" checked><div class="mp-flag fr">FR</div><span>France</span></label>
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="IT" checked><div class="mp-flag it">IT</div><span>Italy</span></label>
                    <label class="mp-item"><input type="checkbox" class="mp-checkbox" value="ES" checked><div class="mp-flag es">ES</div><span>Spain</span></label>
                </div>
            </div>
            
            <div class="sidebar-section">
                <h3><i class="fas fa-chart-bar"></i> Statistics View</h3>
                <div class="stats-toggle">
                    <button class="stats-toggle-btn active" data-view="total">Total</button>
                    <button class="stats-toggle-btn" data-view="t4w">T4W</button>
                    <button class="stats-toggle-btn" data-view="cw">CW</button>
                </div>
            </div>
            
            <div class="model-info show">
                <h4>Report Information</h4>
                <p>Model: SARIMAX</p>
                <p>Generated: {generated_at}</p>
                <p>Confidence: 85%</p>
            </div>
        </aside>

        <main class="main-content" id="mainContent">
            <section class="dashboard-section" id="dashboardSection">
                <div class="tab-navigation">
                    <button class="tab-btn active" data-tab="forecasts"><i class="fas fa-chart-line"></i> Forecasts</button>
                    <button class="tab-btn" data-tab="latest-week"><i class="fas fa-calendar-week"></i> Latest Week Overview</button>
                    <button class="tab-btn" data-tab="historic-deviations"><i class="fas fa-history"></i> Historic Deviations</button>
                </div>

                <div class="tab-content active" id="tab-forecasts">
                    <div class="section-header">
                        <h2 class="section-title"><i class="fas fa-analytics"></i> Statistics</h2>
                        <span class="current-metric" id="currentMetricLabel">Net Ordered Units</span>
                    </div>
                    <div class="stats-grid" id="statsGrid"></div>

                    <div class="section-header" style="margin-top: 1rem;">
                        <h2 class="section-title"><i class="fas fa-chart-line"></i> Forecasts by Marketplace</h2>
                    </div>
                    <div class="charts-grid" id="chartsGrid"></div>
                </div>

                <div class="tab-content" id="tab-latest-week">
                    <div class="section-header">
                        <h2 class="section-title"><i class="fas fa-calendar-week"></i> Latest Week Overview</h2>
                        <span class="latest-week-label" id="latestWeekLabel">--</span>
                    </div>
                    <div class="latest-week-container">
                        <table class="latest-week-table" id="latestWeekTable">
                            <thead>
                                <tr>
                                    <th>Marketplace</th>
                                    <th colspan="3">Net Ordered Units</th>
                                    <th colspan="3">Transits</th>
                                    <th colspan="3">Transit Conversion</th>
                                    <th colspan="3">UPO</th>
                                </tr>
                                <tr class="sub-header">
                                    <th></th>
                                    <th>Actual</th><th>Forecast</th><th>Dev %</th>
                                    <th>Actual</th><th>Forecast</th><th>Dev %</th>
                                    <th>Actual</th><th>Forecast</th><th>Dev %</th>
                                    <th>Actual</th><th>Forecast</th><th>Dev %</th>
                                </tr>
                            </thead>
                            <tbody id="latestWeekTableBody"></tbody>
                        </table>
                    </div>
                    <div class="deviation-legend">
                        <span class="legend-item"><span class="legend-color green"></span> &lt;20% deviation</span>
                        <span class="legend-item"><span class="legend-color yellow"></span> 20-30% deviation</span>
                        <span class="legend-item"><span class="legend-color red"></span> &gt;30% deviation</span>
                    </div>
                </div>

                <div class="tab-content" id="tab-historic-deviations">
                    <div class="section-header">
                        <h2 class="section-title"><i class="fas fa-history"></i> Historic Deviations</h2>
                        <div class="deviation-controls">
                            <select id="deviationMetricSelect" class="deviation-metric-select">
                                <option value="Net Ordered Units">Net Ordered Units</option>
                                <option value="Transits">Transits</option>
                                <option value="Transit Conversion">Transit Conversion</option>
                                <option value="UPO">UPO</option>
                            </select>
                            <select id="deviationMpSelect" class="deviation-mp-select">
                                <option value="EU5">EU5 (All)</option>
                                <option value="UK">UK</option>
                                <option value="DE">DE</option>
                                <option value="FR">FR</option>
                                <option value="IT">IT</option>
                                <option value="ES">ES</option>
                            </select>
                        </div>
                    </div>
                    <div class="deviation-chart-container"><div id="deviationChart" class="deviation-chart" style="height:400px;"></div></div>
                    <div class="historic-deviations-container">
                        <table class="historic-deviations-table" id="historicDeviationsTable">
                            <thead><tr><th>Week</th><th>Actual</th><th>Manual FC</th><th>Manual Dev %</th><th>Model FC</th><th>Model Dev %</th></tr></thead>
                            <tbody id="historicDeviationsTableBody"></tbody>
                        </table>
                    </div>
                    <div class="deviation-summary" id="deviationSummary"></div>
                </div>
            </section>
        </main>
    </div>

    <footer class="footer">Amazon Haul EU5 Forecasting Dashboard | Static Report v{BUILD_VERSION} | 85% Confidence Interval</footer>

    <script>
        // Embedded Data
        const dashboardData = {data_json};
        const forecastsData = {forecasts_json};
        const statisticsData = {statistics_json};
        const accuracyData = {accuracy_json};
        const latestWeekData = {latest_week_json};
        const promoData = {promo_json};
        const hasManualForecast = {'true' if has_manual_forecast else 'false'};
        
        const METRICS = ['Net Ordered Units', 'Transits', 'Transit Conversion', 'UPO'];
        const MARKETPLACES = ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'];
        const MP_COLORS = {{'EU5':'#667eea','UK':'#ff9900','DE':'#00d9ff','FR':'#ff6b9d','IT':'#00e676','ES':'#ffeb3b'}};
        
        let currentMetric = 'Net Ordered Units';
        let currentStatsView = 'total';
        let showManualForecast = true;
        let selectedMarketplaces = ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'];
        
        // Helper function to resize all Plotly charts after DOM updates
        function resizeAllCharts() {{
            setTimeout(function() {{
                document.querySelectorAll('[id^="chart-"]').forEach(function(el) {{
                    if(el && el.data) {{
                        Plotly.Plots.resize(el);
                    }}
                }});
                // Also resize deviation chart
                const devChart = document.getElementById('deviationChart');
                if(devChart && devChart.data) {{
                    Plotly.Plots.resize(devChart);
                }}
            }}, 100);
        }}
        
        // Theme Toggle
        document.getElementById('themeToggle').addEventListener('click', function() {{
            const html = document.documentElement;
            const current = html.getAttribute('data-theme');
            html.setAttribute('data-theme', current === 'light' ? '' : 'light');
            document.getElementById('darkIcon').classList.toggle('active');
            document.getElementById('lightIcon').classList.toggle('active');
        }});
        
        // Tab Navigation
        document.querySelectorAll('.tab-btn').forEach(btn => {{
            btn.addEventListener('click', function() {{
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                this.classList.add('active');
                document.getElementById('tab-' + this.dataset.tab).classList.add('active');
                // Resize charts when switching tabs
                resizeAllCharts();
            }});
        }});
        
        // Metric Select
        document.getElementById('metricSelect').addEventListener('change', function() {{
            currentMetric = this.value;
            document.getElementById('currentMetricLabel').textContent = currentMetric;
            updateDashboard();
            resizeAllCharts();
        }});
        
        // Manual Forecast Toggle
        document.getElementById('manualForecastToggle').addEventListener('change', function() {{
            showManualForecast = this.checked;
            updateCharts();
            resizeAllCharts();
        }});
        
        // Marketplace Checkboxes
        document.querySelectorAll('.mp-checkbox').forEach(cb => {{
            cb.addEventListener('change', function() {{
                selectedMarketplaces = Array.from(document.querySelectorAll('.mp-checkbox:checked')).map(c => c.value);
                updateDashboard();
                resizeAllCharts();
            }});
        }});
        
        // Stats Toggle - also update accuracy and charts
        document.querySelectorAll('.stats-toggle-btn').forEach(btn => {{
            btn.addEventListener('click', function() {{
                document.querySelectorAll('.stats-toggle-btn').forEach(b => b.classList.remove('active'));
                this.classList.add('active');
                currentStatsView = this.dataset.view;
                updateStats();
                updateAccuracy();
                updateCharts();
                resizeAllCharts();
            }});
        }});
        
        // Deviation Selects
        document.getElementById('deviationMetricSelect').addEventListener('change', function() {{
            updateDeviations();
            resizeAllCharts();
        }});
        document.getElementById('deviationMpSelect').addEventListener('change', function() {{
            updateDeviations();
            resizeAllCharts();
        }});
        
        function formatValue(v, m) {{
            if(v==null||v==undefined||isNaN(v)) return '-';
            if(m==='Transit Conversion') return (v*100).toFixed(2)+'%';
            if(m==='UPO') return v.toFixed(2);
            return Math.round(v).toLocaleString();
        }}
        
        function getDevClass(dev) {{
            if(dev===null||dev===undefined) return '';
            const abs = Math.abs(dev);
            if(abs<20) return 'dev-green';
            if(abs<30) return 'dev-yellow';
            return 'dev-red';
        }}
        
        function updateStats() {{
            const grid = document.getElementById('statsGrid');
            grid.innerHTML = '';
            
            const viewLabels = {{ total: 'Total', t4w: 'T4W', cw: 'CW' }};
            const currentViewLabel = viewLabels[currentStatsView] || 'Total';
            
            selectedMarketplaces.forEach(mp => {{
                const stats = statisticsData[currentMetric] && statisticsData[currentMetric][mp];
                if(!stats) return;
                
                let primaryValue, avgValue, minValue, maxValue;
                if(currentStatsView === 'total') {{
                    primaryValue = stats.total;
                    avgValue = stats.average;
                    minValue = stats.min;
                    maxValue = stats.max;
                }} else if(currentStatsView === 't4w') {{
                    primaryValue = stats.t4w_total;
                    avgValue = stats.t4w_avg;
                    minValue = stats.t4w_min;
                    maxValue = stats.t4w_max;
                }} else {{
                    primaryValue = stats.cw_value;
                    avgValue = stats.cw_value;
                    minValue = stats.cw_value;
                    maxValue = stats.cw_value;
                }}
                
                const card = document.createElement('div');
                card.className = 'stat-card';
                card.innerHTML = `
                    <div class="stat-card-header">
                        <h4>${{mp}}</h4>
                        <div class="mp-flag ${{mp.toLowerCase()}}">${{mp}}</div>
                    </div>
                    <div class="stat-card-body">
                        <div class="stat-item"><div class="value">${{formatValue(primaryValue, currentMetric)}}</div><div class="label">${{currentViewLabel}} Total</div></div>
                        <div class="stat-item"><div class="value">${{formatValue(avgValue, currentMetric)}}</div><div class="label">${{currentViewLabel}} Avg</div></div>
                        <div class="stat-item"><div class="value">${{formatValue(minValue, currentMetric)}}</div><div class="label">Min</div></div>
                        <div class="stat-item"><div class="value">${{formatValue(maxValue, currentMetric)}}</div><div class="label">Max</div></div>
                    </div>
                `;
                grid.appendChild(card);
            }});
        }}
        
        function updateAccuracy() {{
            // Accuracy panel was removed - WMAPE now shown on chart cards
            // This function is kept for compatibility but does nothing
        }}
        
        function updateCharts() {{
            const grid = document.getElementById('chartsGrid');
            grid.innerHTML = '';
            
            selectedMarketplaces.forEach(mp => {{
                const metricData = dashboardData[currentMetric];
                if(!metricData || !metricData[mp]) return;
                
                const mpData = metricData[mp];
                const forecast = forecastsData[currentMetric] && forecastsData[currentMetric][mp];
                
                const card = document.createElement('div');
                card.className = 'chart-card';
                const chartId = 'chart-' + mp;
                card.innerHTML = `
                    <div class="chart-card-header">
                        <h4><div class="chart-icon" style="background:${{MP_COLORS[mp]}}">${{mp}}</div>${{mp}} - ${{currentMetric}}</h4>
                    </div>
                    <div class="chart-container"><div id="${{chartId}}" style="height:350px;"></div></div>
                `;
                grid.appendChild(card);
                
                // Get actual container width for proper sizing
                const containerEl = card.querySelector('.chart-container');
                const containerWidth = containerEl ? containerEl.offsetWidth : 300;
                
                // Build traces - filter to only show weeks with actual data
                const traces = [];
                
                // Historical data - filter out null/NaN values
                const allWeeks = mpData.weeks || [];
                const allValues = mpData.values || [];
                const validIndices = [];
                allValues.forEach((v, i) => {{
                    if(v != null && !isNaN(v)) validIndices.push(i);
                }});
                const weeks = validIndices.map(i => allWeeks[i]);
                const values = validIndices.map(i => allValues[i]);
                
                // Set to track the valid week range (from actuals) - used to constrain other traces
                const validWeekSet = new Set(weeks);
                const firstActualWeek = weeks[0];
                const lastActualWeek = weeks[weeks.length - 1];
                
                traces.push({{
                    x: weeks,
                    y: values,
                    name: 'Actual',
                    type: 'scatter',
                    mode: 'lines+markers',
                    line: {{ color: MP_COLORS[mp], width: 2 }},
                    marker: {{ size: 6 }}
                }});
                
                // Manual forecast - only show weeks that are within or after the actuals range
                if(hasManualForecast && showManualForecast && mpData.manual_forecast) {{
                    const mfWeeks = mpData.manual_weeks || [];
                    const mfValues = mpData.manual_forecast || [];
                    
                    // Filter to valid values AND weeks that overlap with or extend from actuals
                    const filteredMfWeeks = [];
                    const filteredMfValues = [];
                    mfWeeks.forEach((wk, i) => {{
                        const val = mfValues[i];
                        if(val != null && !isNaN(val)) {{
                            // Only include if week is in actuals range OR after actuals end
                            if(validWeekSet.has(wk) || weeks.indexOf(wk) === -1) {{
                                // Check if this week comes after the first actual week
                                const mfWeekIdx = allWeeks.indexOf(wk);
                                const firstActualIdx = allWeeks.indexOf(firstActualWeek);
                                if(mfWeekIdx >= firstActualIdx) {{
                                    filteredMfWeeks.push(wk);
                                    filteredMfValues.push(val);
                                }}
                            }}
                        }}
                    }});
                    
                    if(filteredMfWeeks.length > 0) {{
                        traces.push({{
                            x: filteredMfWeeks,
                            y: filteredMfValues,
                            name: 'Manual FC',
                            type: 'scatter',
                            mode: 'lines+markers',
                            line: {{ color: '#9c27b0', width: 2, dash: 'dot' }},
                            marker: {{ size: 5 }}
                        }});
                    }}
                }}
                
                // Model forecast
                if(forecast) {{
                    const fcWeeks = forecast.dates.map(d => {{
                        const date = new Date(d);
                        const y = date.getFullYear();
                        const start = new Date(y,0,1);
                        const days = Math.floor((date-start)/(24*60*60*1000));
                        const w = Math.ceil((days+start.getDay()+1)/7);
                        return 'Wk'+w.toString().padStart(2,'0')+' '+y;
                    }});
                    
                    traces.push({{
                        x: fcWeeks,
                        y: forecast.values,
                        name: 'Model FC',
                        type: 'scatter',
                        mode: 'lines+markers',
                        line: {{ color: '#ff9900', width: 2, dash: 'dash' }},
                        marker: {{ size: 5 }}
                    }});
                    
                    // Confidence interval
                    traces.push({{
                        x: fcWeeks.concat(fcWeeks.slice().reverse()),
                        y: forecast.upper_bound.concat(forecast.lower_bound.slice().reverse()),
                        fill: 'toself',
                        fillcolor: 'rgba(255,153,0,0.1)',
                        line: {{ color: 'transparent' }},
                        name: '85% CI',
                        showlegend: false,
                        hoverinfo: 'skip'
                    }});
                }}
                
                // Calculate Y-axis range based on historical + manual forecast ONLY
                const scaleValues = [...values];
                if(hasManualForecast && mpData.manual_forecast) {{
                    scaleValues.push(...mpData.manual_forecast.filter(v => v != null && !isNaN(v)));
                }}
                const yMax = scaleValues.length > 0 ? Math.max(...scaleValues) * 1.15 : 100;
                
                // Shorten week labels for better fit: "Wk01 2025" -> "W01"
                const shortWeeks = weeks.map((w, i) => {{
                    // Simple extraction: "Wk01 2025" -> "W01"
                    const parts = w.split(' ');
                    if(parts.length >= 1) {{
                        const wkPart = parts[0].replace('Wk', 'W');
                        // Show year at first, last, and every 10th week
                        if(parts.length >= 2 && (i === 0 || i === weeks.length - 1 || i % 10 === 0)) {{
                            return wkPart + "'" + parts[1].slice(-2);
                        }}
                        return wkPart;
                    }}
                    return w;
                }});
                
                // Build x-axis category list: actuals weeks + future weeks (from Manual FC and Model FC)
                const xAxisCategories = [...weeks];
                
                // Add Manual FC weeks that extend beyond actuals (future forecasts)
                if(hasManualForecast && mpData.manual_weeks) {{
                    const mfWeeks = mpData.manual_weeks || [];
                    const mfValues = mpData.manual_forecast || [];
                    mfWeeks.forEach((wk, i) => {{
                        const val = mfValues[i];
                        if(val != null && !isNaN(val) && !xAxisCategories.includes(wk)) {{
                            // Only add if week is AFTER the last actual week (future forecast)
                            const wkIdx = mpData.weeks.indexOf(wk);
                            const lastActualIdx = mpData.weeks.indexOf(lastActualWeek);
                            if(wkIdx > lastActualIdx || wkIdx === -1) {{
                                xAxisCategories.push(wk);
                            }}
                        }}
                    }});
                }}
                
                // Add Model FC weeks that extend beyond actuals
                if(forecast) {{
                    const fcWeeks = forecast.dates.map(d => {{
                        const date = new Date(d);
                        const y = date.getFullYear();
                        const start = new Date(y,0,1);
                        const days = Math.floor((date-start)/(24*60*60*1000));
                        const w = Math.ceil((days+start.getDay()+1)/7);
                        return 'Wk'+w.toString().padStart(2,'0')+' '+y;
                    }});
                    fcWeeks.forEach(fw => {{
                        if(!xAxisCategories.includes(fw)) xAxisCategories.push(fw);
                    }});
                }}
                
                // Calculate tick skip based on FULL data length including forecasts
                const fullDataLen = xAxisCategories.length;
                let tickSkip = 1;
                if(fullDataLen > 40) tickSkip = 5;
                else if(fullDataLen > 30) tickSkip = 4;
                else if(fullDataLen > 15) tickSkip = 2;
                
                // Select ticks to show (first, every Nth, last) - for full range including forecasts
                const tickVals = [];
                const tickText = [];
                xAxisCategories.forEach((w, i) => {{
                    if(i === 0 || i === xAxisCategories.length - 1 || i % tickSkip === 0) {{
                        tickVals.push(w);
                        tickText.push(w);
                    }}
                }});
                
                // Debug: log xAxisCategories for first marketplace
                if(mp === 'EU5') console.log('EU5 xAxisCategories:', xAxisCategories.length, 'weeks from', xAxisCategories[0], 'to', xAxisCategories[xAxisCategories.length-1]);
                
                const layout = {{
                    margin: {{ t: 10, r: 20, b: 70, l: 60 }},
                    paper_bgcolor: 'transparent',
                    plot_bgcolor: 'transparent',
                    font: {{ color: '#8892b0', family: 'Inter', size: 10 }},
                    xaxis: {{ 
                        gridcolor: 'rgba(255,255,255,0.1)', 
                        tickangle: -45, 
                        tickfont: {{ size: 9 }},
                        type: 'category',
                        categoryorder: 'array',
                        categoryarray: xAxisCategories,
                        range: [-0.5, xAxisCategories.length - 0.5]
                    }},
                    yaxis: {{ 
                        gridcolor: 'rgba(255,255,255,0.1)', 
                        tickfont: {{ size: 9 }},
                        tickformat: currentMetric === 'Transit Conversion' ? '.2%' : '.2s',
                        automargin: true,
                        range: [0, yMax],
                        autorange: false
                    }},
                    legend: {{ orientation: 'h', y: -0.25, x: 0.5, xanchor: 'center', font: {{ size: 10 }} }},
                    showlegend: true,
                    hovermode: 'x unified',
                    dragmode: 'pan'
                }};
                
                const config = {{
                    responsive: true,
                    displayModeBar: true,
                    modeBarButtonsToInclude: ['zoom2d', 'pan2d', 'resetScale2d', 'zoomIn2d', 'zoomOut2d'],
                    displaylogo: false
                }};
                
                Plotly.newPlot(chartId, traces, layout, config).then(function() {{
                    // Trigger autoscale after initial render to fit all data
                    Plotly.relayout(chartId, {{'xaxis.autorange': true}});
                }});
                
                // Add forecast stats below chart (including accuracy)
                const statsDiv = document.createElement('div');
                statsDiv.className = 'forecast-stats';
                
                let statsHtml = '';
                
                // Model FC stats
                if(forecast) {{
                    let fcVals = forecast.values;
                    let viewLabel = 'Total';
                    if(currentStatsView === 'cw') {{ fcVals = fcVals.slice(0,1); viewLabel = 'CW'; }}
                    else if(currentStatsView === 't4w') {{ fcVals = fcVals.slice(0,4); viewLabel = 'T4W'; }}
                    const fcTotal = fcVals.reduce((a,b) => a+b, 0);
                    const fcAvg = fcTotal / fcVals.length;
                    const modelName = forecast.model || 'SARIMAX';
                    statsHtml += '<div class="forecast-stat"><div class="value">' + formatValue(fcTotal, currentMetric) + '</div><div class="label">Model FC ' + viewLabel + '</div></div>' +
                        '<div class="forecast-stat"><div class="value">' + formatValue(fcAvg, currentMetric) + '</div><div class="label">' + viewLabel + ' Avg</div></div>' +
                        '<div class="forecast-stat"><div class="value">' + modelName + '</div><div class="label">Model</div></div>';
                }}
                
                // Add Manual FC Accuracy (WMAPE only) to the card - uses timeframe-based data
                if(hasManualForecast && accuracyData && accuracyData[currentStatsView] && accuracyData[currentStatsView][currentMetric] && accuracyData[currentStatsView][currentMetric][mp]) {{
                    const acc = accuracyData[currentStatsView][currentMetric][mp];
                    if(acc && acc.wmape !== null) {{
                        const wmape = acc.wmape;
                        const accCls = wmape < 20 ? 'good' : (wmape < 30 ? 'medium' : 'poor');
                        statsHtml += '<div class="forecast-stat"><div class="value accuracy-' + accCls + '">' + wmape.toFixed(1) + '%</div><div class="label">WMAPE</div></div>';
                    }}
                }}
                
                statsDiv.innerHTML = statsHtml;
                card.appendChild(statsDiv);
            }});
        }}
        
        function updateLatestWeek() {{
            if(!latestWeekData || !latestWeekData.data) return;
            document.getElementById('latestWeekLabel').textContent = latestWeekData.latest_week || '--';
            const tbody = document.getElementById('latestWeekTableBody');
            tbody.innerHTML = '';
            MARKETPLACES.forEach(mp => {{
                const mpData = latestWeekData.data[mp];
                if(!mpData) return;
                let row = '<tr><td class="mp-cell"><div class="mp-flag ' + mp.toLowerCase() + '">' + mp + '</div></td>';
                METRICS.forEach(m => {{
                    const d = mpData[m] || {{}};
                    const actual = formatValue(d.actual, m);
                    const fc = d.manual_forecast !== null ? formatValue(d.manual_forecast, m) : '-';
                    const dev = d.manual_dev_pct;
                    const devCls = getDevClass(dev);
                    const devStr = dev !== null ? (dev > 0 ? '+' : '') + dev.toFixed(1) + '%' : '-';
                    row += '<td class="value-cell">' + actual + '</td><td class="forecast-cell">' + fc + '</td><td class="deviation-cell ' + devCls + '">' + devStr + '</td>';
                }});
                row += '</tr>';
                tbody.innerHTML += row;
            }});
        }}
        
        function updateDeviations() {{
            const metric = document.getElementById('deviationMetricSelect').value;
            const mp = document.getElementById('deviationMpSelect').value;
            const mData = dashboardData[metric] && dashboardData[metric][mp];
            if(!mData) return;
            
            const allWeeks = mData.weeks || [];
            const allActuals = mData.values || [];
            const manualFc = mData.manual_forecast || [];
            const manualWeeks = mData.manual_weeks || [];
            
            // Filter to only valid actuals
            const validIndices = [];
            allActuals.forEach((v, i) => {{
                if(v != null && !isNaN(v)) validIndices.push(i);
            }});
            const weeks = validIndices.map(i => allWeeks[i]);
            const actuals = validIndices.map(i => allActuals[i]);
            
            // Create a map of week to manual forecast
            const manualFcMap = {{}};
            if(manualWeeks.length > 0 && manualFc.length > 0) {{
                manualWeeks.forEach((wk, idx) => {{
                    if(idx < manualFc.length && manualFc[idx] != null && !isNaN(manualFc[idx])) {{
                        manualFcMap[wk] = manualFc[idx];
                    }}
                }});
            }}
            
            // Build aligned data for chart - only weeks where both actual and FC exist
            const chartWeeks = [];
            const chartActuals = [];
            const chartForecasts = [];
            const chartDeviations = [];
            
            weeks.forEach((wk, i) => {{
                const actual = actuals[i];
                const fc = manualFcMap[wk];
                if(actual != null && fc != null && fc !== 0) {{
                    chartWeeks.push(wk);
                    chartActuals.push(actual);
                    chartForecasts.push(fc);
                    const dev = ((actual - fc) / fc) * 100;
                    chartDeviations.push(dev);
                }}
            }});
            
            // Create line chart with deviation bars
            const traces = [
                {{
                    x: chartWeeks,
                    y: chartActuals,
                    name: 'Actual',
                    type: 'scatter',
                    mode: 'lines+markers',
                    line: {{ color: MP_COLORS[mp], width: 2 }},
                    marker: {{ size: 8 }},
                    yaxis: 'y'
                }},
                {{
                    x: chartWeeks,
                    y: chartForecasts,
                    name: 'Manual FC',
                    type: 'scatter',
                    mode: 'lines+markers',
                    line: {{ color: '#9c27b0', width: 2, dash: 'dot' }},
                    marker: {{ size: 6 }},
                    yaxis: 'y'
                }},
                {{
                    x: chartWeeks,
                    y: chartDeviations,
                    name: 'Deviation %',
                    type: 'bar',
                    marker: {{ 
                        color: chartDeviations.map(d => d >= 0 ? 'rgba(76,175,80,0.6)' : 'rgba(244,67,54,0.6)'),
                        line: {{ color: chartDeviations.map(d => d >= 0 ? '#4caf50' : '#f44336'), width: 1 }}
                    }},
                    yaxis: 'y2'
                }}
            ];
            
            // Calculate Y-axis range
            const allVals = [...chartActuals, ...chartForecasts];
            const yMax = allVals.length > 0 ? Math.max(...allVals) * 1.1 : 100;
            const devMax = chartDeviations.length > 0 ? Math.max(Math.abs(Math.min(...chartDeviations)), Math.abs(Math.max(...chartDeviations)), 30) * 1.2 : 50;
            
            Plotly.newPlot('deviationChart', traces, {{
                margin: {{ t: 20, r: 60, b: 80, l: 60 }},
                paper_bgcolor: 'transparent', 
                plot_bgcolor: 'transparent',
                font: {{ color: '#8892b0', family: 'Inter' }},
                xaxis: {{ 
                    gridcolor: 'rgba(255,255,255,0.1)', 
                    tickangle: -45,
                    type: 'category'
                }},
                yaxis: {{ 
                    gridcolor: 'rgba(255,255,255,0.1)',
                    title: {{ text: metric, font: {{ size: 10 }} }},
                    range: [0, yMax]
                }},
                yaxis2: {{
                    overlaying: 'y',
                    side: 'right',
                    title: {{ text: 'Deviation %', font: {{ size: 10 }} }},
                    range: [-devMax, devMax],
                    zeroline: true,
                    zerolinecolor: 'rgba(255,255,255,0.3)',
                    gridcolor: 'rgba(255,255,255,0.05)'
                }},
                legend: {{ orientation: 'h', y: -0.3, font: {{ size: 10 }} }},
                barmode: 'overlay'
            }}, {{ responsive: true, displayModeBar: false }});
            
            // Build deviation table - show most recent weeks first (only compared weeks)
            const tbody = document.getElementById('historicDeviationsTableBody');
            tbody.innerHTML = '';
            const reversedIndices = [...Array(chartWeeks.length).keys()].reverse();
            reversedIndices.forEach(idx => {{
                const wk = chartWeeks[idx];
                const actual = chartActuals[idx];
                const fc = chartForecasts[idx];
                const dev = chartDeviations[idx];
                const devCls = getDevClass(dev);
                tbody.innerHTML += '<tr><td class="week-cell">' + wk + '</td><td class="value-cell">' + formatValue(actual, metric) + '</td><td class="forecast-cell">' + formatValue(fc, metric) + '</td><td class="deviation-cell ' + devCls + '">' + (dev > 0 ? '+' : '') + dev.toFixed(1) + '%</td><td>-</td><td>-</td></tr>';
            }});
            
            // Add deviation summary
            const summary = document.getElementById('deviationSummary');
            if(summary && chartDeviations.length > 0) {{
                const absDeviations = chartDeviations.map(d => Math.abs(d));
                const avgDev = absDeviations.reduce((a,b) => a+b, 0) / absDeviations.length;
                const maxDev = Math.max(...absDeviations);
                const minDev = Math.min(...absDeviations);
                const devClass = avgDev < 20 ? 'summary-good' : (avgDev < 30 ? 'summary-warn' : 'summary-bad');
                summary.innerHTML = '<div class="deviation-summary-grid"><div class="summary-card"><div class="summary-value">' + chartWeeks.length + '</div><div class="summary-label">Compared Weeks</div></div><div class="summary-card ' + devClass + '"><div class="summary-value">' + avgDev.toFixed(1) + '%</div><div class="summary-label">Avg |Dev|</div></div><div class="summary-card"><div class="summary-value">' + maxDev.toFixed(1) + '%</div><div class="summary-label">Max |Dev|</div></div><div class="summary-card"><div class="summary-value">' + minDev.toFixed(1) + '%</div><div class="summary-label">Min |Dev|</div></div></div>';
            }} else {{
                summary.innerHTML = '<p style="color:var(--text-muted);padding:1rem;">No overlapping data available for comparison</p>';
            }}
        }}
        
        function updateDashboard() {{
            updateStats();
            updateAccuracy();
            updateCharts();
            updateLatestWeek();
            updateDeviations();
        }}
        
        document.addEventListener('DOMContentLoaded', function() {{
            updateDashboard();
            // Fix Plotly timing issue - resize charts after DOM is fully rendered
            resizeAllCharts();
        }});
    </script>
</body>
</html>'''
    
    return html


def main():
    parser = argparse.ArgumentParser(description='Build static HTML dashboard')
    
    # Determine default input path based on script location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    
    # Default: look for inputs_forecasting.xlsx in parent directory first, then current directory
    default_input = os.path.join(parent_dir, 'inputs_forecasting.xlsx')
    if not os.path.exists(default_input):
        default_input = 'inputs_forecasting.xlsx'
    
    parser.add_argument('--input', '-i', default=default_input, help='Input Excel file')
    parser.add_argument('--output', '-o', default='dashboard_report.html', help='Output HTML file')
    parser.add_argument('--no-open', action='store_true', help='Do not open browser')
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("  Amazon Haul EU5 Dashboard Builder v" + BUILD_VERSION)
    print("="*60)
    
    if not os.path.exists(args.input):
        print(f"\n  ERROR: Input file not found: {args.input}")
        sys.exit(1)
    
    print(f"\n  Input:  {args.input}")
    print(f"  Output: {args.output}")
    
    print("\n  [1/5] Loading Excel data...")
    data_processor = DataProcessor()
    success, message = data_processor.load_excel(args.input)
    if not success:
        print(f"  ERROR: {message}")
        sys.exit(1)
    print(f"        {message}")
    
    print("  [2/5] Calculating EU5 totals...")
    data_processor.calculate_eu5_totals()
    
    print("  [3/5] Extracting data...")
    data = data_processor.get_all_data()
    
    # Merge manual forecast into actuals data for frontend
    if data_processor.has_manual_forecast:
        manual_fc_data = data_processor.get_manual_forecast_data()
        if manual_fc_data:
            for metric in manual_fc_data:
                if metric in data:
                    for mp in manual_fc_data[metric]:
                        if mp in data[metric]:
                            mf = manual_fc_data[metric][mp]
                            data[metric][mp]['manual_forecast'] = mf.get('values', [])
                            data[metric][mp]['manual_weeks'] = mf.get('weeks', [])
                            data[metric][mp]['manual_dates'] = mf.get('dates', [])
            print(f"        Merged manual forecast data for {len(manual_fc_data)} metrics")
    
    statistics = generate_statistics(data_processor)
    accuracy = generate_accuracy_metrics(data_processor)
    latest_week = data_processor.get_latest_week_overview()
    promo_scores = getattr(data_processor, 'promo_scores', None)
    
    print("  [4/5] Generating SARIMAX forecasts...")
    forecasts = generate_all_forecasts(data_processor)
    fc_count = sum(len(mp_fc) for mp_fc in forecasts.values())
    print(f"        Generated {fc_count} forecasts")
    
    print("  [5/5] Building HTML dashboard...")
    generated_at = datetime.now().strftime('%Y-%m-%d %H:%M')
    html = build_html(data, forecasts, statistics, accuracy, latest_week, promo_scores, 
                      data_processor.has_manual_forecast, generated_at, args.input)
    
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(html)
    
    file_size = os.path.getsize(args.output) / 1024
    print(f"\n  SUCCESS! Dashboard saved to: {args.output}")
    print(f"  File size: {file_size:.1f} KB")
    
    if not args.no_open:
        print("\n  Opening in browser...")
        webbrowser.open('file://' + os.path.realpath(args.output))
    
    print("\n" + "="*60)
    print("  Build complete!")
    print("="*60 + "\n")


if __name__ == '__main__':
    main()
