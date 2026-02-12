"""
Build Dashboard Script - Full 1:1 Replica v2.7.0
Generates a self-contained HTML dashboard matching the localhost version exactly.
Includes: Promo Analysis tab, promo overlays, Model FC +Promo, click-to-expand modal.

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
from datetime import datetime, timedelta
import pandas as pd

# Add parent directory to path for imports when running from html_export/
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from data_processor import DataProcessor
from forecaster import Forecaster

BUILD_VERSION = "2.7.0"

MAX_TRANSIT_CONVERSION = 0.10
UPO_CAP_MULTIPLIER = 2.0
TRANSITS_CAP_MULTIPLIER = 3.0


def get_historical_max(data_processor, metric, marketplace):
    try:
        df = data_processor.get_dataframe(metric, marketplace)
        if df is not None and not df.empty:
            return df['y'].max()
    except Exception:
        pass
    return None


def cap_forecast(forecast, metric, mp, data_processor, eu5_transits_max):
    """Apply caps to forecast values"""
    if forecast is None:
        return forecast
    if metric == 'Transit Conversion':
        forecast['values'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['values']]
        forecast['lower_bound'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['lower_bound']]
        forecast['upper_bound'] = [min(v, MAX_TRANSIT_CONVERSION) for v in forecast['upper_bound']]
    elif metric == 'Transits':
        mp_max = get_historical_max(data_processor, 'Transits', mp)
        if mp_max and eu5_transits_max:
            cap = min(eu5_transits_max, mp_max * TRANSITS_CAP_MULTIPLIER)
            forecast['values'] = [min(v, cap) for v in forecast['values']]
            forecast['lower_bound'] = [min(v, cap) for v in forecast['lower_bound']]
            forecast['upper_bound'] = [min(v, cap) for v in forecast['upper_bound']]
    elif metric == 'UPO':
        mp_max = get_historical_max(data_processor, 'UPO', mp)
        if mp_max:
            cap = mp_max * UPO_CAP_MULTIPLIER
            forecast['values'] = [min(v, cap) for v in forecast['values']]
            forecast['lower_bound'] = [min(v, cap) for v in forecast['lower_bound']]
            forecast['upper_bound'] = [min(v, cap) for v in forecast['upper_bound']]
    return forecast


def prepare_promo_exog(data_processor, metric, marketplace, df, forecast_horizon):
    """Prepare exogenous promo score data for SARIMAX model"""
    try:
        promo_scores = []
        for _, row in df.iterrows():
            week_label = data_processor.format_week_label(row['ds'])
            score = data_processor.get_promo_score_for_week(marketplace, week_label)
            promo_scores.append(score if score is not None else 1.0)

        exog = pd.DataFrame({'ds': df['ds'], 'promo_score': promo_scores})

        last_date = df['ds'].max()
        future_dates = [last_date + timedelta(weeks=i+1) for i in range(forecast_horizon)]
        future_scores = []
        for fd in future_dates:
            wl = data_processor.format_week_label(fd)
            score = data_processor.get_promo_score_for_week(marketplace, wl)
            future_scores.append(score if score is not None else 1.0)

        future_exog = pd.DataFrame({'ds': future_dates, 'promo_score': future_scores})

        promo_info = {
            'future_scores': [{'week': data_processor.format_week_label(d), 'score': s}
                              for d, s in zip(future_dates, future_scores)]
        }
        return exog, future_exog, promo_info
    except Exception as e:
        print(f"  Warning: promo exog prep failed: {e}")
        return None, None, None


def apply_promo_floor(promo_fc, base_fc, future_scores):
    """Promo score > 1 cannot decrease forecast below baseline"""
    floored = []
    for i in range(len(promo_fc['values'])):
        pv = promo_fc['values'][i]
        bv = base_fc['values'][i]
        ps = future_scores[i] if i < len(future_scores) else 1.0
        if ps > 1.0:
            floored.append(max(pv, bv))
        elif ps == 1.0:
            floored.append(bv)
        else:
            floored.append(pv)
    result = dict(promo_fc)
    result['values'] = floored
    return result


def generate_all_forecasts(data_processor, include_promo=False, forecast_horizon=12):
    """Generate baseline and optionally promo-adjusted forecasts"""
    forecaster = Forecaster(forecast_horizon=forecast_horizon)
    base_forecasts = {}
    promo_forecasts = {}
    driver_metrics = ['Transits', 'Transit Conversion', 'UPO']
    eu5_transits_max = get_historical_max(data_processor, 'Transits', 'EU5')

    for metric in driver_metrics:
        base_forecasts[metric] = {}
        promo_forecasts[metric] = {}
        for mp in DataProcessor.MARKETPLACES:
            df = data_processor.get_dataframe(metric, mp)
            if df is None or df.empty or len(df) < 4:
                continue
            try:
                # Baseline
                fc_base = forecaster.forecast_sarimax(df, use_seasonality=True)
                if fc_base:
                    fc_base = cap_forecast(fc_base, metric, mp, data_processor, eu5_transits_max)
                    base_forecasts[metric][mp] = fc_base

                # Promo
                if include_promo and data_processor.has_promo_scores:
                    exog, future_exog, promo_info = prepare_promo_exog(
                        data_processor, metric, mp, df, forecast_horizon)
                    if exog is not None:
                        fc_promo = forecaster.forecast_sarimax(df, use_seasonality=True, exog=exog, future_exog=future_exog)
                        if fc_promo and fc_base and promo_info:
                            future_scores = [item['score'] for item in promo_info.get('future_scores', [])]
                            fc_promo = apply_promo_floor(fc_promo, fc_base, future_scores)
                            fc_promo = cap_forecast(fc_promo, metric, mp, data_processor, eu5_transits_max)
                            promo_forecasts[metric][mp] = fc_promo
            except Exception as e:
                print(f"  Warning: Could not forecast {metric} for {mp}: {e}")

    # Derive NOU
    for fc_dict in [base_forecasts, promo_forecasts]:
        fc_dict['Net Ordered Units'] = {}
        for mp in DataProcessor.MARKETPLACES:
            t = fc_dict.get('Transits', {}).get(mp)
            c = fc_dict.get('Transit Conversion', {}).get(mp)
            u = fc_dict.get('UPO', {}).get(mp)
            if t and c and u:
                vals = [max(0, t['values'][i] * c['values'][i] * u['values'][i]) for i in range(len(t['values']))]
                lower = [max(0, t['lower_bound'][i] * c['lower_bound'][i] * u['lower_bound'][i]) for i in range(len(t['values']))]
                upper = [max(0, t['upper_bound'][i] * c['upper_bound'][i] * u['upper_bound'][i]) for i in range(len(t['values']))]
                fc_dict['Net Ordered Units'][mp] = {
                    'dates': t['dates'], 'values': vals,
                    'lower_bound': lower, 'upper_bound': upper,
                    'model': 'Calculated (T×C×U)'
                }

    return base_forecasts, promo_forecasts


def generate_statistics(data_processor):
    stats = {}
    for metric in DataProcessor.METRICS:
        stats[metric] = {}
        for mp in DataProcessor.MARKETPLACES:
            stat = data_processor.get_summary_statistics(metric, mp)
            if stat:
                df = data_processor.get_dataframe(metric, mp)
                if df is not None and not df.empty:
                    values = df['y'].dropna()
                    t4w = values.tail(4)
                    stat['t4w_total'] = round(float(t4w.sum()), 2) if len(t4w) > 0 else 0
                    stat['t4w_avg'] = round(float(t4w.mean()), 2) if len(t4w) > 0 else 0
                    stat['t4w_min'] = round(float(t4w.min()), 2) if len(t4w) > 0 else 0
                    stat['t4w_max'] = round(float(t4w.max()), 2) if len(t4w) > 0 else 0
                    stat['cw_value'] = round(float(values.iloc[-1]), 2) if len(values) > 0 else 0
                stats[metric][mp] = stat
    return stats


def generate_accuracy_metrics(data_processor):
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


def generate_promo_analysis(data_processor):
    """Generate promo uplift analysis data"""
    if not data_processor.has_promo_scores:
        return None
    try:
        if data_processor.promo_format == 'regressors':
            return data_processor.get_all_regressor_analysis()
        else:
            return data_processor.get_all_promo_analysis()
    except Exception as e:
        print(f"  Warning: promo analysis failed: {e}")
        return None


def get_promo_regressors_json(data_processor):
    """Extract promo regressors for chart overlays"""
    if not data_processor.has_promo_scores or data_processor.promo_format != 'regressors':
        return None
    return data_processor.promo_regressors


def read_css_file():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for base in [script_dir, os.path.dirname(script_dir)]:
        css_path = os.path.join(base, 'static', 'css', 'style.css')
        if os.path.exists(css_path):
            with open(css_path, 'r', encoding='utf-8') as f:
                return f.read()
    return ""


def build_html(data, base_forecasts, promo_forecasts, statistics, accuracy,
               latest_week, promo_analysis, promo_regressors, discount_values,
               has_manual_forecast, has_promo_scores, promo_format,
               generated_at, input_file):
    """Build the complete HTML dashboard"""

    data_json = json.dumps(data, default=str)
    base_fc_json = json.dumps(base_forecasts, default=str)
    promo_fc_json = json.dumps(promo_forecasts, default=str)
    stats_json = json.dumps(statistics, default=str)
    acc_json = json.dumps(accuracy, default=str) if accuracy else 'null'
    lw_json = json.dumps(latest_week, default=str) if latest_week else 'null'
    pa_json = json.dumps(promo_analysis, default=str) if promo_analysis else 'null'
    pr_json = json.dumps(promo_regressors, default=str) if promo_regressors else 'null'
    dv_json = json.dumps(discount_values, default=str) if discount_values else '[]'

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
/* Light mode toggle fix - ensure sliders are visible */
[data-theme="light"] .toggle-slider {{
    background-color: #ccc !important;
    border: 1px solid #999 !important;
}}
[data-theme="light"] .toggle-switch input:checked + .toggle-slider {{
    background-color: #ff9900 !important;
    border-color: #ff9900 !important;
}}
[data-theme="light"] .toggle-slider:before {{
    background-color: #fff !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3) !important;
}}
    </style>
</head>
<body data-theme-default="light">
    <div class="bg-animation"></div>

    <header class="header">
        <div class="header-content">
            <div class="logo">
                <div class="logo-icon">📊</div>
                <div>
                    <h1>Amazon Haul EU5</h1>
                    <span>Static Report v{BUILD_VERSION} | Generated: {generated_at}</span>
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
                <div class="toggle-group" id="promoOverlayToggleGroup" style="display: {'block' if has_promo_scores else 'none'};">
                    <span class="toggle-label">Promo Overlay</span>
                    <label class="toggle-switch">
                        <input type="checkbox" id="promoOverlayToggle">
                        <span class="toggle-slider"></span>
                    </label>
                </div>
                <div class="toggle-group" id="promoUpliftToggleGroup" style="display: {'block' if has_promo_scores else 'none'};">
                    <span class="toggle-label">Promo Uplift FC</span>
                    <label class="toggle-switch">
                        <input type="checkbox" id="promoUpliftToggle" checked>
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
                    <button class="tab-btn" data-tab="promo-analysis" id="promoAnalysisTab" style="display: {'inline-flex' if has_promo_scores else 'none'};"><i class="fas fa-tags"></i> Promo Analysis</button>
                </div>

                <!-- Forecasts Tab -->
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

                <!-- Latest Week Tab -->
                <div class="tab-content" id="tab-latest-week">
                    <div class="section-header">
                        <h2 class="section-title"><i class="fas fa-calendar-week"></i> Latest Week Overview</h2>
                        <span class="latest-week-label" id="latestWeekLabel">--</span>
                    </div>
                    <div class="latest-week-container">
                        <table class="latest-week-table" id="latestWeekTable">
                            <thead>
                                <tr><th>Marketplace</th><th colspan="3">Net Ordered Units</th><th colspan="3">Transits</th><th colspan="3">Transit Conversion</th><th colspan="3">UPO</th></tr>
                                <tr class="sub-header"><th></th><th>Actual</th><th>Forecast</th><th>Dev %</th><th>Actual</th><th>Forecast</th><th>Dev %</th><th>Actual</th><th>Forecast</th><th>Dev %</th><th>Actual</th><th>Forecast</th><th>Dev %</th></tr>
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

                <!-- Historic Deviations Tab -->
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

                <!-- Promo Analysis Tab -->
                <div class="tab-content" id="tab-promo-analysis">
                    <div class="section-header">
                        <h2 class="section-title"><i class="fas fa-tags"></i> Promo Uplift Analysis</h2>
                        <div class="promo-controls">
                            <select id="promoMetricSelect" class="deviation-metric-select">
                                <option value="Net Ordered Units">Net Ordered Units</option>
                                <option value="Transits">Transits</option>
                                <option value="Transit Conversion">Transit Conversion</option>
                                <option value="UPO">UPO</option>
                            </select>
                            <select id="promoTypeFilter" class="deviation-metric-select">
                                <option value="all">All Promo Types</option>
                                <option value="HVE">HVE</option>
                                <option value="Dollar Deals">Dollar Deals</option>
                                <option value="Discount %">Discount %</option>
                            </select>
                        </div>
                    </div>
                    <div class="promo-legend" id="promoLegend">
                        <div class="promo-band-legend">
                            <span class="legend-item"><span class="legend-color" style="background: rgba(128,128,128,0.3);"></span> No Promo</span>
                            <span class="legend-item"><span class="legend-color" style="background: rgba(255,193,7,0.3);"></span> MEDIUM</span>
                            <span class="legend-item"><span class="legend-color" style="background: rgba(255,152,0,0.3);"></span> HIGH</span>
                            <span class="legend-item"><span class="legend-color" style="background: rgba(244,67,54,0.3);"></span> MEGA</span>
                        </div>
                    </div>
                    <div class="promo-analysis-grid" id="promoAnalysisGrid"></div>
                    <div class="promo-info-card">
                        <h4><i class="fas fa-info-circle"></i> How to read this</h4>
                        <p>The <strong>Uplift by Volume Impact</strong> tables show how metric values change during promo weeks at each impact level (MEDIUM, HIGH, MEGA) vs baseline. A 1.5x uplift = 50% increase.</p>
                        <p>The <strong>Promo Type × Volume Impact</strong> cross-tab shows performance by promotion type and volume impact level.</p>
                    </div>
                </div>
            </section>
        </main>
    </div>

    <!-- Chart Modal -->
    <div id="chartModal" class="chart-modal">
        <div class="chart-modal-content">
            <div class="chart-modal-header">
                <h3 id="modalTitle">Chart</h3>
                <button class="modal-close-btn" onclick="closeChartModal()"><i class="fas fa-times"></i></button>
            </div>
            <div class="chart-modal-body">
                <div id="modalChartContainer" class="modal-chart-container"></div>
                <div id="modalForecastStats" class="forecast-stats modal-forecast-stats"></div>
            </div>
        </div>
    </div>

    <footer class="footer">Amazon Haul EU5 Forecasting Dashboard | Static Report v{BUILD_VERSION} | 85% Confidence Interval</footer>

    <script>
        // === EMBEDDED DATA ===
        const dashboardData = {data_json};
        const baseForecasts = {base_fc_json};
        const promoForecasts = {promo_fc_json};
        const statisticsData = {stats_json};
        const accuracyData = {acc_json};
        const latestWeekData = {lw_json};
        const promoAnalysisData = {pa_json};
        const promoRegressors = {pr_json};
        const discountValues = {dv_json};
        const hasManualForecast = {'true' if has_manual_forecast else 'false'};
        const hasPromoScores = {'true' if has_promo_scores else 'false'};
        const promoFormat = '{promo_format}';

        const METRICS = ['Net Ordered Units', 'Transits', 'Transit Conversion', 'UPO'];
        const MARKETPLACES = ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'];
        const MP_COLORS = {{
            'EU5': {{line:'#667eea', fill:'rgba(102,126,234,0.2)'}},
            'UK': {{line:'#ff9900', fill:'rgba(255,153,0,0.2)'}},
            'DE': {{line:'#00d9ff', fill:'rgba(0,217,255,0.2)'}},
            'FR': {{line:'#ff6b9d', fill:'rgba(255,107,157,0.2)'}},
            'IT': {{line:'#00e676', fill:'rgba(0,230,118,0.2)'}},
            'ES': {{line:'#ffeb3b', fill:'rgba(255,235,59,0.2)'}}
        }};
        const viColors = {{
            0: {{bg:'rgba(0,0,0,0)', border:'transparent'}},
            1: {{bg:'rgba(255,193,7,0.15)', border:'rgba(255,193,7,0.4)'}},
            2: {{bg:'rgba(255,152,0,0.18)', border:'rgba(255,152,0,0.5)'}},
            3: {{bg:'rgba(244,67,54,0.20)', border:'rgba(244,67,54,0.5)'}}
        }};

        let currentMetric = 'Net Ordered Units';
        let currentStatsView = 'total';
        let showManualForecast = true;
        let showPromoOverlay = false;
        let showPromoUplift = true;
        let selectedMarketplaces = ['EU5','UK','DE','FR','IT','ES'];

        // === HELPERS ===
        function formatValue(v, m) {{
            if(v==null||v==undefined||isNaN(v)) return '-';
            if(m==='Transit Conversion') return (v*100).toFixed(2)+'%';
            if(m==='UPO') return v.toFixed(2);
            if(Math.abs(v)>=1000000) return (v/1000000).toFixed(2)+'M';
            if(Math.abs(v)>=1000) return (v/1000).toFixed(1)+'K';
            return Math.round(v).toLocaleString();
        }}
        function formatNumber(num) {{
            if(num==null||num==undefined) return '-';
            if(Math.abs(num)>=1e6) return (num/1e6).toFixed(2)+'M';
            if(Math.abs(num)>=1000) return (num/1000).toFixed(1)+'K';
            if(Math.abs(num)<1&&num!==0) return (num*100).toFixed(2)+'%';
            return num.toFixed(0).replace(/\\B(?=(\\d{{3}})+(?!\\d))/g,',');
        }}
        function getDevClass(d) {{
            if(d==null||d==undefined) return '';
            const a=Math.abs(d);
            if(a<20) return 'dev-green';
            if(a<30) return 'dev-yellow';
            return 'dev-red';
        }}
        function formatDateToWeek(dateStr) {{
            const date=new Date(dateStr);
            const target=new Date(date.valueOf());
            const dow=date.getDay();
            const dayOff=dow===0?-3:(4-dow);
            target.setDate(date.getDate()+dayOff);
            const isoYear=target.getFullYear();
            const jan4=new Date(isoYear,0,4);
            const jan4Day=jan4.getDay();
            const monW1=new Date(jan4);
            monW1.setDate(jan4.getDate()-(jan4Day===0?6:jan4Day-1));
            const wn=Math.floor((target-monW1)/(7*864e5))+1;
            return 'Wk'+String(wn).padStart(2,'0')+' '+isoYear;
        }}
        function resizeAllCharts() {{
            setTimeout(()=>{{
                document.querySelectorAll('[id^="chart-"]').forEach(el=>{{if(el&&el.data)Plotly.Plots.resize(el);}});
                const dc=document.getElementById('deviationChart');
                if(dc&&dc.data)Plotly.Plots.resize(dc);
            }},100);
        }}
        function getActiveForecast() {{ return showPromoUplift ? promoForecasts : baseForecasts; }}
        function getMpName(c) {{ return {{'UK':'United Kingdom','DE':'Germany','FR':'France','IT':'Italy','ES':'Spain','EU5':'EU5 Consolidated'}}[c]||c; }}

        // === THEME ===
        // Set light mode as default
        (function(){{
            document.documentElement.setAttribute('data-theme','light');
            const di=document.getElementById('darkIcon');
            const li=document.getElementById('lightIcon');
            if(di) di.classList.remove('active');
            if(li) li.classList.add('active');
        }})();
        document.getElementById('themeToggle').addEventListener('click',function(){{
            const h=document.documentElement, c=h.getAttribute('data-theme');
            const isLight=c==='light';
            h.setAttribute('data-theme',isLight?'dark':'light');
            document.getElementById('darkIcon').classList.toggle('active',isLight);
            document.getElementById('lightIcon').classList.toggle('active',!isLight);
            updateCharts(); resizeAllCharts();
        }});

        // === TABS ===
        document.querySelectorAll('.tab-btn').forEach(btn=>{{
            btn.addEventListener('click',function(){{
                document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
                this.classList.add('active');
                document.getElementById('tab-'+this.dataset.tab).classList.add('active');
                if(this.dataset.tab==='promo-analysis') populatePromoAnalysis();
                resizeAllCharts();
            }});
        }});

        // === CONTROLS ===
        document.getElementById('metricSelect').addEventListener('change',function(){{
            currentMetric=this.value;
            document.getElementById('currentMetricLabel').textContent=currentMetric;
            updateDashboard(); resizeAllCharts();
        }});
        document.getElementById('manualForecastToggle').addEventListener('change',function(){{
            showManualForecast=this.checked; updateCharts(); resizeAllCharts();
        }});
        document.getElementById('promoOverlayToggle')?.addEventListener('change',function(){{
            showPromoOverlay=this.checked; updateCharts(); resizeAllCharts();
        }});
        document.getElementById('promoUpliftToggle')?.addEventListener('change',function(){{
            showPromoUplift=this.checked; updateCharts(); resizeAllCharts();
        }});
        document.querySelectorAll('.mp-checkbox').forEach(cb=>{{
            cb.addEventListener('change',function(){{
                selectedMarketplaces=Array.from(document.querySelectorAll('.mp-checkbox:checked')).map(c=>c.value);
                updateDashboard(); resizeAllCharts();
            }});
        }});
        document.querySelectorAll('.stats-toggle-btn').forEach(btn=>{{
            btn.addEventListener('click',function(){{
                document.querySelectorAll('.stats-toggle-btn').forEach(b=>b.classList.remove('active'));
                this.classList.add('active');
                currentStatsView=this.dataset.view;
                updateStats(); updateCharts(); resizeAllCharts();
            }});
        }});
        document.getElementById('deviationMetricSelect')?.addEventListener('change',function(){{ updateDeviations(); resizeAllCharts(); }});
        document.getElementById('deviationMpSelect')?.addEventListener('change',function(){{ updateDeviations(); resizeAllCharts(); }});
        document.getElementById('promoMetricSelect')?.addEventListener('change',function(){{ populatePromoAnalysis(); }});
        document.getElementById('promoTypeFilter')?.addEventListener('change',function(){{ populatePromoAnalysis(); }});

        // === CHART MODAL ===
        function openChartModal(mp) {{
            const modal=document.getElementById('chartModal');
            document.getElementById('modalTitle').innerHTML='<span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span> '+getMpName(mp)+' - '+currentMetric;
            modal.classList.add('active');
            document.body.style.overflow='hidden';
            setTimeout(()=>renderChart(mp,currentMetric,true),50);
        }}
        function closeChartModal() {{
            document.getElementById('chartModal').classList.remove('active');
            document.body.style.overflow='';
            Plotly.purge('modalChartContainer');
        }}
        document.getElementById('chartModal').addEventListener('click',function(e){{ if(e.target.id==='chartModal') closeChartModal(); }});
        document.addEventListener('keydown',function(e){{ if(e.key==='Escape') closeChartModal(); }});

        // === STATS ===
        function updateStats() {{
            const grid=document.getElementById('statsGrid');
            grid.innerHTML='';
            selectedMarketplaces.forEach(mp=>{{
                const stats=statisticsData[currentMetric]&&statisticsData[currentMetric][mp];
                if(!stats) return;
                let pv,av,mn,mx,lbl;
                if(currentStatsView==='cw'){{ pv=stats.cw_value; av=pv; mn=pv; mx=pv; lbl='CW'; }}
                else if(currentStatsView==='t4w'){{ pv=stats.t4w_total; av=stats.t4w_avg; mn=stats.t4w_min; mx=stats.t4w_max; lbl='T4W'; }}
                else{{ pv=stats.total; av=stats.average; mn=stats.min; mx=stats.max; lbl='Total'; }}
                const card=document.createElement('div');
                card.className='stat-card';
                card.innerHTML='<div class="stat-card-header"><h4><span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span> '+getMpName(mp)+'</h4></div><div class="stat-card-body"><div class="stat-item"><div class="value">'+formatValue(pv,currentMetric)+'</div><div class="label">'+lbl+' Total</div></div><div class="stat-item"><div class="value">'+formatValue(av,currentMetric)+'</div><div class="label">'+lbl+' Avg</div></div><div class="stat-item"><div class="value">'+formatValue(mn,currentMetric)+'</div><div class="label">Min</div></div><div class="stat-item"><div class="value">'+formatValue(mx,currentMetric)+'</div><div class="label">Max</div></div></div>';
                grid.appendChild(card);
            }});
        }}

        // === CHARTS ===
        function updateCharts() {{
            const grid=document.getElementById('chartsGrid');
            grid.innerHTML='';
            selectedMarketplaces.forEach(mp=>{{
                const card=document.createElement('div');
                card.className='chart-card clickable';
                card.setAttribute('onclick',"openChartModal('"+mp+"')");
                card.title='Click to expand';
                const chartId='chart-'+mp;
                card.innerHTML='<div class="chart-card-header"><h4><span class="chart-icon mp-flag '+mp.toLowerCase()+'">'+mp+'</span> '+getMpName(mp)+'</h4><span class="expand-icon"><i class="fas fa-expand-alt"></i></span></div><div class="chart-container" id="'+chartId+'"></div><div class="forecast-stats" id="forecast-stats-'+mp+'"></div>';
                grid.appendChild(card);
                renderChart(mp,currentMetric,false);
            }});
        }}

        function renderChart(mp,metric,isModal) {{
            const cid=isModal?'modalChartContainer':'chart-'+mp;
            const sid=isModal?'modalForecastStats':'forecast-stats-'+mp;
            const container=document.getElementById(cid);
            const statsC=document.getElementById(sid);
            if(!container) return;
            const mData=dashboardData[metric]&&dashboardData[metric][mp];
            if(!mData){{ container.innerHTML='<p style="text-align:center;color:var(--text-muted);padding:2rem;">No data</p>'; return; }}
            const colors=MP_COLORS[mp];
            const isDark=!document.documentElement.getAttribute('data-theme')||document.documentElement.getAttribute('data-theme')!=='light';
            const wks=mData.weeks||[];
            const vals=mData.values||[];
            const forecasts=getActiveForecast();
            const fc=forecasts[metric]&&forecasts[metric][mp];
            const traces=[];

            // Historical
            traces.push({{x:wks,y:vals,type:'scatter',mode:'lines+markers',name:'Historical',line:{{color:colors.line,width:2}},marker:{{size:isModal?6:4}}}});

            // Manual FC
            if(hasManualForecast&&showManualForecast&&mData.manual_forecast){{
                const mfW=mData.manual_weeks||[];
                const mfV=mData.manual_forecast||[];
                const fW=[],fV=[];
                mfW.forEach((w,i)=>{{if(mfV[i]!=null&&!isNaN(mfV[i])){{fW.push(w);fV.push(mfV[i]);}}}});
                if(fW.length>0) traces.push({{x:fW,y:fV,type:'scatter',mode:'lines+markers',name:'Manual FC',line:{{color:'#e040fb',width:2,dash:'dot'}},marker:{{size:isModal?6:4,symbol:'square'}}}});
            }}

            // Model FC
            if(fc){{
                const fcW=fc.dates.map(d=>formatDateToWeek(d));
                // CI
                if(fc.upper_bound&&fc.lower_bound){{
                    traces.push({{x:[...fcW,...fcW.slice().reverse()],y:[...fc.upper_bound,...fc.lower_bound.slice().reverse()],type:'scatter',fill:'toself',fillcolor:colors.fill,line:{{color:'transparent',width:0}},name:'85% CI',showlegend:true,hoverinfo:'skip'}});
                }}
                const fcLabel=showPromoUplift?'Model FC +Promo':'Model FC';
                traces.push({{x:fcW,y:fc.values,type:'scatter',mode:'lines+markers',name:fcLabel,line:{{color:colors.line,width:2,dash:'dash'}},marker:{{size:isModal?6:4,symbol:'diamond'}}}});
            }}

            // Promo overlay shapes
            const shapes=[];
            if(hasPromoScores&&showPromoOverlay&&promoRegressors){{
                wks.forEach((w,idx)=>{{
                    const r=promoRegressors[mp]&&promoRegressors[mp][w];
                    if(r){{
                        const vi=r.volume_impact||0;
                        if(vi>0&&viColors[vi]){{
                            shapes.push({{type:'rect',xref:'x',yref:'paper',x0:idx-0.4,x1:idx+0.4,y0:0,y1:1,fillcolor:viColors[vi].bg,line:{{width:0}},layer:'below'}});
                        }}
                    }}
                }});
                if(fc){{
                    const fcW=fc.dates.map(d=>formatDateToWeek(d));
                    fcW.forEach((w)=>{{
                        const r=promoRegressors[mp]&&promoRegressors[mp][w];
                        if(r){{
                            const vi=r.volume_impact||0;
                            if(vi>0&&viColors[vi]){{
                                shapes.push({{type:'rect',xref:'x',yref:'paper',x0:w,x1:w,y0:0,y1:1,fillcolor:viColors[vi].bg,line:{{color:viColors[vi].border,width:1.5,dash:'dot'}},layer:'below'}});
                            }}
                        }}
                    }});
                }}
            }}

            // Y range from historical + manual FC only
            const sv=[...vals.filter(v=>v!=null&&!isNaN(v)),...(mData.manual_forecast||[]).filter(v=>v!=null&&!isNaN(v))];
            const yMax=sv.length>0?Math.max(...sv)*1.15:100;

            const layout={{
                paper_bgcolor:'transparent',plot_bgcolor:'transparent',
                font:{{color:isDark?'rgba(255,255,255,0.7)':'rgba(26,26,46,0.8)',family:'Inter',size:isModal?12:10}},
                margin:isModal?{{l:80,r:40,t:40,b:80}}:{{l:60,r:30,t:30,b:60}},
                xaxis:{{gridcolor:isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',tickangle:-45,tickfont:{{size:isModal?11:9}}}},
                yaxis:{{gridcolor:isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',tickformat:metric==='Transit Conversion'?'.2%':'.2s',tickfont:{{size:isModal?11:9}},automargin:true,range:[0,yMax],autorange:false}},
                legend:{{orientation:'h',y:isModal?-0.15:-0.25,x:0.5,xanchor:'center',font:{{size:isModal?12:10}}}},
                hovermode:'x unified',shapes:shapes
            }};
            Plotly.newPlot(cid,traces,layout,{{responsive:true,displayModeBar:isModal,displaylogo:false}});

            // Stats below chart
            if(statsC&&fc){{
                let fcV=fc.values, vl='Total';
                if(currentStatsView==='cw'){{fcV=fcV.slice(0,1);vl='CW';}}
                else if(currentStatsView==='t4w'){{fcV=fcV.slice(0,4);vl='T4W';}}
                const fcT=fcV.reduce((a,b)=>a+b,0);
                const fcA=fcT/fcV.length;
                const mn=fc.model||'SARIMAX';
                let h='<div class="forecast-stat"><div class="value">'+formatValue(fcT,metric)+'</div><div class="label">Model FC '+vl+'</div></div><div class="forecast-stat"><div class="value">'+formatValue(fcA,metric)+'</div><div class="label">'+vl+' Avg</div></div><div class="forecast-stat"><div class="value">'+mn+(showPromoUplift?' +Promo':'')+'</div><div class="label">Model</div></div>';
                // WMAPE
                if(hasManualForecast&&accuracyData&&accuracyData[currentStatsView]&&accuracyData[currentStatsView][metric]&&accuracyData[currentStatsView][metric][mp]){{
                    const acc=accuracyData[currentStatsView][metric][mp];
                    if(acc&&acc.wmape!=null){{
                        const cls=acc.wmape<20?'good':(acc.wmape<30?'medium':'poor');
                        h+='<div class="forecast-stat"><div class="value accuracy-'+cls+'">'+acc.wmape.toFixed(1)+'%</div><div class="label">WMAPE</div></div>';
                    }}
                }}
                statsC.innerHTML=h;
            }}
        }}

        // === LATEST WEEK ===
        function updateLatestWeek() {{
            if(!latestWeekData||!latestWeekData.data) return;
            document.getElementById('latestWeekLabel').textContent=latestWeekData.latest_week||'--';
            const tb=document.getElementById('latestWeekTableBody');
            tb.innerHTML='';
            MARKETPLACES.forEach(mp=>{{
                const d=latestWeekData.data[mp]; if(!d) return;
                let r='<tr><td class="mp-cell"><div class="mp-flag '+mp.toLowerCase()+'">'+mp+'</div></td>';
                METRICS.forEach(m=>{{
                    const x=d[m]||{{}};
                    const a=formatValue(x.actual,m);
                    const f=x.manual_forecast!=null?formatValue(x.manual_forecast,m):'-';
                    const dv=x.manual_dev_pct;
                    const dc=getDevClass(dv);
                    const ds=dv!=null?(dv>0?'+':'')+dv.toFixed(1)+'%':'-';
                    r+='<td class="value-cell">'+a+'</td><td class="forecast-cell">'+f+'</td><td class="deviation-cell '+dc+'">'+ds+'</td>';
                }});
                r+='</tr>'; tb.innerHTML+=r;
            }});
        }}

        // === HISTORIC DEVIATIONS ===
        function updateDeviations() {{
            const metric=document.getElementById('deviationMetricSelect').value;
            const mp=document.getElementById('deviationMpSelect').value;
            const mData=dashboardData[metric]&&dashboardData[metric][mp];
            if(!mData) return;
            const aw=mData.weeks||[], av=mData.values||[];
            const mfW=mData.manual_weeks||[], mfV=mData.manual_forecast||[];
            const mfMap={{}};
            mfW.forEach((w,i)=>{{if(i<mfV.length&&mfV[i]!=null&&!isNaN(mfV[i]))mfMap[w]=mfV[i];}});
            const cw=[],ca=[],cf=[],cd=[];
            aw.forEach((w,i)=>{{
                const a=av[i]; const f=mfMap[w];
                if(a!=null&&!isNaN(a)&&f!=null&&f!==0){{
                    cw.push(w);ca.push(a);cf.push(f);
                    cd.push(((a-f)/f)*100);
                }}
            }});
            // Chart
            const colors=MP_COLORS[mp];
            const isDark=!document.documentElement.getAttribute('data-theme')||document.documentElement.getAttribute('data-theme')!=='light';
            const allV=[...ca,...cf];
            const yM=allV.length>0?Math.max(...allV)*1.1:100;
            const dM=cd.length>0?Math.max(Math.abs(Math.min(...cd)),Math.abs(Math.max(...cd)),30)*1.2:50;
            Plotly.newPlot('deviationChart',[
                {{x:cw,y:ca,name:'Actual',type:'scatter',mode:'lines+markers',line:{{color:colors.line,width:2}},marker:{{size:8}},yaxis:'y'}},
                {{x:cw,y:cf,name:'Manual FC',type:'scatter',mode:'lines+markers',line:{{color:'#9c27b0',width:2,dash:'dot'}},marker:{{size:6}},yaxis:'y'}},
                {{x:cw,y:cd,name:'Deviation %',type:'bar',marker:{{color:cd.map(d=>d>=0?'rgba(76,175,80,0.6)':'rgba(244,67,54,0.6)')}},yaxis:'y2'}}
            ],{{
                margin:{{t:20,r:60,b:80,l:60}},paper_bgcolor:'transparent',plot_bgcolor:'transparent',
                font:{{color:isDark?'rgba(255,255,255,0.7)':'rgba(26,26,46,0.8)',family:'Inter'}},
                xaxis:{{gridcolor:isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',tickangle:-45,type:'category'}},
                yaxis:{{gridcolor:isDark?'rgba(255,255,255,0.1)':'rgba(0,0,0,0.1)',title:{{text:metric,font:{{size:10}}}},range:[0,yM]}},
                yaxis2:{{overlaying:'y',side:'right',title:{{text:'Deviation %',font:{{size:10}}}},range:[-dM,dM],zeroline:true,zerolinecolor:isDark?'rgba(255,255,255,0.3)':'rgba(0,0,0,0.3)'}},
                legend:{{orientation:'h',y:-0.3,font:{{size:10}}}},barmode:'overlay'
            }},{{responsive:true,displayModeBar:false}});
            // Table
            const tb=document.getElementById('historicDeviationsTableBody');
            tb.innerHTML='';
            [...Array(cw.length).keys()].reverse().forEach(idx=>{{
                const dc=getDevClass(cd[idx]);
                tb.innerHTML+='<tr><td class="week-cell">'+cw[idx]+'</td><td class="value-cell">'+formatValue(ca[idx],metric)+'</td><td class="forecast-cell">'+formatValue(cf[idx],metric)+'</td><td class="deviation-cell '+dc+'">'+(cd[idx]>0?'+':'')+cd[idx].toFixed(1)+'%</td><td>-</td><td>-</td></tr>';
            }});
            // Summary
            const sm=document.getElementById('deviationSummary');
            if(sm&&cd.length>0){{
                const ad=cd.map(d=>Math.abs(d));
                const avg=ad.reduce((a,b)=>a+b,0)/ad.length;
                const mx=Math.max(...ad);const mn=Math.min(...ad);
                const cls=avg<20?'summary-good':(avg<30?'summary-warn':'summary-bad');
                sm.innerHTML='<div class="deviation-summary-grid"><div class="summary-card"><div class="summary-value">'+cw.length+'</div><div class="summary-label">Compared Weeks</div></div><div class="summary-card '+cls+'"><div class="summary-value">'+avg.toFixed(1)+'%</div><div class="summary-label">Avg |Dev|</div></div><div class="summary-card"><div class="summary-value">'+mx.toFixed(1)+'%</div><div class="summary-label">Max |Dev|</div></div><div class="summary-card"><div class="summary-value">'+mn.toFixed(1)+'%</div><div class="summary-label">Min |Dev|</div></div></div>';
            }}
        }}

        // === PROMO ANALYSIS ===
        function populatePromoAnalysis() {{
            const grid=document.getElementById('promoAnalysisGrid');
            if(!grid||!promoAnalysisData) return;
            if(promoFormat==='regressors') populateRegressorGrid(grid);
            else grid.innerHTML='<p style="color:var(--text-muted);">Legacy promo format not supported in static export.</p>';
        }}
        function populateRegressorGrid(grid) {{
            const metrics=METRICS;
            const mpOrder=MARKETPLACES;
            const impactLabels=['No Promo','MEDIUM','HIGH','MEGA'];
            const impactColors={{'No Promo':'#6c757d','MEDIUM':'#ffc107','HIGH':'#ff9800','MEGA':'#f44336'}};
            const regressorLabels={{'promo_type':'Promo Type','discount_pct':'Discount %','volume_impact':'Volume Impact','promo_count':'Promo Count'}};
            const continuousRegressors=['discount_pct','promo_count'];
            const continuousUnits={{'discount_pct':'/pp','promo_count':'/promo'}};
            let h='';

            // Section 1: Uplift by Volume Impact
            h+='<h3 class="promo-section-title"><i class="fas fa-chart-bar"></i> Uplift by Volume Impact Level</h3>';
            h+='<div class="promo-matrix-container">';
            for(const metric of metrics){{
                const analysis=promoAnalysisData[metric]; if(!analysis) continue;
                h+='<div class="promo-matrix-card"><h4 class="promo-matrix-title">'+metric+'</h4><table class="promo-matrix-table"><thead><tr><th>MP</th>';
                for(const il of impactLabels) h+='<th><span style="color:'+impactColors[il]+';">'+il+'</span></th>';
                h+='</tr></thead><tbody>';
                for(const mp of mpOrder){{
                    if(!analysis[mp]) continue;
                    const ubi=analysis[mp].uplift_by_impact||{{}};
                    h+='<tr><td><span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span></td>';
                    for(const il of impactLabels){{
                        const d=ubi[il];
                        if(d&&d.count>0){{
                            if(il==='No Promo'){{
                                h+='<td title="avg '+formatNumber(d.average)+', '+d.count+' weeks">'+formatNumber(d.average)+' <span class="week-count">('+d.count+'w)</span></td>';
                            }}else{{
                                const up=d.uplift_pct||0;
                                const uc=up>10?'uplift-positive':(up<-10?'uplift-negative':'uplift-neutral');
                                const lc=d.count<=2?' <span class="low-confidence" title="Low confidence">⚠</span>':'';
                                h+='<td class="'+uc+'" title="'+(up>0?'+':'')+up.toFixed(0)+'% uplift, '+d.count+' weeks">'+(d.uplift_factor?d.uplift_factor.toFixed(2)+'x':'--')+' <span class="week-count">('+d.count+'w)</span>'+lc+'</td>';
                            }}
                        }}else{{ h+='<td class="no-data">--</td>'; }}
                    }}
                    h+='</tr>';
                }}
                h+='</tbody></table></div>';
            }}
            h+='</div>';

            // Section 2: Promo Type x VI Crosstab
            const selMetric=document.getElementById('promoMetricSelect')?.value||'Net Ordered Units';
            const selType=document.getElementById('promoTypeFilter')?.value||'all';
            h+='<h3 class="promo-section-title" style="margin-top:2rem;"><i class="fas fa-th"></i> Promo Type × Volume Impact Breakdown</h3>';
            h+='<p class="promo-section-desc">Uplift as multiplier vs baseline. <span class="low-confidence">⚠</span> = ≤2 weeks.</p>';
            const promoTypes=['HVE','Dollar Deals','Discount %'];
            const viLevels=['MEDIUM','HIGH','MEGA'];
            const viCols={{'MEDIUM':'#ffc107','HIGH':'#ff9800','MEGA':'#f44336'}};
            const ctA=promoAnalysisData[selMetric];
            if(ctA){{
                h+='<div class="promo-matrix-container">';
                const types=selType==='all'?promoTypes:[selType];
                for(const pt of types){{
                    if(pt==='Discount %'){{
                        h+=buildDiscountCrosstab(selMetric,ctA,mpOrder,viLevels,viCols);
                        continue;
                    }}
                    h+='<div class="promo-matrix-card"><h4 class="promo-matrix-title">'+pt+' — '+selMetric+'</h4><table class="promo-matrix-table"><thead><tr><th>MP</th>';
                    for(const vi of viLevels) h+='<th><span style="color:'+viCols[vi]+';">'+vi+'</span></th>';
                    h+='</tr></thead><tbody>';
                    for(const mp of mpOrder){{
                        if(!ctA[mp]) continue;
                        const ct=ctA[mp].crosstab||{{}};
                        const ptD=ct[pt]||{{}};
                        h+='<tr><td><span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span></td>';
                        for(const vi of viLevels){{
                            const cell=ptD[vi];
                            if(cell&&cell.count>0){{
                                const up=cell.uplift_pct||0;const uc=up>10?'uplift-positive':(up<-10?'uplift-negative':'uplift-neutral');
                                const lc=cell.count<=2?' <span class="low-confidence">⚠</span>':'';
                                h+='<td class="'+uc+'">'+(cell.uplift_factor?cell.uplift_factor.toFixed(2)+'x':'--')+' <span class="week-count">('+cell.count+'w)</span>'+lc+'</td>';
                            }}else h+='<td class="no-data">--</td>';
                        }}
                        h+='</tr>';
                    }}
                    h+='</tbody></table></div>';
                }}
                h+='</div>';
            }}
            // Section 3: R²
            h+='<h3 class="promo-section-title" style="margin-top:2rem;"><i class="fas fa-chart-line"></i> Regression Fit (R²)</h3>';
            h+='<div class="promo-matrix-container">';
            for(const metric of metrics){{
                const analysis=promoAnalysisData[metric]; if(!analysis) continue;
                h+='<div class="promo-matrix-card"><h4 class="promo-matrix-title">'+metric+'</h4><table class="promo-matrix-table"><thead><tr><th>MP</th>';
                for(const [rk,rl] of Object.entries(regressorLabels)) h+='<th>'+rl+'</th>';
                h+='</tr></thead><tbody>';
                for(const mp of mpOrder){{
                    if(!analysis[mp]) continue;
                    const co=analysis[mp].regression_coefficients||{{}};
                    h+='<tr><td><span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span></td>';
                    for(const rk of Object.keys(regressorLabels)){{
                        const c=co[rk];
                        if(c&&c.r_squared>0){{
                            const rs=c.r_squared;const sc=rs>0.3?'coeff-strong':(rs>0.1?'coeff-moderate':'coeff-weak');
                            h+='<td class="'+sc+'">R²='+rs.toFixed(2);
                            if(continuousRegressors.includes(rk)){{ const pi=c.pct_impact||0; h+=' <span class="week-count">'+(pi>0?'+':'')+pi.toFixed(1)+'%'+(continuousUnits[rk]||'')+'</span>'; }}
                            h+='</td>';
                        }}else h+='<td class="no-data">--</td>';
                    }}
                    h+='</tr>';
                }}
                h+='</tbody></table></div>';
            }}
            h+='</div>';
            grid.innerHTML=h;
        }}
        function buildDiscountCrosstab(sm,ctA,mps,vis,vc){{
            const selDisc=document.getElementById('discountSubFilter')?.value||'all';
            let h='<div class="promo-matrix-card" style="min-width:100%;"><h4 class="promo-matrix-title">Discount % — '+sm;
            if(discountValues&&discountValues.length>0){{
                h+=' <select id="discountSubFilter" class="discount-sub-filter" onchange="populatePromoAnalysis()">';
                h+='<option value="all"'+(selDisc==='all'?' selected':'')+'>All Discounts</option>';
                for(const dv of discountValues){{
                    const lbl=Math.round(dv)+'%';
                    h+='<option value="'+dv+'"'+(selDisc==dv?' selected':'')+'>'+lbl+'</option>';
                }}
                h+='</select>';
            }}
            h+='</h4>';
            h+='<table class="promo-matrix-table"><thead><tr><th>MP</th>';
            for(const vi of vis) h+='<th><span style="color:'+vc[vi]+';">'+vi+'</span></th>';
            h+='</tr></thead><tbody>';
            for(const mp of mps){{
                if(!ctA[mp]) continue;
                let ptD;
                if(selDisc!=='all'){{
                    const discLabel=Math.round(parseFloat(selDisc))+'%';
                    const discCt=ctA[mp].discount_crosstab||{{}};
                    ptD=discCt[discLabel]||{{}};
                }}else{{
                    const ct=ctA[mp].crosstab||{{}};
                    ptD=ct['Discount %']||{{}};
                }}
                h+='<tr><td><span class="mp-flag '+mp.toLowerCase()+'">'+mp+'</span></td>';
                for(const vi of vis){{
                    const cell=ptD[vi];
                    if(cell&&cell.count>0){{
                        const up=cell.uplift_pct||0;const uc=up>10?'uplift-positive':(up<-10?'uplift-negative':'uplift-neutral');
                        const lc=cell.count<=2?' <span class="low-confidence">⚠</span>':'';
                        h+='<td class="'+uc+'">'+(cell.uplift_factor?cell.uplift_factor.toFixed(2)+'x':'--')+' <span class="week-count">('+cell.count+'w)</span>'+lc+'</td>';
                    }}else h+='<td class="no-data">--</td>';
                }}
                h+='</tr>';
            }}
            h+='</tbody></table></div>';
            return h;
        }}

        // === INIT ===
        function updateDashboard(){{ updateStats(); updateCharts(); updateLatestWeek(); updateDeviations(); }}
        document.addEventListener('DOMContentLoaded',function(){{ updateDashboard(); resizeAllCharts(); }});
    </script>
</body>
</html>'''

    return html


def main():
    parser = argparse.ArgumentParser(description='Build static HTML dashboard')
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
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

    print("\n  [1/7] Loading Excel data...")
    dp = DataProcessor()
    success, message = dp.load_excel(args.input)
    if not success:
        print(f"  ERROR: {message}")
        sys.exit(1)
    print(f"        {message}")

    print("  [2/7] Calculating EU5 totals...")
    dp.calculate_eu5_totals()

    print("  [3/7] Extracting data...")
    data = dp.get_all_data()
    if dp.has_manual_forecast:
        mf = dp.get_manual_forecast_data()
        if mf:
            for metric in mf:
                if metric in data:
                    for mp in mf[metric]:
                        if mp in data[metric]:
                            data[metric][mp]['manual_forecast'] = mf[metric][mp].get('values', [])
                            data[metric][mp]['manual_weeks'] = mf[metric][mp].get('weeks', [])

    statistics = generate_statistics(dp)
    accuracy = generate_accuracy_metrics(dp)
    latest_week = dp.get_latest_week_overview()

    print("  [4/7] Generating promo analysis...")
    promo_analysis = generate_promo_analysis(dp)
    promo_regressors = get_promo_regressors_json(dp)
    discount_values = dp.get_distinct_discount_values() if dp.has_promo_scores else []
    promo_format = getattr(dp, 'promo_format', 'legacy')

    print("  [5/7] Generating baseline SARIMAX forecasts...")
    base_fc, promo_fc = generate_all_forecasts(dp, include_promo=dp.has_promo_scores)
    fc_count = sum(len(v) for v in base_fc.values())
    print(f"        Generated {fc_count} baseline + {sum(len(v) for v in promo_fc.values())} promo forecasts")

    print("  [6/7] Building HTML dashboard...")
    generated_at = datetime.now().strftime('%Y-%m-%d %H:%M')
    html = build_html(data, base_fc, promo_fc, statistics, accuracy, latest_week,
                      promo_analysis, promo_regressors, discount_values,
                      dp.has_manual_forecast, dp.has_promo_scores, promo_format,
                      generated_at, args.input)

    print("  [7/7] Writing output file...")
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
