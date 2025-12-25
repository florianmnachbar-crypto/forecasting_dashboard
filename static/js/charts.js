/**
 * Amazon Haul EU5 Forecasting Dashboard - Charts v2.0.0
 * Handles data visualization, theme switching, and user interactions
 * v2.0.0 - Added manual forecast support with toggle and accuracy metrics
 */

// Global state
const state = {
    data: null,
    forecasts: null,
    manualForecast: null,
    accuracy: null,
    statistics: null,
    currentMetric: 'Net Ordered Units',
    selectedMarketplaces: ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'],
    model: 'sarimax',
    seasonality: true,
    showManualForecast: true,
    statsView: 'total', // 'total' or 't4w'
    theme: 'dark',
    hasManualForecast: false
};

// Marketplace colors
const mpColors = {
    'EU5': { line: '#667eea', fill: 'rgba(102, 126, 234, 0.2)' },
    'UK': { line: '#ff9900', fill: 'rgba(255, 153, 0, 0.2)' },
    'DE': { line: '#00d9ff', fill: 'rgba(0, 217, 255, 0.2)' },
    'FR': { line: '#ff6b9d', fill: 'rgba(255, 107, 157, 0.2)' },
    'IT': { line: '#00e676', fill: 'rgba(0, 230, 118, 0.2)' },
    'ES': { line: '#ffeb3b', fill: 'rgba(255, 235, 59, 0.2)' }
};

// Manual forecast color (purple/magenta)
const manualForecastColor = { line: '#e040fb', fill: 'rgba(224, 64, 251, 0.15)' };

// Initialize on DOM load
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    initializeTheme();
    initializeModal();
    checkExistingData();
});

function initializeEventListeners() {
    // File upload
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const uploadBtn = document.getElementById('uploadBtn');
    
    dropzone?.addEventListener('click', () => fileInput?.click());
    dropzone?.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    dropzone?.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone?.addEventListener('drop', handleFileDrop);
    fileInput?.addEventListener('change', handleFileSelect);
    uploadBtn?.addEventListener('click', () => fileInput?.click());
    
    // Controls
    document.getElementById('metricSelect')?.addEventListener('change', handleMetricChange);
    document.getElementById('modelSelect')?.addEventListener('change', handleModelChange);
    document.getElementById('seasonalityToggle')?.addEventListener('change', handleSeasonalityChange);
    document.getElementById('manualForecastToggle')?.addEventListener('change', handleManualForecastToggle);
    document.getElementById('refreshBtn')?.addEventListener('click', refreshDashboard);
    
    // Marketplace checkboxes
    document.querySelectorAll('.mp-checkbox').forEach(cb => {
        cb.addEventListener('change', handleMarketplaceChange);
    });
    
    // Theme toggle
    document.getElementById('themeToggle')?.addEventListener('click', toggleTheme);
    
    // Stats toggle (Total vs T4W)
    document.querySelectorAll('.stats-toggle-btn').forEach(btn => {
        btn.addEventListener('click', handleStatsToggle);
    });
}

function initializeTheme() {
    // Check localStorage for saved theme
    const savedTheme = localStorage.getItem('dashboardTheme') || 'dark';
    state.theme = savedTheme;
    applyTheme(savedTheme);
}

function initializeModal() {
    // Create modal HTML if it doesn't exist
    if (!document.getElementById('chartModal')) {
        const modalHTML = `
            <div id="chartModal" class="chart-modal">
                <div class="chart-modal-content">
                    <div class="chart-modal-header">
                        <h3 id="modalTitle">Chart</h3>
                        <button class="modal-close-btn" onclick="closeChartModal()">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                    <div class="chart-modal-body">
                        <div id="modalChartContainer" class="modal-chart-container"></div>
                        <div id="modalForecastStats" class="forecast-stats modal-forecast-stats"></div>
                    </div>
                </div>
            </div>
        `;
        document.body.insertAdjacentHTML('beforeend', modalHTML);
        
        // Close modal on backdrop click
        document.getElementById('chartModal').addEventListener('click', (e) => {
            if (e.target.id === 'chartModal') {
                closeChartModal();
            }
        });
        
        // Close modal on Escape key
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                closeChartModal();
            }
        });
    }
}

function toggleTheme() {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    applyTheme(state.theme);
    localStorage.setItem('dashboardTheme', state.theme);
    
    // Re-render charts with new theme
    if (state.data && state.forecasts) {
        updateCharts();
    }
}

function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    
    const darkIcon = document.getElementById('darkIcon');
    const lightIcon = document.getElementById('lightIcon');
    
    if (theme === 'dark') {
        darkIcon?.classList.add('active');
        lightIcon?.classList.remove('active');
    } else {
        darkIcon?.classList.remove('active');
        lightIcon?.classList.add('active');
    }
}

function handleStatsToggle(e) {
    const view = e.target.dataset.view;
    state.statsView = view;
    
    // Update button states
    document.querySelectorAll('.stats-toggle-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.view === view);
    });
    
    // Update stats display
    if (state.statistics) {
        updateStatistics();
    }
}

function handleManualForecastToggle(e) {
    state.showManualForecast = e.target.checked;
    updateCharts();
}

async function checkExistingData() {
    try {
        const response = await fetch('/api/status');
        const result = await response.json();
        
        if (result.data_loaded) {
            showLoading();
            await loadData();
            await generateForecasts();
            showDashboard();
            hideLoading();
        }
    } catch (error) {
        console.error('Error checking status:', error);
    }
}

function handleFileDrop(e) {
    e.preventDefault();
    e.target.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (file) uploadFile(file);
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) uploadFile(file);
}

async function uploadFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showToast('error', 'Invalid File', 'Please upload an Excel file (.xlsx or .xls)');
        return;
    }
    
    showLoading();
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            showToast('success', 'File Uploaded', `Successfully loaded ${result.filename}`);
            updateFileStatus(result.filename);
            await loadData();
            await generateForecasts();
            showDashboard();
        } else {
            showToast('error', 'Upload Failed', result.error);
        }
    } catch (error) {
        showToast('error', 'Error', 'Failed to upload file');
        console.error('Upload error:', error);
    }
    
    hideLoading();
}

async function loadData() {
    try {
        const response = await fetch('/api/data');
        const result = await response.json();
        
        if (result.success) {
            state.data = result.data;
            state.hasManualForecast = result.has_manual_forecast || false;
            
            if (result.manual_forecast) {
                state.manualForecast = result.manual_forecast;
            }
        }
        
        // Load statistics
        const statsResponse = await fetch('/api/statistics');
        const statsResult = await statsResponse.json();
        
        if (statsResult.success) {
            state.statistics = statsResult.statistics;
        }
        
        // Load accuracy metrics if manual forecast exists
        if (state.hasManualForecast) {
            const accResponse = await fetch('/api/accuracy');
            const accResult = await accResponse.json();
            
            if (accResult.success && accResult.accuracy) {
                state.accuracy = accResult.accuracy;
            }
        }
        
        // Update UI to show/hide manual forecast toggle
        updateManualForecastToggleVisibility();
        
    } catch (error) {
        console.error('Error loading data:', error);
    }
}

function updateManualForecastToggleVisibility() {
    const toggleGroup = document.getElementById('manualForecastToggleGroup');
    if (toggleGroup) {
        toggleGroup.style.display = state.hasManualForecast ? 'flex' : 'none';
    }
}

async function generateForecasts() {
    try {
        const response = await fetch('/api/forecast/all', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                model: state.model,
                seasonality: state.seasonality
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            state.forecasts = result.forecasts;
            updateModelInfo(result.model);
        }
    } catch (error) {
        console.error('Error generating forecasts:', error);
    }
}

function showDashboard() {
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('dashboardSection').style.display = 'block';
    document.getElementById('modelInfo').classList.add('show');
    
    updateStatistics();
    updateCharts();
    updateAccuracyPanel();
}

function updateAccuracyPanel() {
    const accuracyPanel = document.getElementById('accuracyPanel');
    if (!accuracyPanel) return;
    
    if (!state.hasManualForecast || !state.accuracy) {
        accuracyPanel.style.display = 'none';
        return;
    }
    
    accuracyPanel.style.display = 'block';
    
    const metric = state.currentMetric;
    const accuracy = state.accuracy[metric];
    
    if (!accuracy) {
        accuracyPanel.innerHTML = '<p style="color: var(--text-muted);">No accuracy data for this metric</p>';
        return;
    }
    
    let html = '<div class="accuracy-grid">';
    
    state.selectedMarketplaces.forEach(mp => {
        if (!accuracy[mp]) return;
        
        const acc = accuracy[mp];
        const accuracyClass = acc.accuracy >= 80 ? 'good' : acc.accuracy >= 60 ? 'medium' : 'poor';
        const biasClass = acc.bias > 0 ? 'over' : 'under';
        
        html += `
            <div class="accuracy-card">
                <div class="accuracy-header">
                    <span class="mp-flag ${mp.toLowerCase()}">${mp}</span>
                    <span class="accuracy-value ${accuracyClass}">${acc.accuracy}%</span>
                </div>
                <div class="accuracy-details">
                    <div class="accuracy-item">
                        <span class="label">WMAPE</span>
                        <span class="value">${acc.wmape}%</span>
                    </div>
                    <div class="accuracy-item">
                        <span class="label">Bias</span>
                        <span class="value ${biasClass}">${acc.bias > 0 ? '+' : ''}${acc.bias}%</span>
                    </div>
                    <div class="accuracy-item">
                        <span class="label">Overlap</span>
                        <span class="value">${acc.overlap_weeks} wks</span>
                    </div>
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    accuracyPanel.innerHTML = html;
}

function updateStatistics() {
    const statsGrid = document.getElementById('statsGrid');
    const metric = state.currentMetric;
    const stats = state.statistics?.[metric];
    
    if (!stats) {
        statsGrid.innerHTML = '<p style="color: var(--text-muted);">No statistics available</p>';
        return;
    }
    
    let html = '';
    
    state.selectedMarketplaces.forEach(mp => {
        if (!stats[mp]) return;
        
        const s = stats[mp];
        const isT4W = state.statsView === 't4w';
        
        // Calculate T4W stats from the data if available
        let displayStats;
        if (isT4W && state.data?.[metric]?.[mp]) {
            const values = state.data[metric][mp].values;
            const last4 = values.slice(-4);
            displayStats = {
                total: last4.reduce((a, b) => a + b, 0),
                average: last4.reduce((a, b) => a + b, 0) / last4.length,
                min: Math.min(...last4),
                max: Math.max(...last4),
                count: last4.length
            };
        } else {
            displayStats = s;
        }
        
        html += `
            <div class="stat-card">
                <div class="stat-card-header">
                    <h4>
                        <span class="mp-flag ${mp.toLowerCase()}">${mp}</span>
                        ${mp === 'EU5' ? 'All Marketplaces' : getMarketplaceName(mp)}
                    </h4>
                </div>
                <div class="stat-card-body">
                    <div class="stat-item">
                        <div class="value">${formatNumber(displayStats.total)}</div>
                        <div class="label">${isT4W ? 'T4W Total' : 'Total'}</div>
                    </div>
                    <div class="stat-item">
                        <div class="value">${formatNumber(displayStats.average)}</div>
                        <div class="label">${isT4W ? 'T4W Avg' : 'Average'}</div>
                    </div>
                    <div class="stat-item">
                        <div class="value">${formatNumber(displayStats.min)}</div>
                        <div class="label">${isT4W ? 'T4W Min' : 'Minimum'}</div>
                    </div>
                    <div class="stat-item">
                        <div class="value">${formatNumber(displayStats.max)}</div>
                        <div class="label">${isT4W ? 'T4W Max' : 'Maximum'}</div>
                    </div>
                </div>
            </div>
        `;
    });
    
    statsGrid.innerHTML = html;
    document.getElementById('currentMetricLabel').textContent = metric;
    
    // Update accuracy panel when metric changes
    updateAccuracyPanel();
}

function updateCharts() {
    const chartsGrid = document.getElementById('chartsGrid');
    const metric = state.currentMetric;
    
    let html = '';
    
    state.selectedMarketplaces.forEach(mp => {
        html += `
            <div class="chart-card clickable" onclick="openChartModal('${mp}')" title="Click to expand">
                <div class="chart-card-header">
                    <h4>
                        <span class="chart-icon mp-flag ${mp.toLowerCase()}">${mp}</span>
                        ${mp === 'EU5' ? 'EU5 Consolidated' : getMarketplaceName(mp)}
                    </h4>
                    <span class="expand-icon"><i class="fas fa-expand-alt"></i></span>
                </div>
                <div class="chart-container" id="chart-${mp}"></div>
                <div class="forecast-stats" id="forecast-stats-${mp}"></div>
            </div>
        `;
    });
    
    chartsGrid.innerHTML = html;
    
    // Render each chart
    state.selectedMarketplaces.forEach(mp => {
        renderChart(mp, metric, false);
    });
}

function openChartModal(marketplace) {
    const modal = document.getElementById('chartModal');
    const modalTitle = document.getElementById('modalTitle');
    const metric = state.currentMetric;
    
    // Set modal title
    const mpName = marketplace === 'EU5' ? 'EU5 Consolidated' : getMarketplaceName(marketplace);
    modalTitle.innerHTML = `<span class="mp-flag ${marketplace.toLowerCase()}">${marketplace}</span> ${mpName} - ${metric}`;
    
    // Show modal
    modal.classList.add('active');
    document.body.style.overflow = 'hidden'; // Prevent background scroll
    
    // Render expanded chart after modal is visible (slight delay for animation)
    setTimeout(() => {
        renderChart(marketplace, metric, true);
    }, 50);
}

function closeChartModal() {
    const modal = document.getElementById('chartModal');
    modal.classList.remove('active');
    document.body.style.overflow = ''; // Restore scrolling
    
    // Clean up the chart
    Plotly.purge('modalChartContainer');
}

function renderChart(marketplace, metric, isModal = false) {
    const containerId = isModal ? 'modalChartContainer' : `chart-${marketplace}`;
    const statsContainerId = isModal ? 'modalForecastStats' : `forecast-stats-${marketplace}`;
    
    const container = document.getElementById(containerId);
    const statsContainer = document.getElementById(statsContainerId);
    
    if (!container) return;
    
    const historicalData = state.data?.[metric]?.[marketplace];
    const forecast = state.forecasts?.[metric]?.[marketplace];
    const manualForecast = state.manualForecast?.[metric]?.[marketplace];
    
    if (!historicalData) {
        container.innerHTML = '<p style="text-align: center; color: var(--text-muted); padding: 2rem;">No data available</p>';
        return;
    }
    
    const colors = mpColors[marketplace];
    const isDark = state.theme === 'dark';
    
    // Use week labels for x-axis
    const weekLabels = historicalData.weeks || historicalData.dates.map(d => formatDateToWeek(d));
    
    const traces = [];
    
    // Historical data trace
    traces.push({
        x: weekLabels,
        y: historicalData.values,
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Historical',
        line: { color: colors.line, width: 2 },
        marker: { size: isModal ? 6 : 4 }
    });
    
    // Manual forecast trace (dotted line) - show if toggle is on
    if (manualForecast && state.showManualForecast) {
        const manualWeeks = manualForecast.weeks || manualForecast.dates.map(d => formatDateToWeek(d));
        
        traces.push({
            x: manualWeeks,
            y: manualForecast.values,
            type: 'scatter',
            mode: 'lines+markers',
            name: 'Manual FC',
            line: { color: manualForecastColor.line, width: 2, dash: 'dot' },
            marker: { size: isModal ? 6 : 4, symbol: 'square' }
        });
    }
    
    // Model forecast trace (dashed line)
    if (forecast) {
        const forecastWeeks = forecast.dates.map(d => formatDateToWeek(d));
        
        // Confidence interval (85%) - add FIRST so it renders behind the line
        // Filter out null/NaN values to prevent cutouts in the polygon
        if (forecast.upper_bound && forecast.lower_bound && 
            forecast.upper_bound.length > 0 && forecast.lower_bound.length > 0) {
            
            // Build arrays of valid CI points (no nulls/NaN)
            const validCIData = [];
            for (let i = 0; i < forecastWeeks.length; i++) {
                const upper = forecast.upper_bound[i];
                const lower = forecast.lower_bound[i];
                if (upper != null && lower != null && !isNaN(upper) && !isNaN(lower)) {
                    validCIData.push({
                        week: forecastWeeks[i],
                        upper: upper,
                        lower: lower
                    });
                }
            }
            
            // Only render CI if we have valid points
            if (validCIData.length > 0) {
                const ciWeeks = validCIData.map(d => d.week);
                const ciUpper = validCIData.map(d => d.upper);
                const ciLower = validCIData.map(d => d.lower);
                
                traces.push({
                    x: [...ciWeeks, ...ciWeeks.slice().reverse()],
                    y: [...ciUpper, ...ciLower.slice().reverse()],
                    type: 'scatter',
                    fill: 'toself',
                    fillcolor: colors.fill,
                    line: { color: 'transparent', width: 0 },
                    name: '85% CI',
                    showlegend: true,
                    hoverinfo: 'skip'
                });
            }
        }
        
        // Forecast line - add AFTER CI so it renders on top
        traces.push({
            x: forecastWeeks,
            y: forecast.values,
            type: 'scatter',
            mode: 'lines+markers',
            name: 'Model FC',
            line: { color: colors.line, width: 2, dash: 'dash' },
            marker: { size: isModal ? 6 : 4, symbol: 'diamond' }
        });
        
        // Update forecast stats
        if (statsContainer) {
            const modelDisplay = forecast.model || 'SARIMAX';
            const isDerived = modelDisplay.includes('Calculated') || (forecast.model_info && forecast.model_info.method === 'derived');
            
            // Calculate accuracy info if available
            const accuracy = state.accuracy?.[metric]?.[marketplace];
            let accuracyHtml = '';
            if (accuracy && state.hasManualForecast) {
                const accClass = accuracy.accuracy >= 80 ? 'good' : accuracy.accuracy >= 60 ? 'medium' : 'poor';
                accuracyHtml = `
                    <div class="forecast-stat">
                        <div class="value accuracy-badge ${accClass}">${accuracy.accuracy}%</div>
                        <div class="label">Manual FC Acc</div>
                    </div>
                `;
            }
            
            statsContainer.innerHTML = `
                <div class="forecast-stat">
                    <div class="value">${formatNumber(forecast.values.reduce((a, b) => a + b, 0))}</div>
                    <div class="label">Model FC Total</div>
                </div>
                <div class="forecast-stat">
                    <div class="value">${formatNumber(forecast.values.reduce((a, b) => a + b, 0) / forecast.values.length)}</div>
                    <div class="label">Model FC Avg</div>
                </div>
                <div class="forecast-stat">
                    <div class="value" title="${isDerived ? 'Net Ordered Units = Transits × Conversion × UPO' : ''}">${modelDisplay}</div>
                    <div class="label">${isDerived ? 'Derived' : 'Model'}</div>
                </div>
                ${accuracyHtml}
            `;
        }
    }
    
    // Calculate dynamic left margin based on max value
    const allValues = [
        ...historicalData.values, 
        ...(forecast?.values || []), 
        ...(forecast?.upper_bound || []),
        ...(manualForecast?.values || [])
    ];
    const maxVal = Math.max(...allValues);
    let leftMargin = 60; // Default
    if (maxVal >= 10000000) leftMargin = 80;
    else if (maxVal >= 1000000) leftMargin = 75;
    else if (maxVal >= 100000) leftMargin = 70;
    
    const layout = {
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: { 
            color: isDark ? 'rgba(255,255,255,0.7)' : 'rgba(26,26,46,0.8)',
            family: 'Inter, sans-serif',
            size: isModal ? 12 : 10
        },
        margin: isModal 
            ? { l: leftMargin + 20, r: 40, t: 40, b: 80 }
            : { l: leftMargin, r: 30, t: 30, b: 60 },
        xaxis: {
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickangle: -45,
            tickfont: { size: isModal ? 11 : 9 }
        },
        yaxis: {
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickformat: metric === 'Transit Conversion' ? '.2%' : '.2s',
            tickfont: { size: isModal ? 11 : 9 },
            automargin: true
        },
        legend: {
            orientation: 'h',
            y: isModal ? -0.15 : -0.25,
            x: 0.5,
            xanchor: 'center',
            font: { size: isModal ? 12 : 10 }
        },
        hovermode: 'x unified'
    };
    
    const config = {
        responsive: true,
        displayModeBar: isModal,
        modeBarButtonsToRemove: ['pan2d', 'select2d', 'lasso2d', 'autoScale2d'],
        displaylogo: false
    };
    
    Plotly.newPlot(container, traces, layout, config);
}

function formatDateToWeek(dateStr) {
    const date = new Date(dateStr);
    
    // Use proper ISO 8601 week calculation
    // ISO week: Week 1 is the week containing January 4th
    // Weeks start on Monday
    const target = new Date(date.valueOf());
    
    // Set to nearest Thursday (for ISO week calculation)
    // Monday = 1, Sunday = 0 -> we need to adjust Sunday to be 7
    const dayOfWeek = date.getDay();
    const dayOffset = dayOfWeek === 0 ? -3 : (4 - dayOfWeek);
    target.setDate(date.getDate() + dayOffset);
    
    // Get the year of the Thursday (which determines the ISO week year)
    const isoYear = target.getFullYear();
    
    // Get January 4th of that year (always in week 1)
    const jan4 = new Date(isoYear, 0, 4);
    
    // Find the Monday of week 1
    const jan4Day = jan4.getDay();
    const mondayOfWeek1 = new Date(jan4);
    mondayOfWeek1.setDate(jan4.getDate() - (jan4Day === 0 ? 6 : jan4Day - 1));
    
    // Calculate week number
    const diffMs = target - mondayOfWeek1;
    const weekNum = Math.floor(diffMs / (7 * 24 * 60 * 60 * 1000)) + 1;
    
    return `Wk${String(weekNum).padStart(2, '0')} ${isoYear}`;
}

function handleMetricChange(e) {
    state.currentMetric = e.target.value;
    updateStatistics();
    updateCharts();
}

function handleModelChange(e) {
    state.model = e.target.value;
    refreshForecasts();
}

function handleSeasonalityChange(e) {
    state.seasonality = e.target.checked;
    refreshForecasts();
}

function handleMarketplaceChange() {
    const checkboxes = document.querySelectorAll('.mp-checkbox:checked');
    state.selectedMarketplaces = Array.from(checkboxes).map(cb => cb.value);
    updateStatistics();
    updateCharts();
}

async function refreshForecasts() {
    showLoading();
    await generateForecasts();
    updateCharts();
    hideLoading();
}

async function refreshDashboard() {
    showLoading();
    await loadData();
    await generateForecasts();
    updateStatistics();
    updateCharts();
    hideLoading();
    showToast('success', 'Refreshed', 'Dashboard data has been refreshed');
}

function updateFileStatus(filename) {
    const fileStatus = document.getElementById('fileStatus');
    const fileNameSpan = document.getElementById('fileName');
    
    fileStatus.classList.remove('no-file');
    fileNameSpan.textContent = filename;
}

function updateModelInfo(model) {
    document.getElementById('modelName').textContent = `Model: ${model.toUpperCase()}`;
    document.getElementById('modelDetails').textContent = `Seasonality: ${state.seasonality ? 'Enabled' : 'Disabled'}`;
}

function getMarketplaceName(code) {
    const names = {
        'UK': 'United Kingdom',
        'DE': 'Germany',
        'FR': 'France',
        'IT': 'Italy',
        'ES': 'Spain',
        'EU5': 'EU5 Consolidated'
    };
    return names[code] || code;
}

function formatNumber(num) {
    if (num === null || num === undefined) return '-';
    if (Math.abs(num) >= 1000000) {
        return (num / 1000000).toFixed(2) + 'M';
    } else if (Math.abs(num) >= 1000) {
        return (num / 1000).toFixed(1) + 'K';
    } else if (Math.abs(num) < 1 && num !== 0) {
        return (num * 100).toFixed(2) + '%';
    }
    return num.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function showLoading() {
    document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
}

function showToast(type, title, message) {
    const container = document.getElementById('toastContainer');
    const icons = {
        success: 'fas fa-check-circle',
        error: 'fas fa-exclamation-circle',
        warning: 'fas fa-exclamation-triangle'
    };
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        <i class="${icons[type]}"></i>
        <div class="toast-content">
            <div class="toast-title">${title}</div>
            <div class="toast-message">${message}</div>
        </div>
        <button class="toast-close" onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => toast.remove(), 5000);
}
