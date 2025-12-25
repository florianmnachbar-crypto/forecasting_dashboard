/**
 * Amazon Haul EU5 Forecasting Dashboard - Charts v1.1.0
 * Handles data visualization, theme switching, and user interactions
 */

// Global state
const state = {
    data: null,
    forecasts: null,
    statistics: null,
    currentMetric: 'Net Ordered Units',
    selectedMarketplaces: ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'],
    model: 'sarimax',
    seasonality: true,
    statsView: 'total', // 'total' or 't4w'
    theme: 'dark'
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

// Initialize on DOM load
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    initializeTheme();
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
        }
        
        // Load statistics
        const statsResponse = await fetch('/api/statistics');
        const statsResult = await statsResponse.json();
        
        if (statsResult.success) {
            state.statistics = statsResult.statistics;
        }
    } catch (error) {
        console.error('Error loading data:', error);
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
}

function updateCharts() {
    const chartsGrid = document.getElementById('chartsGrid');
    const metric = state.currentMetric;
    
    let html = '';
    
    state.selectedMarketplaces.forEach(mp => {
        html += `
            <div class="chart-card">
                <div class="chart-card-header">
                    <h4>
                        <span class="chart-icon mp-flag ${mp.toLowerCase()}">${mp}</span>
                        ${mp === 'EU5' ? 'EU5 Consolidated' : getMarketplaceName(mp)}
                    </h4>
                </div>
                <div class="chart-container" id="chart-${mp}"></div>
                <div class="forecast-stats" id="forecast-stats-${mp}"></div>
            </div>
        `;
    });
    
    chartsGrid.innerHTML = html;
    
    // Render each chart
    state.selectedMarketplaces.forEach(mp => {
        renderChart(mp, metric);
    });
}

function renderChart(marketplace, metric) {
    const container = document.getElementById(`chart-${marketplace}`);
    const statsContainer = document.getElementById(`forecast-stats-${marketplace}`);
    
    if (!container) return;
    
    const historicalData = state.data?.[metric]?.[marketplace];
    const forecast = state.forecasts?.[metric]?.[marketplace];
    
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
        marker: { size: 4 }
    });
    
    // Forecast trace
    if (forecast) {
        const forecastWeeks = forecast.dates.map(d => formatDateToWeek(d));
        
        // Forecast line
        traces.push({
            x: forecastWeeks,
            y: forecast.values,
            type: 'scatter',
            mode: 'lines+markers',
            name: 'Forecast',
            line: { color: colors.line, width: 2, dash: 'dash' },
            marker: { size: 4, symbol: 'diamond' }
        });
        
        // Confidence interval (85%)
        traces.push({
            x: [...forecastWeeks, ...forecastWeeks.slice().reverse()],
            y: [...forecast.upper_bound, ...forecast.lower_bound.slice().reverse()],
            type: 'scatter',
            fill: 'toself',
            fillcolor: colors.fill,
            line: { color: 'transparent' },
            name: '85% CI',
            showlegend: true,
            hoverinfo: 'skip'
        });
        
        // Update forecast stats
        if (statsContainer) {
            const modelDisplay = forecast.model || 'SARIMAX';
            const isDerived = modelDisplay.includes('Calculated') || (forecast.model_info && forecast.model_info.method === 'derived');
            
            statsContainer.innerHTML = `
                <div class="forecast-stat">
                    <div class="value">${formatNumber(forecast.values.reduce((a, b) => a + b, 0))}</div>
                    <div class="label">Forecast Total</div>
                </div>
                <div class="forecast-stat">
                    <div class="value">${formatNumber(forecast.values.reduce((a, b) => a + b, 0) / forecast.values.length)}</div>
                    <div class="label">Forecast Avg</div>
                </div>
                <div class="forecast-stat">
                    <div class="value" title="${isDerived ? 'Net Ordered Units = Transits × Conversion × UPO' : ''}">${modelDisplay}</div>
                    <div class="label">${isDerived ? 'Derived' : 'Model'}</div>
                </div>
            `;
        }
    }
    
    const layout = {
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: { 
            color: isDark ? 'rgba(255,255,255,0.7)' : 'rgba(26,26,46,0.8)',
            family: 'Inter, sans-serif'
        },
        margin: { l: 50, r: 30, t: 30, b: 60 },
        xaxis: {
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickangle: -45,
            tickfont: { size: 10 }
        },
        yaxis: {
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickformat: metric === 'Transit Conversion' ? '.2%' : ',d'
        },
        legend: {
            orientation: 'h',
            y: -0.2,
            x: 0.5,
            xanchor: 'center'
        },
        hovermode: 'x unified'
    };
    
    const config = {
        responsive: true,
        displayModeBar: false
    };
    
    Plotly.newPlot(container, traces, layout, config);
}

function formatDateToWeek(dateStr) {
    const date = new Date(dateStr);
    // Adjust for Sunday start
    const adjusted = new Date(date.getTime() + 24 * 60 * 60 * 1000);
    const year = adjusted.getFullYear();
    const startOfYear = new Date(year, 0, 1);
    const days = Math.floor((adjusted - startOfYear) / (24 * 60 * 60 * 1000));
    const weekNum = Math.ceil((days + startOfYear.getDay() + 1) / 7);
    return `Wk${String(weekNum).padStart(2, '0')} ${year}`;
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
