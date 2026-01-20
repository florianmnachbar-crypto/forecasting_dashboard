/**
 * Amazon Haul EU5 Forecasting Dashboard - Charts v2.3.0
 * Handles data visualization, theme switching, and user interactions
 */

// Global state
const state = {
    data: null,
    forecasts: null,
    manualForecast: null,
    forecastUplift: null,
    accuracy: null,
    statistics: null,
    promoScores: null,
    promoAnalysis: null,
    currentMetric: 'Net Ordered Units',
    selectedMarketplaces: ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'],
    model: 'sarimax',
    seasonality: true,
    showManualForecast: true,
    showPromoOverlay: true,
    showPromoUplift: false,
    statsView: 'total', // 'total' or 't4w'
    theme: 'dark',
    hasManualForecast: false,
    hasPromoScores: false
};

// Promo band colors (for chart overlays)
const promoBandColors = {
    'No/Low Promo': { bg: 'rgba(128, 128, 128, 0.1)', border: 'rgba(128, 128, 128, 0.3)' },
    'Light Promo': { bg: 'rgba(100, 181, 246, 0.15)', border: 'rgba(100, 181, 246, 0.4)' },
    'Medium Promo': { bg: 'rgba(255, 193, 7, 0.15)', border: 'rgba(255, 193, 7, 0.4)' },
    'Strong Promo': { bg: 'rgba(255, 87, 34, 0.2)', border: 'rgba(255, 87, 34, 0.5)' }
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
    initializeTabs();
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
    document.getElementById('promoOverlayToggle')?.addEventListener('change', handlePromoOverlayToggle);
    document.getElementById('promoUpliftToggle')?.addEventListener('change', handlePromoUpliftToggle);
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
    
    // Tab navigation
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', handleTabChange);
    });
    
    // Historic deviations selects
    document.getElementById('deviationMetricSelect')?.addEventListener('change', loadHistoricDeviations);
    document.getElementById('deviationMpSelect')?.addEventListener('change', loadHistoricDeviations);
    
    // Promo analysis select
    document.getElementById('promoMetricSelect')?.addEventListener('change', populatePromoAnalysisGrid);
    
    // Export dropdown
    initializeExportDropdown();
}

function initializeExportDropdown() {
    const exportBtn = document.getElementById('exportBtn');
    const exportDropdown = document.getElementById('exportDropdown');
    const exportOptions = document.querySelectorAll('.export-option');
    
    // Toggle dropdown on button click
    exportBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        if (exportBtn.disabled) return;
        exportDropdown.classList.toggle('open');
    });
    
    // Handle export option clicks
    exportOptions.forEach(option => {
        option.addEventListener('click', (e) => {
            const format = e.currentTarget.dataset.format;
            handleExport(format);
            exportDropdown.classList.remove('open');
        });
    });
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!exportDropdown?.contains(e.target)) {
            exportDropdown?.classList.remove('open');
        }
    });
}

function updateExportButtonState() {
    const exportBtn = document.getElementById('exportBtn');
    if (exportBtn) {
        exportBtn.disabled = !state.data;
    }
}

async function handleExport(format) {
    if (!state.data) {
        showToast('error', 'No Data', 'Please upload data first before exporting');
        return;
    }
    
    showLoading();
    
    try {
        const endpoint = format === 'csv' ? '/api/export/csv' : '/api/export/excel';
        const response = await fetch(endpoint);
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Export failed');
        }
        
        // Get filename from Content-Disposition header or use default
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = `amazon_haul_eu5_export.${format === 'csv' ? 'csv' : 'xlsx'}`;
        if (contentDisposition) {
            const match = contentDisposition.match(/filename=(.+)/);
            if (match) {
                filename = match[1];
            }
        }
        
        // Download the file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showToast('success', 'Export Complete', `Data exported as ${format.toUpperCase()}`);
    } catch (error) {
        console.error('Export error:', error);
        showToast('error', 'Export Failed', error.message);
    }
    
    hideLoading();
}

function initializeTabs() {
    // Set default tab as active
    const defaultTab = document.querySelector('.tab-btn[data-tab="forecasts"]');
    if (defaultTab) {
        defaultTab.classList.add('active');
    }
}

function handleTabChange(e) {
    const tabName = e.currentTarget.dataset.tab;
    
    // Update button states
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabName);
    });
    
    // Update tab content visibility
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabName}`);
    });
    
    // Load latest week data if switching to that tab
    if (tabName === 'latest-week') {
        loadLatestWeekOverview();
    }
    
    // Load historic deviations if switching to that tab
    if (tabName === 'historic-deviations') {
        loadHistoricDeviations();
    }
    
    // Load promo analysis if switching to that tab
    if (tabName === 'promo-analysis') {
        populatePromoAnalysisGrid();
    }
}

async function loadLatestWeekOverview() {
    try {
        const response = await fetch('/api/latest-week');
        const result = await response.json();
        
        if (result.success) {
            renderLatestWeekTable(result.overview, result.has_manual_forecast);
        } else {
            console.error('Failed to load latest week overview:', result.error);
        }
    } catch (error) {
        console.error('Error loading latest week overview:', error);
    }
}

function renderLatestWeekTable(overview, hasManualForecast) {
    const tableBody = document.getElementById('latestWeekTableBody');
    const weekLabel = document.getElementById('latestWeekLabel');
    
    if (!tableBody) return;
    
    // Update the week label
    if (weekLabel && overview.latest_week) {
        weekLabel.textContent = overview.latest_week;
    }
    
    const marketplaces = ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'];
    const metrics = ['Net Ordered Units', 'Transits', 'Transit Conversion', 'UPO'];
    
    let html = '';
    
    marketplaces.forEach(mp => {
        const mpData = overview.data[mp];
        if (!mpData) return;
        
        html += `<tr>
            <td class="mp-cell"><span class="mp-flag ${mp.toLowerCase()}">${mp}</span></td>`;
        
        metrics.forEach(metric => {
            const data = mpData[metric] || {};
            
            // Format values based on metric type
            const actualDisplay = formatMetricValue(data.actual, metric);
            const forecastDisplay = formatMetricValue(data.manual_forecast, metric);
            const devPct = data.manual_dev_pct;
            
            // Get deviation color class
            const devColorClass = getDeviationColorClass(devPct);
            const devDisplay = devPct !== null && devPct !== undefined 
                ? `${devPct > 0 ? '+' : ''}${devPct.toFixed(1)}%` 
                : '--';
            
            html += `
                <td class="value-cell">${actualDisplay}</td>
                <td class="value-cell forecast-cell">${forecastDisplay}</td>
                <td class="deviation-cell ${devColorClass}">${devDisplay}</td>`;
        });
        
        html += '</tr>';
    });
    
    tableBody.innerHTML = html;
}

function formatMetricValue(value, metric) {
    if (value === null || value === undefined) return '--';
    
    if (metric === 'Transit Conversion') {
        // Show as percentage
        return (value * 100).toFixed(2) + '%';
    } else if (metric === 'UPO') {
        // Show with 2 decimal places
        return value.toFixed(2);
    } else {
        // Net Ordered Units and Transits - format with K/M suffix
        if (Math.abs(value) >= 1000000) {
            return (value / 1000000).toFixed(2) + 'M';
        } else if (Math.abs(value) >= 1000) {
            return (value / 1000).toFixed(1) + 'K';
        }
        return Math.round(value).toLocaleString();
    }
}

function getDeviationColorClass(devPct) {
    if (devPct === null || devPct === undefined) return '';
    
    const absDeviation = Math.abs(devPct);
    
    if (absDeviation < 20) {
        return 'dev-green';
    } else if (absDeviation >= 20 && absDeviation <= 30) {
        return 'dev-yellow';
    } else {
        return 'dev-red';
    }
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

async function handleStatsToggle(e) {
    const view = e.target.dataset.view;
    state.statsView = view;
    
    // Update button states
    document.querySelectorAll('.stats-toggle-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.view === view);
    });
    
    // Re-fetch accuracy with new timeframe if we have manual forecast
    if (state.hasManualForecast) {
        await loadAccuracy();
    }
    
    // Update stats display and charts (to update forecast stats)
    if (state.statistics) {
        updateStatistics();
    }
    if (state.forecasts) {
        updateCharts();
    }
}

async function loadAccuracy() {
    try {
        // Map statsView to API timeframe parameter
        const timeframe = state.statsView; // 'total', 't4w', or 'cw'
        const accResponse = await fetch(`/api/accuracy?timeframe=${timeframe}`);
        const accResult = await accResponse.json();
        
        if (accResult.success && accResult.accuracy) {
            state.accuracy = accResult.accuracy;
        }
    } catch (error) {
        console.error('Error loading accuracy:', error);
    }
}

function handleManualForecastToggle(e) {
    state.showManualForecast = e.target.checked;
    updateCharts();
}

function handlePromoOverlayToggle(e) {
    state.showPromoOverlay = e.target.checked;
    updateCharts();
}

async function handlePromoUpliftToggle(e) {
    state.showPromoUplift = e.target.checked;
    // Regenerate SARIMAX forecasts with/without promo regressor
    await refreshForecasts();
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
            await loadAccuracy();
        }
        
        // Load promo scores if available
        await loadPromoData();
        
        // Update UI to show/hide manual forecast toggle
        updateManualForecastToggleVisibility();
        
    } catch (error) {
        console.error('Error loading data:', error);
    }
}

async function loadPromoData() {
    try {
        // Load promo scores
        const promoResponse = await fetch('/api/promo-scores');
        const promoResult = await promoResponse.json();
        
        if (promoResult.success && promoResult.has_promo_scores) {
            state.promoScores = promoResult.promo_data;
            state.hasPromoScores = true;
            console.log('Promo scores loaded:', Object.keys(state.promoScores.scores || {}));
        } else {
            state.hasPromoScores = false;
        }
        
        // Load promo analysis
        if (state.hasPromoScores) {
            const analysisResponse = await fetch('/api/promo-analysis');
            const analysisResult = await analysisResponse.json();
            
            if (analysisResult.success && analysisResult.analysis) {
                state.promoAnalysis = analysisResult.analysis;
                console.log('Promo analysis loaded for metrics:', Object.keys(state.promoAnalysis));
            }
            
            // Load forecast uplift data
            const upliftResponse = await fetch('/api/forecast-uplift');
            const upliftResult = await upliftResponse.json();
            
            if (upliftResult.success && upliftResult.has_uplift_data) {
                state.forecastUplift = upliftResult.uplift_data;
                console.log('Forecast uplift loaded for metrics:', Object.keys(state.forecastUplift));
            }
        }
    } catch (error) {
        console.error('Error loading promo data:', error);
        state.hasPromoScores = false;
    }
}

// Get promo band from score
function getPromoBand(score) {
    if (score === null || score === undefined) return null;
    if (score <= 1) return 'No/Low Promo';
    if (score <= 2) return 'Light Promo';
    if (score <= 3) return 'Medium Promo';
    return 'Strong Promo';
}

// Get promo score for a specific marketplace and week
function getPromoScoreForWeek(marketplace, weekLabel) {
    if (!state.hasPromoScores || !state.promoScores?.scores) return null;
    
    const mpScores = state.promoScores.scores[marketplace];
    if (!mpScores) return null;
    
    // Try exact match
    if (mpScores[weekLabel] !== undefined) {
        return mpScores[weekLabel];
    }
    
    // Try normalized match (convert week format)
    // weekLabel might be "Wk19 2025" or "Wk01 2026"
    return null;
}

function updateManualForecastToggleVisibility() {
    const toggleGroup = document.getElementById('manualForecastToggleGroup');
    if (toggleGroup) {
        toggleGroup.style.display = state.hasManualForecast ? 'flex' : 'none';
    }
    
    // Also update promo overlay toggle visibility
    const promoToggleGroup = document.getElementById('promoOverlayToggleGroup');
    if (promoToggleGroup) {
        promoToggleGroup.style.display = state.hasPromoScores ? 'flex' : 'none';
    }
    
    // Promo uplift toggle visibility - only show when we have both promo scores and manual forecast
    const promoUpliftToggleGroup = document.getElementById('promoUpliftToggleGroup');
    if (promoUpliftToggleGroup) {
        promoUpliftToggleGroup.style.display = (state.hasPromoScores && state.hasManualForecast) ? 'flex' : 'none';
    }
    
    // Show/hide the Promo Analysis tab based on promo data availability
    const promoAnalysisTab = document.getElementById('promoAnalysisTab');
    if (promoAnalysisTab) {
        promoAnalysisTab.style.display = state.hasPromoScores ? 'inline-flex' : 'none';
    }
}

async function generateForecasts() {
    try {
        const response = await fetch('/api/forecast/all', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                model: state.model,
                seasonality: state.seasonality,
                include_promo: state.showPromoUplift  // Include promo as SARIMAX regressor if toggle is on
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
    updateExportButtonState();
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
        const isCW = state.statsView === 'cw';
        
        // Calculate stats based on view
        let displayStats;
        let viewLabel = 'Total';
        
        if (isCW && state.data?.[metric]?.[mp]) {
            // Current Week - just the latest value
            const values = state.data[metric][mp].values;
            const latestValue = values[values.length - 1];
            displayStats = {
                total: latestValue,
                average: latestValue,
                min: latestValue,
                max: latestValue,
                count: 1
            };
            viewLabel = 'CW';
        } else if (isT4W && state.data?.[metric]?.[mp]) {
            const values = state.data[metric][mp].values;
            const last4 = values.slice(-4);
            displayStats = {
                total: last4.reduce((a, b) => a + b, 0),
                average: last4.reduce((a, b) => a + b, 0) / last4.length,
                min: Math.min(...last4),
                max: Math.max(...last4),
                count: last4.length
            };
            viewLabel = 'T4W';
        } else {
            displayStats = s;
            viewLabel = 'Total';
        }
        
        // For CW, show simplified card
        if (isCW) {
            html += `
                <div class="stat-card">
                    <div class="stat-card-header">
                        <h4>
                            <span class="mp-flag ${mp.toLowerCase()}">${mp}</span>
                            ${mp === 'EU5' ? 'All Marketplaces' : getMarketplaceName(mp)}
                        </h4>
                    </div>
                    <div class="stat-card-body">
                        <div class="stat-item stat-item-large">
                            <div class="value">${formatNumber(displayStats.total)}</div>
                            <div class="label">Current Week Value</div>
                        </div>
                    </div>
                </div>
            `;
        } else {
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
                            <div class="label">${viewLabel} Total</div>
                        </div>
                        <div class="stat-item">
                            <div class="value">${formatNumber(displayStats.average)}</div>
                            <div class="label">${viewLabel} Avg</div>
                        </div>
                        <div class="stat-item">
                            <div class="value">${formatNumber(displayStats.min)}</div>
                            <div class="label">${viewLabel} Min</div>
                        </div>
                        <div class="stat-item">
                            <div class="value">${formatNumber(displayStats.max)}</div>
                            <div class="label">${viewLabel} Max</div>
                        </div>
                    </div>
                </div>
            `;
        }
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
    
    // Build promo score annotations for tooltips
    const promoAnnotations = [];
    if (state.hasPromoScores && state.showPromoOverlay) {
        weekLabels.forEach((week, idx) => {
            const score = getPromoScoreForWeek(marketplace, week);
            if (score !== null) {
                promoAnnotations.push({
                    week: week,
                    score: score,
                    band: getPromoBand(score)
                });
            }
        });
    }
    
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
    
    // Manual forecast trace (dotted line) - show if toggle is on (no uplift applied - manual FC shown as-is)
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
        
        // Update forecast stats based on statsView
        if (statsContainer) {
            let modelDisplay = forecast.model || 'SARIMAX';
            // Add promo indicator if promo regressor was used
            if (state.showPromoUplift && state.hasPromoScores && forecast.promo_info) {
                modelDisplay += ' +Promo';
            }
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
            
            // Calculate values based on stats view
            let fcValues = forecast.values;
            let viewLabel = 'Total';
            
            if (state.statsView === 'cw') {
                // Current week - just the first forecast value
                fcValues = fcValues.slice(0, 1);
                viewLabel = 'CW';
            } else if (state.statsView === 't4w') {
                // Trailing 4 weeks - take first 4 forecast values
                fcValues = fcValues.slice(0, 4);
                viewLabel = 'T4W';
            }
            
            const fcTotal = fcValues.reduce((a, b) => a + b, 0);
            const fcAvg = fcTotal / fcValues.length;
            
            statsContainer.innerHTML = `
                <div class="forecast-stat">
                    <div class="value">${formatNumber(fcTotal)}</div>
                    <div class="label">Model FC ${viewLabel}</div>
                </div>
                <div class="forecast-stat">
                    <div class="value">${formatNumber(fcAvg)}</div>
                    <div class="label">${viewLabel === 'CW' ? 'CW Value' : viewLabel + ' Avg'}</div>
                </div>
                <div class="forecast-stat">
                    <div class="value" title="${isDerived ? 'Net Ordered Units = Transits × Conversion × UPO' : ''}">${modelDisplay}</div>
                    <div class="label">${isDerived ? 'Derived' : 'Model'}</div>
                </div>
                ${accuracyHtml}
            `;
        }
    }
    
    // Calculate Y-axis range based on historical + manual forecast ONLY
    // This ensures the chart is readable even if model forecast explodes
    const scaleValues = [
        ...historicalData.values.filter(v => v != null && !isNaN(v)),
        ...(manualForecast?.values?.filter(v => v != null && !isNaN(v)) || [])
    ];
    const yMax = scaleValues.length > 0 ? Math.max(...scaleValues) * 1.15 : 100; // 15% padding
    const yMin = 0; // Start from 0
    
    // Calculate dynamic left margin based on max value (use scale values, not model forecast)
    const maxVal = yMax;
    let leftMargin = 60; // Default
    if (maxVal >= 10000000) leftMargin = 80;
    else if (maxVal >= 1000000) leftMargin = 75;
    else if (maxVal >= 100000) leftMargin = 70;
    
    // Build promo overlay shapes
    const promoShapes = [];
    if (state.hasPromoScores && state.showPromoOverlay && promoAnnotations.length > 0) {
        promoAnnotations.forEach((anno, idx) => {
            const band = anno.band;
            if (band && promoBandColors[band]) {
                promoShapes.push({
                    type: 'rect',
                    xref: 'x',
                    yref: 'paper',
                    x0: idx - 0.4,
                    x1: idx + 0.4,
                    y0: 0,
                    y1: 1,
                    fillcolor: promoBandColors[band].bg,
                    line: { width: 0 },
                    layer: 'below'
                });
            }
        });
    }
    
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
            automargin: true,
            range: [yMin, yMax],  // Fixed range based on historical + manual FC only
            autorange: false      // Disable autorange so model FC doesn't expand the scale
        },
        legend: {
            orientation: 'h',
            y: isModal ? -0.15 : -0.25,
            x: 0.5,
            xanchor: 'center',
            font: { size: isModal ? 12 : 10 }
        },
        hovermode: 'x unified',
        shapes: promoShapes
    };
    
    const config = {
        responsive: true,
        displayModeBar: isModal,
        modeBarButtonsToRemove: ['pan2d', 'select2d', 'lasso2d', 'autoScale2d'],
        displaylogo: false
    };
    
    Plotly.newPlot(container, traces, layout, config);
}

// Populate promo analysis grid
function populatePromoAnalysisGrid() {
    const grid = document.getElementById('promoAnalysisGrid');
    if (!grid || !state.promoAnalysis) return;
    
    const metric = document.getElementById('promoMetricSelect')?.value || 'Net Ordered Units';
    const analysis = state.promoAnalysis[metric];
    
    if (!analysis) {
        grid.innerHTML = '<p style="text-align: center; color: var(--text-muted);">No promo analysis data available for this metric.</p>';
        return;
    }
    
    const mpNames = {
        'UK': 'United Kingdom',
        'DE': 'Germany', 
        'FR': 'France',
        'IT': 'Italy',
        'ES': 'Spain',
        'EU5': 'EU5 (All)'
    };
    
    let html = '';
    const mpOrder = ['EU5', 'UK', 'DE', 'FR', 'IT', 'ES'];
    
    for (const mp of mpOrder) {
        if (!analysis[mp]) continue;
        
        const mpData = analysis[mp];
        const bands = mpData.bands || {};
        
        html += `
            <div class="promo-card">
                <div class="promo-card-header">
                    <h4>
                        <div class="mp-flag ${mp.toLowerCase()}">${mp}</div>
                        ${mpNames[mp]}
                    </h4>
                    <span style="color: var(--text-muted); font-size: 0.85rem;">${mpData.total_weeks_analyzed || 0} weeks</span>
                </div>
                <table class="promo-band-table">
                    <thead>
                        <tr>
                            <th>Promo Band</th>
                            <th>Avg</th>
                            <th>Uplift</th>
                            <th>Weeks</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        const bandOrder = ['No/Low Promo', 'Light Promo', 'Medium Promo', 'Strong Promo'];
        const bandClasses = {
            'No/Low Promo': 'no-promo',
            'Light Promo': 'light-promo',
            'Medium Promo': 'medium-promo',
            'Strong Promo': 'strong-promo'
        };
        
        for (const bandLabel of bandOrder) {
            const band = bands[bandLabel];
            if (!band) continue;
            
            const upliftPct = band.uplift_pct || 0;
            const upliftClass = upliftPct > 5 ? 'uplift-positive' : (upliftPct < -5 ? 'uplift-negative' : 'uplift-neutral');
            const upliftSign = upliftPct > 0 ? '+' : '';
            
            html += `
                <tr>
                    <td>
                        <span class="promo-band-name">
                            <span class="promo-band-dot ${bandClasses[bandLabel]}"></span>
                            ${bandLabel}
                        </span>
                    </td>
                    <td>${formatNumber(band.average)}</td>
                    <td class="uplift-value ${upliftClass}">${band.uplift_factor ? band.uplift_factor.toFixed(2) + 'x' : '1.00x'} (${upliftSign}${upliftPct.toFixed(0)}%)</td>
                    <td>${band.count}</td>
                </tr>
            `;
        }
        
        html += `
                    </tbody>
                </table>
                <div style="margin-top: 0.75rem; font-size: 0.8rem; color: var(--text-muted);">
                    Baseline: ${formatNumber(mpData.baseline_avg)} avg
                </div>
            </div>
        `;
    }
    
    grid.innerHTML = html || '<p style="text-align: center; color: var(--text-muted);">No promo analysis data available.</p>';
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

// Historic Deviations functions
async function loadHistoricDeviations() {
    const metricSelect = document.getElementById('deviationMetricSelect');
    const mpSelect = document.getElementById('deviationMpSelect');
    
    if (!metricSelect || !mpSelect) return;
    
    const metric = metricSelect.value;
    const marketplace = mpSelect.value;
    
    try {
        const response = await fetch(`/api/historic-deviations?metric=${encodeURIComponent(metric)}&marketplace=${encodeURIComponent(marketplace)}`);
        const result = await response.json();
        
        if (result.success) {
            renderHistoricDeviationsTable(result.deviations, result.summary, metric);
        } else {
            console.error('Failed to load historic deviations:', result.error);
            showToast('error', 'Error', result.error || 'Failed to load historic deviations');
        }
    } catch (error) {
        console.error('Error loading historic deviations:', error);
        showToast('error', 'Error', 'Failed to load historic deviations');
    }
}

function renderHistoricDeviationsTable(deviations, summary, metric) {
    const tableBody = document.getElementById('historicDeviationsTableBody');
    const summaryContainer = document.getElementById('deviationSummary');
    
    if (!tableBody) return;
    
    // Sort deviations by date ascending for chart (oldest first)
    const sortedForChart = [...deviations].sort((a, b) => new Date(a.date) - new Date(b.date));
    
    // Render the deviation chart
    renderDeviationChart(sortedForChart, metric);
    
    // Sort deviations by date descending for table (most recent first)
    const sortedDeviations = [...deviations].sort((a, b) => new Date(b.date) - new Date(a.date));
    
    let html = '';
    
    sortedDeviations.forEach(d => {
        const actualDisplay = formatMetricValue(d.actual, metric);
        const manualFcDisplay = d.manual_forecast !== null ? formatMetricValue(d.manual_forecast, metric) : '--';
        const manualDevDisplay = d.manual_dev_pct !== null 
            ? `${d.manual_dev_pct > 0 ? '+' : ''}${d.manual_dev_pct.toFixed(1)}%` 
            : '--';
        const manualDevClass = getDeviationColorClass(d.manual_dev_pct);
        
        const modelFcDisplay = d.model_forecast !== null ? formatMetricValue(d.model_forecast, metric) : '--';
        const modelDevDisplay = d.model_dev_pct !== null 
            ? `${d.model_dev_pct > 0 ? '+' : ''}${d.model_dev_pct.toFixed(1)}%` 
            : '--';
        const modelDevClass = getDeviationColorClass(d.model_dev_pct);
        
        html += `
            <tr>
                <td class="week-cell">${d.week}</td>
                <td class="value-cell">${actualDisplay}</td>
                <td class="value-cell forecast-cell">${manualFcDisplay}</td>
                <td class="deviation-cell ${manualDevClass}">${manualDevDisplay}</td>
                <td class="value-cell forecast-cell">${modelFcDisplay}</td>
                <td class="deviation-cell ${modelDevClass}">${modelDevDisplay}</td>
            </tr>
        `;
    });
    
    tableBody.innerHTML = html || '<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">No data available</td></tr>';
    
    // Render summary
    if (summaryContainer && summary) {
        let summaryHtml = '<div class="deviation-summary-grid">';
        
        summaryHtml += `
            <div class="summary-card">
                <div class="summary-value">${summary.total_weeks}</div>
                <div class="summary-label">Total Weeks</div>
            </div>
        `;
        
        if (summary.manual_forecast_weeks > 0) {
            summaryHtml += `
                <div class="summary-card">
                    <div class="summary-value">${summary.manual_forecast_weeks}</div>
                    <div class="summary-label">Manual FC Weeks</div>
                </div>
                <div class="summary-card ${getDeviationSummaryClass(summary.manual_avg_abs_dev)}">
                    <div class="summary-value">${summary.manual_avg_abs_dev !== null ? summary.manual_avg_abs_dev.toFixed(1) + '%' : '--'}</div>
                    <div class="summary-label">Manual Avg |Dev|</div>
                </div>
                <div class="summary-card">
                    <div class="summary-value">${summary.manual_avg_dev !== null ? (summary.manual_avg_dev > 0 ? '+' : '') + summary.manual_avg_dev.toFixed(1) + '%' : '--'}</div>
                    <div class="summary-label">Manual Avg Bias</div>
                </div>
            `;
        }
        
        if (summary.model_forecast_weeks > 0) {
            summaryHtml += `
                <div class="summary-card">
                    <div class="summary-value">${summary.model_forecast_weeks}</div>
                    <div class="summary-label">Model FC Weeks</div>
                </div>
                <div class="summary-card ${getDeviationSummaryClass(summary.model_avg_abs_dev)}">
                    <div class="summary-value">${summary.model_avg_abs_dev !== null ? summary.model_avg_abs_dev.toFixed(1) + '%' : '--'}</div>
                    <div class="summary-label">Model Avg |Dev|</div>
                </div>
            `;
        }
        
        summaryHtml += '</div>';
        summaryContainer.innerHTML = summaryHtml;
    }
}

function getDeviationSummaryClass(avgDev) {
    if (avgDev === null) return '';
    if (avgDev < 20) return 'summary-good';
    if (avgDev < 30) return 'summary-warn';
    return 'summary-bad';
}

function renderDeviationChart(deviations, metric) {
    const container = document.getElementById('deviationChart');
    if (!container || !deviations || deviations.length === 0) {
        if (container) {
            container.innerHTML = '<p style="text-align: center; color: var(--text-muted); padding: 2rem;">No deviation data available for chart</p>';
        }
        return;
    }
    
    const isDark = state.theme === 'dark';
    const mpSelect = document.getElementById('deviationMpSelect');
    const marketplace = mpSelect?.value || 'EU5';
    const colors = mpColors[marketplace] || mpColors['EU5'];
    
    // Prepare data arrays
    const weeks = deviations.map(d => d.week);
    const actuals = deviations.map(d => d.actual);
    const manualForecasts = deviations.map(d => d.manual_forecast);
    const modelForecasts = deviations.map(d => d.model_forecast);
    const manualDevs = deviations.map(d => d.manual_dev_pct);
    const modelDevs = deviations.map(d => d.model_dev_pct);
    
    const traces = [];
    
    // Actual values trace
    traces.push({
        x: weeks,
        y: actuals,
        type: 'scatter',
        mode: 'lines+markers',
        name: 'Actual',
        line: { color: colors.line, width: 2 },
        marker: { size: 6 },
        yaxis: 'y'
    });
    
    // Manual forecast trace (if exists)
    const hasManualFC = manualForecasts.some(v => v !== null);
    if (hasManualFC) {
        traces.push({
            x: weeks,
            y: manualForecasts,
            type: 'scatter',
            mode: 'lines+markers',
            name: 'Manual FC',
            line: { color: manualForecastColor.line, width: 2, dash: 'dot' },
            marker: { size: 6, symbol: 'square' },
            yaxis: 'y'
        });
    }
    
    // Model forecast trace (if exists)
    const hasModelFC = modelForecasts.some(v => v !== null);
    if (hasModelFC) {
        traces.push({
            x: weeks,
            y: modelForecasts,
            type: 'scatter',
            mode: 'lines+markers',
            name: 'Model FC',
            line: { color: '#00bcd4', width: 2, dash: 'dash' },
            marker: { size: 6, symbol: 'diamond' },
            yaxis: 'y'
        });
    }
    
    // Deviation bars (manual) on secondary axis
    if (hasManualFC) {
        // Color bars based on deviation threshold
        const manualBarColors = manualDevs.map(d => {
            if (d === null) return 'rgba(128,128,128,0.3)';
            const abs = Math.abs(d);
            if (abs < 20) return 'rgba(0, 230, 118, 0.7)';  // green
            if (abs < 30) return 'rgba(255, 193, 7, 0.7)';  // yellow
            return 'rgba(244, 67, 54, 0.7)';  // red
        });
        
        traces.push({
            x: weeks,
            y: manualDevs,
            type: 'bar',
            name: 'Manual Dev %',
            marker: { color: manualBarColors },
            yaxis: 'y2',
            opacity: 0.7
        });
    }
    
    const layout = {
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: { 
            color: isDark ? 'rgba(255,255,255,0.7)' : 'rgba(26,26,46,0.8)',
            family: 'Inter, sans-serif',
            size: 11
        },
        margin: { l: 70, r: 70, t: 30, b: 70 },
        xaxis: {
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickangle: -45,
            tickfont: { size: 10 }
        },
        yaxis: {
            title: metric,
            titlefont: { size: 11 },
            gridcolor: isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
            tickformat: metric === 'Transit Conversion' ? '.2%' : '.2s',
            tickfont: { size: 10 },
            side: 'left'
        },
        yaxis2: {
            title: 'Deviation %',
            titlefont: { size: 11 },
            overlaying: 'y',
            side: 'right',
            tickformat: '+.0f',
            ticksuffix: '%',
            tickfont: { size: 10 },
            gridcolor: 'transparent',
            zeroline: true,
            zerolinecolor: isDark ? 'rgba(255,255,255,0.3)' : 'rgba(0,0,0,0.3)',
            zerolinewidth: 1
        },
        legend: {
            orientation: 'h',
            y: -0.2,
            x: 0.5,
            xanchor: 'center',
            font: { size: 11 }
        },
        hovermode: 'x unified',
        barmode: 'group',
        bargap: 0.3
    };
    
    const config = {
        responsive: true,
        displayModeBar: false,
        displaylogo: false
    };
    
    Plotly.newPlot(container, traces, layout, config);
}
