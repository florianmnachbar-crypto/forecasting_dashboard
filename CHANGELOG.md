# Changelog

All notable changes to the Amazon Haul EU5 Forecasting Dashboard will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.3.0] - 2026-01-20

### Added
- **Transit Conversion Cap**: SARIMAX forecasts for Transit Conversion are now capped at 10% (0.10) to prevent unrealistic extrapolations
- **Y-Axis Scaling Fix**: Charts now scale based on historical data and manual forecast only, ensuring readability even when model forecasts explode
- **Team Sharing**: Added `setup_and_run.bat` one-click installer for easy team distribution
- Updated README with comprehensive team sharing instructions

### Fixed
- ES forecast explosion issue caused by Transit Conversion spike extrapolation
- Chart readability when model forecasts significantly exceed historical ranges

## [2.2.0] - 2026-01-12

### Added
- **SARIMAX Promo Regressor**: Promo scores can now be used as exogenous variables in SARIMAX model
- **Promo Floor Logic**: When promo toggle is ON, promo weeks cannot decrease forecast below baseline
- **Derived Net Ordered Units**: NOU is now calculated as Transits × Transit Conversion × UPO
- Baseline forecast used for weeks without promo scores (score = 1.0)

### Changed
- Model label shows "+Promo" when promo regressor is applied
- Model label shows "(Floored)" when floor logic prevents forecast decrease

## [2.1.0] - 2026-01-10

### Added
- **Promo Scores Integration**: Support for "Promo Scores" sheet in Excel input
- **Promo Overlay**: Visual promo band overlays on charts (No/Low, Light, Medium, Strong)
- **Promo Analysis Tab**: New tab showing uplift analysis by promo band per marketplace
- Promo overlay toggle in sidebar
- Promo uplift toggle for SARIMAX model

### Changed
- Data processor now parses promo scores from multiple week formats
- Charts show color-coded promo intensity backgrounds

## [2.0.0] - 2026-01-05

### Added
- **Manual Forecast Support**: Upload "Forecast" sheet alongside actuals for comparison
- **Accuracy Metrics**: WMAPE, MAPE, and Bias calculations for manual forecast vs actuals
- **Manual Forecast Toggle**: Show/hide manual forecast line on charts
- **Historic Deviations Tab**: View week-by-week deviation history
- **Latest Week Overview Tab**: Quick view of most recent week's performance
- Accuracy badge display on chart cards
- Metric name aliases (CVR → Transit Conversion)

### Changed
- Dashboard now displays both historical actuals and manual forecast
- Manual forecast shown as purple dotted line on charts
- Accuracy color coding: Green (<20%), Yellow (20-30%), Red (>30%)

## [1.0.0] - 2025-12-20

### Added
- Initial release
- **SARIMAX Forecasting**: Seasonal ARIMA with eXogenous factors
- **Prophet Support**: Facebook's Prophet forecasting library
- **12-week Forecast Horizon**: Forecasts generated for 12 weeks ahead
- **EU5 Marketplace Support**: UK, DE, FR, IT, ES, EU5 consolidated
- **Metrics**: Net Ordered Units, Transits, Transit Conversion, UPO
- Interactive Plotly.js charts with confidence intervals
- Dark/Light theme toggle
- Export to CSV and Excel
- File upload via drag-and-drop
- Summary statistics (Total, Average, Min, Max, T4W)
- Modern responsive dashboard design

---

## Version Numbering

- **Major (X.0.0)**: Breaking changes or major feature additions
- **Minor (0.X.0)**: New features, backwards compatible
- **Patch (0.0.X)**: Bug fixes, minor improvements
