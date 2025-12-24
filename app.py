"""
Amazon Haul EU5 Forecasting Dashboard
Flask Application - Main Entry Point
"""

import os
import json
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from data_processor import DataProcessor
from forecaster import Forecaster

# Initialize Flask app
app = Flask(__name__)

# App version
APP_VERSION = "1.1.0"

# Configuration
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global data processor instance
data_processor = None
current_file = None


def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    """Render the main dashboard page"""
    return render_template('dashboard.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload"""
    global data_processor, current_file
    
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Please upload an Excel file (.xlsx or .xls)'}), 400
    
    try:
        # Save the file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Process the file
        data_processor = DataProcessor()
        success, message = data_processor.load_excel(filepath)
        
        if success:
            current_file = filename
            # Recalculate EU5 totals
            data_processor.calculate_eu5_totals()
            return jsonify({
                'success': True,
                'message': message,
                'filename': filename,
                'metrics': list(data_processor.data.keys()),
                'marketplaces': list(set(mp for metric in data_processor.data.values() for mp in metric.keys()))
            })
        else:
            return jsonify({'success': False, 'error': message}), 400
            
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/data', methods=['GET'])
def get_data():
    """Get all loaded data"""
    global data_processor
    
    if data_processor is None:
        return jsonify({'success': False, 'error': 'No data loaded. Please upload a file first.'}), 400
    
    try:
        data = data_processor.get_all_data()
        return jsonify({
            'success': True,
            'data': data,
            'metrics': DataProcessor.METRICS,
            'marketplaces': DataProcessor.MARKETPLACES
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/statistics', methods=['GET'])
def get_statistics():
    """Get summary statistics for all metrics and marketplaces"""
    global data_processor
    
    if data_processor is None:
        return jsonify({'success': False, 'error': 'No data loaded'}), 400
    
    try:
        stats = {}
        for metric in DataProcessor.METRICS:
            stats[metric] = {}
            for mp in DataProcessor.MARKETPLACES:
                stat = data_processor.get_summary_statistics(metric, mp)
                if stat:
                    stats[metric][mp] = stat
        
        return jsonify({
            'success': True,
            'statistics': stats
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/forecast', methods=['POST'])
def generate_forecast():
    """Generate forecast for specified metric and marketplace"""
    global data_processor
    
    if data_processor is None:
        return jsonify({'success': False, 'error': 'No data loaded'}), 400
    
    try:
        params = request.get_json()
        metric = params.get('metric', 'Net Ordered Units')
        marketplace = params.get('marketplace', 'UK')
        model_type = params.get('model', 'sarimax')
        use_seasonality = params.get('seasonality', True)
        
        # Get the data
        df = data_processor.get_dataframe(metric, marketplace)
        
        if df is None or df.empty:
            return jsonify({
                'success': False,
                'error': f'No data available for {metric} - {marketplace}'
            }), 400
        
        # Generate forecast
        forecaster = Forecaster(forecast_horizon=12)
        forecast = forecaster.generate_forecast(df, model_type=model_type, use_seasonality=use_seasonality)
        
        if forecast is None:
            return jsonify({
                'success': False,
                'error': 'Failed to generate forecast. Insufficient data.'
            }), 400
        
        # Calculate forecast statistics
        forecast_stats = {
            'total': round(sum(forecast['values']), 2),
            'average': round(sum(forecast['values']) / len(forecast['values']), 2),
            'min': round(min(forecast['values']), 2),
            'max': round(max(forecast['values']), 2)
        }
        
        return jsonify({
            'success': True,
            'forecast': forecast,
            'forecast_statistics': forecast_stats,
            'metric': metric,
            'marketplace': marketplace
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/forecast/all', methods=['POST'])
def generate_all_forecasts():
    """Generate forecasts for all metrics and marketplaces"""
    global data_processor
    
    if data_processor is None:
        return jsonify({'success': False, 'error': 'No data loaded'}), 400
    
    try:
        params = request.get_json()
        model_type = params.get('model', 'sarimax')
        use_seasonality = params.get('seasonality', True)
        
        forecasts = {}
        forecaster = Forecaster(forecast_horizon=12)
        
        for metric in DataProcessor.METRICS:
            forecasts[metric] = {}
            for mp in DataProcessor.MARKETPLACES:
                df = data_processor.get_dataframe(metric, mp)
                
                if df is not None and not df.empty and len(df) >= 4:
                    forecast = forecaster.generate_forecast(df, model_type=model_type, use_seasonality=use_seasonality)
                    if forecast:
                        forecasts[metric][mp] = forecast
        
        return jsonify({
            'success': True,
            'forecasts': forecasts,
            'model': model_type,
            'seasonality': use_seasonality
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/status', methods=['GET'])
def get_status():
    """Get current application status"""
    global data_processor, current_file
    
    return jsonify({
        'success': True,
        'data_loaded': data_processor is not None,
        'current_file': current_file,
        'metrics': DataProcessor.METRICS,
        'marketplaces': DataProcessor.MARKETPLACES
    })


# Static file serving for development
@app.route('/static/<path:filename>')
def serve_static(filename):
    """Serve static files"""
    return send_from_directory('static', filename)


if __name__ == '__main__':
    print("\n" + "="*60)
    print("  Amazon Haul EU5 Forecasting Dashboard")
    print("="*60)
    print("\n  Starting server at http://localhost:5000")
    print("  Press Ctrl+C to stop the server\n")
    print("="*60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
