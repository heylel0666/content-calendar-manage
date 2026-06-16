from flask import Flask, render_template, request, jsonify, send_file
from flask_cors import CORS
from excel_manager import ExcelCalendarManager
import json

app = Flask(__name__)
CORS(app)
manager = ExcelCalendarManager()

@app.route('/')
def index():
    """Serve the main page"""
    return render_template('index.html')

@app.route('/api/entries', methods=['GET'])
def get_entries():
    """Get all entries"""
    entries = manager.get_all_entries()
    return jsonify(entries)

@app.route('/api/entries', methods=['POST'])
def add_entry():
    """Add a new entry"""
    data = request.json
    success = manager.add_entry(
        date=data['date'],
        main_page=data['main_page'],
        main_status=data['main_status'],
        canada_content=data['canada_content'],
        canada_status=data['canada_status'],
        youtube_content=data['youtube_content'],
        youtube_status=data['youtube_status']
    )
    return jsonify({'success': success})

@app.route('/api/entries/<int:row>', methods=['PUT'])
def update_entry(row):
    """Update an entry"""
    data = request.json
    success = manager.update_entry(
        row=row,
        date=data['date'],
        main_page=data['main_page'],
        main_status=data['main_status'],
        canada_content=data['canada_content'],
        canada_status=data['canada_status'],
        youtube_content=data['youtube_content'],
        youtube_status=data['youtube_status']
    )
    return jsonify({'success': success})

@app.route('/api/entries/<int:row>', methods=['DELETE'])
def delete_entry(row):
    """Delete an entry"""
    success = manager.delete_entry(row)
    return jsonify({'success': success})

@app.route('/api/search', methods=['GET'])
def search_entries():
    """Search entries"""
    keyword = request.args.get('q', '')
    results = manager.search_entries(keyword)
    return jsonify(results)

@app.route('/api/filter', methods=['GET'])
def filter_entries():
    """Filter entries by status"""
    status = request.args.get('status', '')
    results = manager.filter_by_status(status)
    return jsonify(results)

@app.route('/api/export', methods=['GET'])
def export_file():
    """Download the Excel file"""
    return send_file(manager.filename, as_attachment=True)
if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)