from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
import image_processor
import stats_processor
import template_generator
import support_manager
import recommendation_generator

app = Flask(__name__)
CORS(app)

os.makedirs(config.BASE_TEMPLATE_DIR, exist_ok=True)
os.makedirs(config.BASE_IMAGE_DIR, exist_ok=True)
os.makedirs(config.OUTPUT_DIR_WITH_IMAGES, exist_ok=True)
os.makedirs(config.OUTPUT_DIR, exist_ok=True)

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({
        "status": "healthy",
        "grafana_url": config.GRAFANA_URL,
        "grafana_configured": bool(config.API_KEY),
        "directories": {
            "templates": os.path.exists(config.BASE_TEMPLATE_DIR),
            "images": os.path.exists(config.BASE_IMAGE_DIR),
            "output_images": os.path.exists(config.OUTPUT_DIR_WITH_IMAGES),
            "output_final": os.path.exists(config.OUTPUT_DIR)
        }
    })

@app.route('/api/config', methods=['GET'])
def get_config():
    return jsonify({
        "grafana_url": config.GRAFANA_URL,
        "dashboard_map": config.DASHBOARD_MAP,
        "directories": {
            "templates": config.BASE_TEMPLATE_DIR,
            "images": config.BASE_IMAGE_DIR,
            "output_with_images": config.OUTPUT_DIR_WITH_IMAGES,
            "output_final": config.OUTPUT_DIR
        }
    })

@app.route('/api/customers', methods=['GET'])
def list_customers():
    customers = []
    if os.path.exists(config.BASE_IMAGE_DIR):
        for item in os.listdir(config.BASE_IMAGE_DIR):
            item_path = os.path.join(config.BASE_IMAGE_DIR, item)
            if os.path.isdir(item_path) and item not in ['template', 'completed_with_images', 'completed_final']:
                customers.append(item)
    return jsonify({"customers": sorted(customers)})

@app.route('/api/process/images', methods=['POST'])
def process_images():
    data = request.get_json() or {}
    customer_name = data.get('customer_name', None)
    
    results = image_processor.process_images(customer_name)
    
    return jsonify(results), 200 if results["success"] else 400

@app.route('/api/process/statistics', methods=['POST'])
def process_statistics():
    data = request.get_json() or {}
    customer_name = data.get('customer_name', None)
    
    results = stats_processor.process_statistics(customer_name)
    
    return jsonify(results), 200 if results["success"] else 400

@app.route('/api/process/all', methods=['POST'])
def process_all():
    results = {
        "success": True,
        "images": {},
        "statistics": {},
        "errors": []
    }
    
    try:
        results["images"] = image_processor.process_images()
        
        if not results["images"]["success"]:
            results["success"] = False
            results["errors"].append("이미지 삽입 실패")
        else:
            results["statistics"] = stats_processor.process_statistics()
            
            if not results["statistics"]["success"]:
                results["success"] = False
                results["errors"].append("통계 삽입 실패")
    
    except Exception as e:
        results["success"] = False
        results["errors"].append(f"전체 프로세스 실패: {str(e)}")
    
    return jsonify(results), 200 if results["success"] else 400

@app.route('/api/template/generate', methods=['POST'])
def generate_template():
    data = request.get_json() or {}
    customer_name = data.get('customer_name')
    selected_panels = data.get('selected_panels', None)
    
    if not customer_name:
        return jsonify({"success": False, "errors": ["customer_name이 필요합니다."]}), 400
    
    results = template_generator.generate_template(customer_name, selected_panels)
    
    return jsonify(results), 200 if results["success"] else 400

@app.route('/api/support/add', methods=['POST'])
def add_support():
    data = request.get_json() or {}
    file_path = data.get('file_path')
    support_entries = data.get('support_entries', [])
    
    if not file_path:
        return jsonify({"success": False, "errors": ["file_path가 필요합니다."]}), 400
    
    if not support_entries:
        return jsonify({"success": False, "errors": ["support_entries가 필요합니다."]}), 400
    
    results = support_manager.add_support_history(file_path, support_entries)
    
    return jsonify(results), 200 if results["success"] else 400

@app.route('/api/recommendation/auto', methods=['POST'])
def auto_recommend():
    data = request.get_json() or {}
    file_path = data.get('file_path')
    customer_name = data.get('customer_name')
    threshold = data.get('threshold', 90)
    
    if not file_path:
        return jsonify({"success": False, "errors": ["file_path가 필요합니다."]}), 400
    
    if not customer_name:
        return jsonify({"success": False, "errors": ["customer_name이 필요합니다."]}), 400
    
    results = recommendation_generator.add_recommendations(file_path, customer_name, threshold)
    
    return jsonify(results), 200 if results["success"] else 400

@app.errorhandler(Exception)
def handle_error(error):
    app.logger.error(f"Unhandled exception: {str(error)}\n{traceback.format_exc()}")
    return jsonify({
        "error": str(error),
        "traceback": traceback.format_exc()
    }), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5001))
    app.run(host='0.0.0.0', port=port, debug=True)
