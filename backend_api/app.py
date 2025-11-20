from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import sys
import traceback
import zipfile
import io

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
import image_processor
import stats_processor

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

@app.route('/api/upload/templates', methods=['POST'])
def upload_templates():
    if 'files' not in request.files:
        return jsonify({"success": False, "error": "파일이 없습니다."}), 400
    
    files = request.files.getlist('files')
    customer = request.form.get('customer', '')
    
    uploaded_files = []
    errors = []
    
    for file in files:
        if file.filename == '':
            continue
        
        if file and file.filename.endswith('.pptx'):
            filename = secure_filename(file.filename)
            
            if customer:
                target_dir = os.path.join(config.BASE_TEMPLATE_DIR, customer)
            else:
                target_dir = config.BASE_TEMPLATE_DIR
            
            os.makedirs(target_dir, exist_ok=True)
            filepath = os.path.join(target_dir, filename)
            file.save(filepath)
            uploaded_files.append(filepath)
        else:
            errors.append(f"'{file.filename}'은(는) .pptx 파일이 아닙니다.")
    
    return jsonify({
        "success": len(uploaded_files) > 0,
        "uploaded_files": uploaded_files,
        "errors": errors
    })

@app.route('/api/upload/images', methods=['POST'])
def upload_images():
    if 'files' not in request.files:
        return jsonify({"success": False, "error": "파일이 없습니다."}), 400
    
    files = request.files.getlist('files')
    customer = request.form.get('customer', '')
    
    if not customer:
        return jsonify({"success": False, "error": "customer가 필요합니다."}), 400
    
    uploaded_files = []
    errors = []
    
    target_dir = os.path.join(config.BASE_IMAGE_DIR, customer)
    os.makedirs(target_dir, exist_ok=True)
    
    for file in files:
        if file.filename == '':
            continue
        
        if file and file.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif')):
            filename = secure_filename(file.filename)
            filepath = os.path.join(target_dir, filename)
            file.save(filepath)
            uploaded_files.append(filepath)
        else:
            errors.append(f"'{file.filename}'은(는) 지원되지 않는 이미지 형식입니다.")
    
    return jsonify({
        "success": len(uploaded_files) > 0,
        "uploaded_files": uploaded_files,
        "errors": errors
    })

@app.route('/api/download/results', methods=['GET'])
def download_results():
    customer = request.args.get('customer', None)
    
    if customer:
        target_dir = os.path.join(config.OUTPUT_DIR, customer)
        if not os.path.exists(target_dir):
            return jsonify({"success": False, "error": f"고객사 '{customer}' 결과가 없습니다."}), 404
        
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(target_dir):
                for file in files:
                    if file.endswith('.pptx'):
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, config.OUTPUT_DIR)
                        zipf.write(file_path, arcname)
        
        memory_file.seek(0)
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'{customer}_results.zip'
        )
    else:
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(config.OUTPUT_DIR):
                for file in files:
                    if file.endswith('.pptx'):
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, config.OUTPUT_DIR)
                        zipf.write(file_path, arcname)
        
        memory_file.seek(0)
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='all_results.zip'
        )

@app.route('/api/files/list', methods=['GET'])
def list_files():
    file_type = request.args.get('type', 'results')
    customer = request.args.get('customer', None)
    
    if file_type == 'results':
        base_dir = config.OUTPUT_DIR
    elif file_type == 'templates':
        base_dir = config.BASE_TEMPLATE_DIR
    else:
        return jsonify({"success": False, "error": "잘못된 타입입니다."}), 400
    
    files_list = []
    
    if customer:
        target_dir = os.path.join(base_dir, customer)
        if os.path.exists(target_dir):
            for root, dirs, files in os.walk(target_dir):
                for file in files:
                    if file.endswith('.pptx'):
                        rel_path = os.path.relpath(os.path.join(root, file), base_dir)
                        files_list.append(rel_path)
    else:
        for root, dirs, files in os.walk(base_dir):
            for file in files:
                if file.endswith('.pptx'):
                    rel_path = os.path.relpath(os.path.join(root, file), base_dir)
                    files_list.append(rel_path)
    
    return jsonify({
        "success": True,
        "files": files_list,
        "count": len(files_list)
    })

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
