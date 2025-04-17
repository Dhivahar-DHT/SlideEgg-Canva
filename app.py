from flask import Flask, request, jsonify, render_template, send_from_directory, send_file
from pptx_fabric_converter import PPTXFabricConverter
import os
import io

app = Flask(__name__)

# Configuration
ENABLE_UI = True  # Set to False to use API-only mode
UPLOAD_FOLDER = os.path.join('static', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Add static folder configuration
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET'])
def index():
    if ENABLE_UI:
        return render_template('upload.html')
    else:
        return jsonify({"message": "API-only mode. Use POST /convert endpoints to convert files."})

@app.route('/pptx-to-fabric', methods=['POST'])
def convert_pptx_to_fabric():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    if not file.filename.endswith('.pptx'):
        return jsonify({"error": "Invalid file type"}), 400

    try:
        converter = PPTXFabricConverter()
        fabric_json = converter.pptx_to_fabric(file)
        return jsonify({"fabric": fabric_json})
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("Error details:", error_details)  # This will show in your console
        return jsonify({
            "error": str(e),
            "details": error_details
        }), 500

@app.route('/fabric-to-pptx', methods=['POST'])
def convert_fabric_to_pptx():
    try:
        fabric_data = request.json
        if not fabric_data or 'fabric' not in fabric_data:
            return jsonify({"error": "No Fabric.js data provided"}), 400
        
        converter = PPTXFabricConverter()
        prs = converter.fabric_to_pptx(fabric_data['fabric'])
        
        # Save to memory buffer
        pptx_buffer = io.BytesIO()
        prs.save(pptx_buffer)
        pptx_buffer.seek(0)
        
        return send_file(
            pptx_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name='converted.pptx'
        )
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print("Error details:", error_details)
        return jsonify({
            "error": str(e),
            "details": error_details
        }), 500

# Optional: Add a route to serve images directly if needed
@app.route('/static/uploads/<path:filename>')
def serve_image(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)
