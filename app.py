import os
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename

# Import the new function from our refactored script
from script import script

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/generate-report', methods=['POST'])
def generate_report_endpoint():
    file_paths = {}
    try:
        # --- 1. Get Form Data ---
        school_name = request.form.get('school_name', 'Default School Name')
        principal_name = request.form.get('principal_name', 'Default Principal Name')

        # --- 2. Get and Save All 5 Uploaded Files ---
        required_files = [
            'goals_by_team_file', 'player_data_file', 'school_summary_file',
            'educator_file', 'family_file'
        ]
        
        # Check if all files are present
        for file_key in required_files:
            if file_key not in request.files:
                return jsonify({"error": f"Missing file: {file_key}"}), 400
        
        # Save files and store their paths
        
        for file_key in required_files:
            file = request.files[file_key]
            if file.filename == '':
                return jsonify({"error": f"No file selected for {file_key}"}), 400
            
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            file_paths[file_key] = filepath

        # --- 3. Call the Report Generation Logic ---
        report_buffer = script(
            school_name=school_name,
            principal_name=principal_name,
            goals_path=file_paths['goals_by_team_file'],
            player_data_path=file_paths['player_data_file'],
            summary_path=file_paths['school_summary_file'],
            educator_path=file_paths['educator_file'],
            family_path=file_paths['family_file']
        )

        # --- 4. Send the Generated File Back to the User ---
        return send_file(
            report_buffer,
            as_attachment=True,
            download_name=f'{school_name}_FIM_Report.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        # If anything goes wrong, return a clear error message
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

    finally:
        # --- 5. Clean Up ---
        # Clean up the temporarily saved files
        for file_key in required_files:
            if file_key in file_paths and os.path.exists(file_paths[file_key]):
                os.remove(file_paths[file_key])

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)