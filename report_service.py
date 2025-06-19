from flask import Flask, request, send_file
from flask_cors import CORS
import report_handler as rh

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file:
        
        # Save the Word document
        word_file_path = rh.generate_report(file)
        
        # Send the Word document as a response
        return send_file(word_file_path, as_attachment=True)
    
@app.route('/reinforce', methods=['POST'])
def reinforce():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    if file:
        contents = request.form.get('contents')
        # Save the Word document
        word_file_path = rh.reinforce_report(file, contents)
        
        # Send the Word document as a response
        return send_file(word_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=8081)
