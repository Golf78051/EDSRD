from flask import Flask, render_template, request, send_from_directory
import os
from werkzeug.utils import secure_filename
from logic import process_exp_and_trace

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'static/processed_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['expfile']
        depth = int(request.form['depth'])
        source_nets = request.form['source_nets']
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            summary_path, trace_path = process_exp_and_trace(filepath, depth, source_nets, app.config['RESULT_FOLDER'])
            return render_template('index.html',
                                   summary_file=os.path.basename(summary_path),
                                   trace_file=os.path.basename(trace_path))
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
