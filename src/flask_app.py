import os
from flask import Flask, flash, request, redirect, url_for, render_template, session
from flask_dropzone import Dropzone
from werkzeug.utils import secure_filename
from .excel_viewer_v2 import excel2html

UPLOAD_FOLDER = 'C:/Temp/uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)
ALLOWED_EXTENSIONS = set(['csv', 'xls', 'xlsx', 'xlsm', 'xlsb'])

app = Flask(__name__)
dropzone = Dropzone(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = 'supersecretkeygoeshere'

# Dropzone settings
app.config['DROPZONE_UPLOAD_MULTIPLE'] = True
app.config['DROPZONE_MAX_FILE_SIZE'] = 20
# app.config['DROPZONE_ALLOWED_FILE_CUSTOM'] = True
# app.config['DROPZONE_ALLOWED_FILE_TYPE'] = 'image/*'
app.config['DROPZONE_REDIRECT_VIEW'] = 'results'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    session['file_urls'] = []
    if request.method == 'POST':
        file_obj = request.files
        for f in file_obj:
            file = request.files.get(f)
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                session['file_urls'].append(filename)
        return 'Uploading...'
    return render_template('index.html')


@app.route('/uploads/<filenames>')
def uploaded_file(filenames):
    filename = filenames.split(',')[0]
    return excel2html(os.path.join(app.config['UPLOAD_FOLDER'], filename))


@app.route('/results')
def results():
    # print(session['file_urls'][0])
    # print(excel2html(os.path.join(app.config['UPLOAD_FOLDER'], session['file_urls'][0])))
    return render_template('results.html',
                           result_html=excel2html(os.path.join(app.config['UPLOAD_FOLDER'], session['file_urls'][0])))
