import os
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
from flask import Flask, jsonify, request, send_file, make_response
from flask_cors import CORS
from excel_file import clean_files, generate_txt_file_data_extra_hours, PATH_OUTPUT

load_dotenv()

UPLOAD_FOLDER = './uploads-excel'
ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'xlsm'}


app = Flask(__name__)
CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
  return '.' in filename and \
          filename.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

if __name__ == '__main__':
  app.run()

@app.route("/", methods=['GET'])
def hello_world():
  return "<h1>Hello, World</h1>"


@app.route("/health-check", methods=['GET'])
def hello_world():
  return _corsify_actual_response(jsonify({"status":"OK"}))


@app.route('/upload-excel', methods=['POST', 'OPTIONS'])
def upload_excel():
  if request.method == "OPTIONS":  # CORS preflight
      return _build_cors_preflight_response()
  if request.method == 'POST':
    list_files = []
    for file_key in request.files:
      print(f"Filename:  {file_key}")
      file_excel = request.files[file_key]
      if file_excel and allowed_file(file_excel.filename):
        filename = secure_filename(file_excel.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file_excel.save(path)
        list_files.append(path)

    code = generate_txt_file_data_extra_hours(list_files)
    return _corsify_actual_response(jsonify({"data": {"code": code}}))
  else:
    return _corsify_actual_response(jsonify({"data":"ERROR","request":request.method}))



@app.route('/download-files/<id_file>', methods=['GET'])
def download_files(id_file):
  if request.method == 'GET':
    clean_files()
    path = os.path.join(PATH_OUTPUT, f'PLANOS-{id_file}.zip')
    path_final = os.path.join(PATH_OUTPUT, f'PLANOS.zip')
    os.rename(path, path_final)

    print(f"path: {path_final}")
    return send_file(path_final,  as_attachment=True)


def _build_cors_preflight_response():
    response = make_response()
    response.headers.add("Access-Control-Allow-Origin", "*")
    response.headers.add('Access-Control-Allow-Headers', "*")
    response.headers.add('Access-Control-Allow-Methods', "*")
    return response
    
    
def _corsify_actual_response(response):
    response.headers.add("Access-Control-Allow-Origin", "*")
    return response
