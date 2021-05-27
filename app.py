import os
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
from flask import Flask, jsonify, request, send_file
from excel_file import excel_file

load_dotenv()

UPLOAD_FOLDER = '.\\uploads-excel'
ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'xlsm'}


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
  return '.' in filename and \
          filename.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

if __name__ == '__main__':
  app.run(debug=True, port=4000)

@app.route("/")
def hello_world():
  return "<p>Hello, World</p>"



@app.route('/upload-excel', methods=['POST'])
def upload_excel():
  if request.method == 'POST':
    for file_key in request.files:
      print(f"Filename:  {file_key}")
      file_excel = request.files[file_key]
      if file_excel and allowed_file(file_excel.filename):
        filename = secure_filename(file_excel.filename)
        file_excel.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

      
  
  return jsonify({"data":"recibido","request":request.method})



@app.route('/download-files', methods=['GET'])
def download_files():
  if request.method == 'GET':
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'test_2.xlsx')
    print(f"path: {path}")
    return send_file(path,  as_attachment=True)
