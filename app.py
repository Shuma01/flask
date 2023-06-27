from flask import Flask,render_template, request
from flask_sqlalchemy import SQLAlchemy
import os

app = Flask(__name__)


@app.route('/')
def index():
  return render_template('hello.html')



@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
  if request.method == 'POST':
    file = request.files['file']
    file.save(os.path.join('./uploads', file.filename))
    return f'{file.filename}がアップロードされました'
  else:
    return render_template('upload.html')

if __name__ == '__main__':
  app.run(debug=True)
