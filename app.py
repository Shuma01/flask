from flask import Flask,render_template, request ,redirect,url_for,flash
from flask_sqlalchemy import SQLAlchemy

import os

app = Flask(__name__)

# app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'
# db = SQLAlchemy(app)

# class Post(db.Model):
#     id = db.Column(db.Integer, primary_key=True)
#     name = db.Column(db.String(30), nullable=False)
#     email = db.Column(db.String(100))
#     password= db.Column(db.String(30), nullable=False)

# @app.route('/', methods=['GET', 'POST'])
# def index():
#     if request.method == 'GET':
#         posts = Post.query.all()
#         return render_template('index.html', posts=posts)
#     else:
#         name = request.form.get('name')
#         email = request.form.get('email')
#         password = request.form.get('password')
#         new_post = Post(name=name, email=email, password=password)

#         db.session.add(new_post)
#         db.session.commit()
#         return redirect('/')
app.config["SECRET_KEY"] = "sample1203"
@app.route('/', methods=['GET', 'POST'])
def upload_file():
  if request.method == 'POST' :
    file = request.files['file']
    if ".xlsx" in str(file.filename):
      file.save(os.path.join('./uploads', file.filename))
      return f'{file.filename}がアップロードされました'
    else:
      flash('無効なファイルです')
      return render_template('index.html')
    
  else:
    return render_template('index.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    return render_template('signup.html')



if __name__ == '__main__':
  app.run(debug=True)
