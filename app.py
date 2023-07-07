from flask import Flask,render_template, request ,redirect,url_for,flash
from flask_sqlalchemy import SQLAlchemy
import data_cleansing as dc
import choice

import os

app = Flask(__name__)


app.secret_key = b'icecream2022'
@app.route('/', methods=['GET', 'POST'])
def upload_file():
  global table_column,table_row
  if request.method == 'POST' :
    file = request.files['file']
    if ".xlsx" in str(file.filename):
      file.save(os.path.join('./uploads', file.filename))
      table_row,table_column=dc.cleansing(file.filename)
      choice.make_excel()
      return redirect(url_for('makedata'))
    else:
      flash('無効なファイルがアップロードされました。')
      return render_template('index.html')
    
  else:
    return render_template('index.html')

@app.route('/makedata', methods=['GET', 'POST'])
def makedata():
  if os.path.isfile("./uploads/print.xlsx"):
    choice.make_df_file()
  df=choice.make_df()
  df=choice.add_total_df(df)
  if choice.is_brother(df):
    enable_nums=choice.enable_day(df)
    enable_days=[choice.num_to_day(num) for num in enable_nums]
    name=df.index[0][1]
  else:
    choose_num=choice.total_min(df)
    choose_day=choice.num_to_day(choose_num)
    choice.write_data(choose_day)



  return render_template('makedata.html',enable_days=enable_days,name=name)
  



if __name__ == '__main__':
  app.run(debug=True)
