from flask import Flask,render_template, request ,redirect,url_for,flash,send_file
import data_cleansing as dc
import choice
import shutil

import os

app = Flask(__name__)



@app.route('/', methods=['GET', 'POST'])
def index():
  if os.path.isfile("./uploads/replace.xlsx"):
    os.remove("./uploads/replace.xlsx")
  if os.path.isfile("./uploads/print.xlsx"):
    os.remove("./uploads/print.xlsx")
  global table_column,table_row
  if request.method == 'POST' :
    file = request.files['file']
    print(file)
    if ".xlsx" in str(file.filename):
      file.save(os.path.join('./uploads', "replace.xlsx"))
      table_row,table_column=dc.cleansing("./uploads/replace.xlsx")
      choice.make_excel(table_row,table_column)
      return redirect(url_for('makedata'))
    else:
      flash('無効なファイルがアップロードされました。')
      return render_template('index.html')
    
  else:
    return render_template('index.html')

@app.route('/how_to_use', methods=['GET'])
def how_to_use():
  return render_template('how_to_use.html')

@app.route('/sample', methods=['GET','POST'])
def sample():
    global table_column,table_row
    shutil.copyfile("./static/sample.xlsx", "./uploads/replace.xlsx")
    table_row,table_column=dc.cleansing("./uploads/replace.xlsx")
    print(table_column)
    choice.make_excel(table_row,table_column)
    return redirect(url_for('makedata'))

@app.route('/makedata', methods=['GET','POST'])
def makedata():
  if not os.path.isfile("./uploads/replace.xlsx"):
      return redirect('page_not_found')
  time_id = None
  choice.make_df_file()
  df=choice.make_df()
  df=choice.add_total_df(df)
  print(df)
  if request.method == 'GET':
    if not os.path.isfile("./uploads/print.xlsx"):
      choice.make_excel(table_row,table_column)
  else:
    time_id = request.form.get('item').split(",")
    print(time_id)
    choice.write_data(time_id,df,table_row,table_column)
    df=choice.delete_data(time_id,df,table_row,table_column)
    print(df)
    df.to_excel('uploads/replace.xlsx', sheet_name='makedata')
  if df.empty:
    return(redirect('download'))
  if choice.is_brother(df):
    enable_nums=choice.enable_day(df)
    enable_days=[choice.num_to_day(num,table_row,table_column) for num in enable_nums]
    name=df.index[0][1]
  else:
    enable_nums=choice.enable_day(df)
    enable_days=[choice.num_to_day(num,table_row,table_column) for num in enable_nums]
    name=df.index[0][1]
  rows=choice.print_table(table_row,table_column)
  return render_template('makedata.html',enable_days=enable_days,name=name,rows=rows,columns=table_column,time_id=time_id)

@app.route('/download',methods=['GET','POST'])
def download():
  if request.method == 'GET':
    return(render_template('download.html'))
  else:
    return send_file("./uploads/print.xlsx",as_attachment=True)
  
@app.errorhandler(Exception)
def page_not_found(error):
  return render_template('page_not_found.html'), 404

if __name__ == '__main__':
  app.run(debug=True)
