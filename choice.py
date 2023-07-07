import pandas as pd
import data_cleansing as dc
import openpyxl
table_row,table_column=dc.cleansing("三者面談（回答）.xlsx")

#dfファイルの作成
def make_df_file():
    df=pd.read_excel('uploads/replace.xlsx', index_col=[0,1,2])
    df=df.fillna(0)
    df.to_excel('uploads/print.xlsx', sheet_name='makedata')

#最初にdfの作成
def make_df():
    df=pd.read_excel('uploads/print.xlsx', index_col=[0,1,2])
    return df


#合計数の少ない人でほかの人のかぶりが少ない日の数を返す
def total_min(df):
    df = df.loc[:, df.iloc[0] != 0]
    min_day=df.iloc[-1].idxmin()
    return min_day

#集計
def add_total_df(df):
    if 'Total' in df.columns :
        df.drop('Total', axis=0)
        df.drop('Total', axis=1)
    df["Total"] = df.sum(axis=1)
    df=df.sort_values(by=["兄弟","Total"],ascending=[False,True])
    df.loc["Total"] = df.sum()
    return df

#出席可能な日を求める(番号)
def enable_day(df):
    df = df.loc[:, df.iloc[0] != 0]
    return df.columns.to_list()[:-1]

#番号から日を求める
def num_to_day(day_num):
    q, mod = divmod(day_num, len(table_row))
    if mod is None:
        mod=0
    return[table_column[q],table_row[mod-1]]

#日から番号を求める
def day_to_num(day_data):
    q=table_column.index(day_data[0])
    mod=table_row.index(day_data[1])
    return q*len(table_row)+mod

def make_excel():
    wb = openpyxl.load_workbook("uploads/replace.xlsx")
    ws= wb.create_sheet(index=2, title="print")
    for i,value in enumerate(table_row):
        ws.cell(row=2+i,column=1).value=value
    for i,value in enumerate(table_column):
        ws.cell(row=1,column=2+i).value=value
    wb.save('uploads/replace.xlsx')
    wb.close()

def write_data(day_data):
    wb = openpyxl.load_workbook("uploads/replace.xlsx")
    ws= wb["print"]
    column=table_column.index(day_data[0])+2
    row=table_row.index(day_data[1])+2
    ws.cell(row=row,column=column).value=1
    wb.save("uploads/replace.xlsx")
    wb.close()


def delete_data(day_data,df):
    del_num=day_to_num(day_data)
    df=df.drop(del_num,axis=1)
    df=df.drop(df.index[[0, -1]])
    df=df.drop(columns=df.columns[[-1]])
    return df

#兄弟がいるか調べる
def is_brother(df):
    if int(df.index[0][2])==1:
        return True
    else:
        return False
    

df=make_df()
# df=add_total_df(df)
# df=delete_data(num_to_day(enable_day(df))[1],df)
print(df)
print(df.index[0])