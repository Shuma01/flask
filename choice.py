import pandas as pd
from data_cleansing import table_row,table_column


def make_df():
    df= pd.read_excel('uploads/replace.xlsx', index_col=[0,1,2])
    df=df.fillna(0)
    df["Total"] = df.sum(axis=1)
    df=df.sort_values(by=["兄弟","Total"],ascending=[False,True])
    df.loc["Total"] = df.sum()
    return df

def min_total(df):
    df = df.loc[:, df.iloc[0] != 0]
    min_day=df.iloc[-1].idxmin()
    return min_day

def remake_df(df):
    df.drop('Total', axis=0)
    df.drop('Total', axis=1)
    df["Total"] = df.sum(axis=1)
    df=df.sort_values(by=["兄弟","Total"],ascending=[False,True])
    df.loc["Total"] = df.sum()
    return df

def enable_day(df):
    df = df.loc[:, df.iloc[0] != 0]
    return df.columns.to_list()[:-1]

def change_num_to_day(arr):
    day_list=[]
    for i in arr:
        #書き込む行と列を指定
        q, mod = divmod(i, len(table_row))
        if mod is None:
            mod=0
        day_list.append([table_column[q],table_row[mod-1]])
    return day_list




df=make_df()
is_brother=df.index[0][2]
print(df)
print(change_num_to_day(enable_day(df)))