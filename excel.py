import pandas as pd


df = pd.read_excel('三者面談（回答）.xlsx', index_col=1)
df=df.iloc[:,1:]


#実行前に置換を行う
def generateReplace():
    df = pd.read_excel('三者面談（回答）.xlsx', index_col=1)
    df=df.iloc[:,1:]
    df=df.replace('(.*) (.*)', r'\1\2', regex=True)
    arrays = df.at['先生', '横軸を入力してください'].split(',')
    arrays=[i.replace(' ', '') for i in arrays]
    return arrays
replace_item=generateReplace()

#置換してNANを0にする
replace_item_change={}
j=1
df=df.fillna(0)
for i in replace_item:
    replace_item_change[j]={i}
    df=df.replace('(.*) (.*)', r'\1\2', regex=True)
    df=df.replace(f'(.*){i}(.*)', f'\\1 {j} \\2', regex=True)
    df=df.replace('(.*) (.*)', r'\1\2', regex=True)
    j+=1
# print(replace_item_change)


print (df.columns.get_loc('出席可能日 [１日目]'))
print (df)