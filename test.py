import openpyxl


#置換する文字を返す関数
def text_to_array(cell):
    _str=cell.value
    _str=_str.replace(' ', '')
    arrays=_str.split(',')
    return arrays
# 特定の列を複数検索
def search_columns(column, keyword):
    result = []
    for cell in column:
        # セルのデータを文字列に変換
        try:
            value = str(cell.value)
        # 文字列に変換できないデータはスキップ
        except:
            continue
        # キーワードに一致するセルの番地を取得
        if keyword in value:
            cell_address = openpyxl.utils.get_column_letter(cell.column) +  str(cell.row)
            result.append(cell_address)
            
    return result



wb = openpyxl.load_workbook("uploads/三者面談（回答）.xlsx")
ws = wb.worksheets[0]

#先生のセルの横軸取得
teacher_cell=ws["C2"]
table_row=text_to_array(ws["C2"])
replace_after=[i for i in range(len(table_row))]
print(replace_after)

print(table_row)
#不要行の削除
ws.delete_cols(3)
ws.delete_cols(1)
ws.delete_rows(2)

change_cell=ws["D2"]
check_row=ws["D1"]
check_column=ws["A2"]
#ここで置換を行う
i=0
while(True):
    j=0
    while(True):
        #1行目の端まで見たらループを抜ける
        if check_row.offset(0,j).value is None:
            break
        #日程の置換を行う
        for index,item in enumerate(table_row):
            if change_cell.offset(i,j).value is None:
                change_cell.offset(i,j).value=str(-1)
                break
            else:
                change_cell.offset(i,j).value=change_cell.offset(i,j).value.replace(item, str(index))
        j+=1
    i+=1
    #A列の端まで見たらループを抜ける
    if check_column.offset(i,0).value is None:
        break

change_cell=ws["C2"]
i=0
while(True):
    if change_cell.offset(i,0).value is None:
        break
    if change_cell.offset(i,0).value=="いる":
        change_cell.offset(i,0).value = 1
    else:
        change_cell.offset(i,0).value=0
    i+=1




#新しいテーブルの横軸作成
new_ws= wb.create_sheet(index=0, title="makedata")
new_ws["A1"].value=ws["A1"].value
new_ws["B1"].value=ws["B1"].value
new_ws["C1"].value=ws["C1"].value

table_column=[]
j=0
day_cells_num=search_columns(ws["1"], "出席可能")
for cell_num in day_cells_num:
    for replace_item in range(len(table_row)):
        _column=ws[cell_num].value.replace('出席可能日 [', '').replace("]", "")
        ws[cell_num].value=ws[cell_num].value.replace('出席可能日 [', '').replace("]", "")
        replace_column=_column+","+str(replace_item)
        new_ws["D1"].offset(0,j).value=replace_column
        j+=1
    table_column.append(str(_column))
print(table_column)




i=0
#データの転記
while True:
    if ws["A2"].offset(i,0).value is None:
        break
    for j in range(3):
        new_ws["A2"].offset(i,j).value=ws["A2"].offset(i,j).value
    i+=1


#できた表に01を入力していく。また、個人の合計も加えていく
i=0
while True:
    if ws["A2"].offset(i,0).value is None:
        break
    j=0
    sum=0
    while True:
        if ws["D2"].offset(i,j).value is None:
            break
        day_list=text_to_array(ws["D2"].offset(i,j))
        day_list= [int(i) for i in day_list]
        for k,value in enumerate(replace_after):
            if int(value) in day_list:
                new_ws["D2"].offset(i,k+j*len(replace_after)).value=1
                sum+=1
            else:
                new_ws["D2"].offset(i,k+j*len(replace_after)).value=0
        j+=1
        new_ws["D2"].offset(i,j*len(replace_after)).value=sum
    i+=1

#日付の合計を記入していく
j=0
while True:
    sum=0
    i=0
    if new_ws["D1"].offset(0,j).value is None:
        break
    while True:
        if new_ws["D2"].offset(i,j).value is None:
            break
        sum+=new_ws["D2"].offset(i,j).value
        i+=1
    new_ws["D2"].offset(i,j).value=sum
    j+=1
ws.title="replace"

wb.save('uploads/replace.xlsx')

wb.close()