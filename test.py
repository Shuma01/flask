import openpyxl

wb = openpyxl.load_workbook("三者面談（回答）.xlsx")
ws = wb.worksheets[0]


# 特定の行を検索
def search_row(row, keyword):
    for cell in row:
        # セルのデータを文字列に変換
        try:
            value = str(cell.value)
        # 文字列に変換できないデータはスキップ
        except:
            continue
        # キーワードに一致するセルの番地を取得
        if value == keyword:
            result = openpyxl.utils.get_column_letter(cell.column) +  str(cell.row)

    return result


# 特定の列を検索
def search_column(column, keyword):
    for cell in column:
        # セルのデータを文字列に変換
        try:
            value = str(cell.value)
        # 文字列に変換できないデータはスキップ
        except:
            continue
        # キーワードに一致するセルの番地を取得
        if value == keyword:
            result = openpyxl.utils.get_column_letter(cell.column) +  str(cell.row)

    return result

#先生と書かれたセルの取得
teacher_cell=search_row(ws['B'], "先生")


#置換する文字を返す関数
def generateReplace(cell):
    _time_str=ws[cell].offset(0,1).value
    _time_str=_time_str.replace(' ', '')
    time_arrays=_time_str.split(',')
    return time_arrays


replace_items=generateReplace(teacher_cell)
replace_after=[i for i in range(len(replace_items))]
print(replace_after)

print(replace_items)

#文字の置換をおこなう
cell=ws["F3"]
check_row=ws["F1"]
check_column=ws["A3"]
while(True):
    i=0
    while(True):
        #1行目の端まで見たらループを抜ける
        if check_row.offset(0,i).value is None:
            break
        #文字の置換を行う
        for index,item in enumerate(replace_items):
            if cell.offset(0,i).value is None:
                cell.offset(0,i).value=int(-1)
                break
            else:
                cell.offset(0,i).value=cell.offset(0,i).value.replace(item, str(index))
        i+=1
    #A列の端まで見たらループを抜ける
    if check_column.offset(1,0).value is None:
        break
    #次の行に移動する
    cell=cell.offset(1,0)
    check_column=check_column.offset(1,0)




chande_cell=ws["E3"]
while(True):
    print("a")
    if  chande_cell.offset(0,i).value is None:
        break
    if chande_cell.offset(0,i).value=="いる":
        chande_cell.offset(0,i).value=int(1)
    else:
        chande_cell.offset(0,i).value=int(0)
    i+=1




ws.title="replace"
wb.save('replace.xlsx')

wb.close()