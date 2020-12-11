import json
import xlrd
from flask import Flask, request

app = Flask(__name__)


def search(department, graduationYear, name):
    workbook = xlrd.open_workbook(r'Archive.xlsx')
    sheet_name = workbook.sheets()[0]
    rowNum = sheet_name.nrows
    colNum = sheet_name.ncols
    list = []
    for i in range(rowNum):
        rowlist = []
        for j in range(colNum):
            rowlist.append(sheet_name.cell_value(i, j))
        list.append(rowlist)
    for i in range(1, rowNum):
        if list[i][1] == department and int(list[i][2]) == int(graduationYear) and list[i][4] == name:
            return list[i][3], list[i][5], True
    return "None", "None", False


@app.route('/get_msg', methods=['GET'])
def get_msg():
    department = request.args.get('department')
    graduationYear = request.args.get('graduationYear')
    name = request.args.get('name')
    major, location, flag = search(department, graduationYear, name)
    message = '如您的档案所在地显示为学校，请拨打：188-1480-4321。'
    result = {'major': major, 'location': location, 'message': message, 'flag': flag}
    return json.dumps(result, ensure_ascii=False)


if __name__ == '__main__':
    app.run(port=50180, host='0.0.0.0', debug=False)
