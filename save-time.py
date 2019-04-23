# pip install requests
# pip install xlwt

# encoding:utf-8
import requests
import json
import xlwt
import datetime
import time
import calendar

week_day_dict = {
    0 : '星期一',
    1 : '星期二',
    2 : '星期三',
    3 : '星期四',
    4 : '星期五',
    5 : '星期六',
    6 : '星期天',
  }
# 表头
xls_header = ['姓名', '日期', '星期', '最早打卡', '最晚打卡', '加班时长', '请假时长', '补卡（早补或晚补', '备注（入职、外出打卡等）']

# 从文件读取用户名密码
def getParams():
    f = open('params.json', 'r')
    str = f.read()
    f.close()
    return json.loads(str)

# 保存的文件
def getFile(username):
    timestr = datetime.datetime.today().strftime("%Y-%m")
    return f'{username}-{timestr}.xls'

xls_file = 'default.xls'

# 日期字符串转星期
def getWeekDay(dateStr):
    date = datetime.datetime.strptime(dateStr, "%Y-%m-%d")
    return week_day_dict[date.weekday()]

def getCurrentMonth():
    d = datetime.datetime.today()
    return d.month

def getMonth(dateStr):
    fmt = '%Y-%m-%d'
    d = datetime.datetime.strptime(dateStr, fmt)
    return d.month

def login():
    # login
    url = 'https://sim-pv.saicmotor.com/login/local'
    headers = {'Content-Type': 'application/json;charset=UTF-8', 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36'}

    try:
      params = getParams()
    except Exception:
      print('请在 params.json 文件设置用户名和密码！！！')
      return

    payload = {'account':params['account'],'password':params['password']}

    r = requests.post(url, headers=headers, json=payload)
    cookies = r.cookies
    resp = r.json()

    # print(r.json())

    # get data
    if '000000' != resp['code'] :
        print('Login fail !!!!!!')
        return
    global xls_file
    xls_file = getFile(resp['data']['name'])
    
    headers['latToken'] = resp['data']['accessToken']
    url = 'https://sim-pv.saicmotor.com/checkStatisticsDay/listByDaySupplierUser'
    # 已上班天数
    size = datetime.datetime.today().day
    payload = {'page': {'curr':'1','size': size}}
    r = requests.post(url, headers=headers, cookies=cookies, json=payload)

    # print(r.json())

    resp = r.json()
    if '000000' == resp['code'] :
        processData(resp)
        print('Success !')

def writeRow(ws, row, rowData):
    for c in range(len(rowData)):
        ws.write(row, c, rowData[c])

# 生成 xls
def genXls(list):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Sheet')

    writeRow(ws, 0, xls_header)

    month = getCurrentMonth()
    r = 1
    for index in range(len(list) - 1, 0, -1):
        d = list[index]
        rowData = []
    
        date = d['checkFirst'].split(' ')[0]
        weekday = getWeekDay(date)
        start = d['checkFirst'].split(' ')[1]
        end = d['checkLast'].split(' ')[1]

        m = getMonth(date)
        
        if month != m:
          continue

        rowData.append(d['supplierUserName'])
        rowData.append(date)
        rowData.append(weekday)
        rowData.append(start)
        rowData.append(end)

        writeRow(ws, r, rowData)

        r += 1

    wb.save(xls_file)

# test json
def getjson():
    f = open('time.json', 'r')
    str = f.read()
    f.close()
    return json.loads(str)

def processData(resp):
    try:
        list = resp['data']['rows']
        genXls(list)
    except Exception:
        print("Load data fail!")

# processData(getjson)

login()

