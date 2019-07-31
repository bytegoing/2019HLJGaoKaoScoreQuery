import urllib.request
import urllib.error
import time
from bs4 import BeautifulSoup
import xlrd
import openpyxl

finalResultRes = ""

# 爬取页面代码
def getPage(resultReq):
    try:
        finalResultRes = urllib.request.urlopen(resultReq).read().decode('utf-8')  # 发送请求并保存页面
    except Exception as e:
        print("遇到HTTP错误:" + str(e))
        print("等待一秒")
        time.sleep(1)  # 等待1秒
    else:
        # 开始检查返回页面正确性
        if (str(finalResultRes).find('请重新核实您的信息') != -1):
            # 信息不正确
            print("请重新核实您的信息!")
            finalResultRes = ""
        else:
            print("检查通过!")
    return finalResultRes
# 爬取页面代码结束

# studentListExcel(studentList.xlsx)为学生信息表.该表结构如下：（第一行就应该开始放学生信息，不要有题头）
# |姓名|身份证号|准考证号|
studentListExcel = 'D:/StudentList.xlsx'
# finalResult.xlsx为最终结果存储表，该表结构如下：(会自动生成题头)
# |准考证号|姓名|语文|数学|外语|综合|总分|听力|
finalResultExcel = 'D:/finalResult.xlsx'

# 读取xlsx表中信息并存入列表
sheet = xlrd.open_workbook(studentListExcel)
table = sheet.sheets()[0]  # 打开Sheet1
allStudents = table.nrows  # 查询总行数

# 第一名学生插入List中为第0个
print("正在读取第1个学生，姓名:"+table.col_values(0)[0]+" 身份证号:"+table.col_values(1)[0]+" 准考证号:"+table.col_values(2)[0])
studentList = [[0, table.col_values(0)[0], table.col_values(1)[0], table.col_values(1)[0][-6:], table.col_values(2)[0]]]
# 从第二名(List中第1个)开始循环读取到最后一名学生插入List中
i = 1
while i < allStudents:
    print("正在读取第"+str(i+1)+"个学生，姓名:"+table.col_values(0)[i]+" 身份证号:"+table.col_values(1)[i]+" 准考证号:"+table.col_values(2)[i])
    studentList.append([0, table.col_values(0)[i], table.col_values(1)[i], table.col_values(1)[i][-6:], table.col_values(2)[i]])
    i = i + 1

print("本次共导入"+str(allStudents)+"个学生")
print("正在打开存储数据文件...")
workbook = openpyxl.Workbook()  # 打开表
sheet = workbook.active  # 默认的第一张sheet
rawList = []
sheet.append(['准考证号','姓名','语文','数学','外语','综合','总分','听力'])
workbook.save(finalResultExcel)  # 记得保存数据
workbook = openpyxl.load_workbook(finalResultExcel)  # 读取新建的文件
sheet = workbook.active
print("准备完成。")
print("开始爬取。")
j = 0
while j < allStudents:
    # 开始准备...
    xm = studentList[j][1]
    ksh = str(studentList[j][4])
    sfz = str(studentList[j][3])
    print("进度: 第"+str(j+1)+"个，共"+str(allStudents)+"个。正在爬取: 姓名:" + xm + " 准考证号: " + ksh + " 身份证后六位: "+ sfz)
    coreUrls = ['http://gk.lzk.hl.cn/JWWebGaokaoNew/wechat/']
    from random import choice
    coreUrl = choice(coreUrls)
    mainUrl = coreUrl + 'studentloginCheckScore'
    headers = {
        "Content-type": "application/x-www-form-urlencoded",
        'Accept-Language': 'zh-CN,zh;q=0.8',
        'User-Agent': "Mozilla/5.0 (Linux; Android 7.1.1; MI 6 Build/NMF26X; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/57.0.2987.132 MQQBrowser/6.2 TBS/043807 Mobile Safari/537.36 MicroMessenger/6.6.1.1220(0x26060135) NetType/4G Language/zh_CN",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "close",
        "Cache-Control": "no-cache"
    }
    print("本次使用url: " + mainUrl)
    # 开始爬取成绩页面
    # 开始请求
    resultPage = "ERROR"
    while(resultPage == "ERROR"):
        data = {
            'xm': xm,
            'ksh': ksh,
            'sfzh': sfz,
        }  # POST数据体
        # 构建请求
        resultReq = urllib.request.Request(url=mainUrl, data=urllib.parse.urlencode(data).encode('utf-8'), headers=headers)
        resultPage = str(getPage(resultReq))  # 返回结果
    # 判断是否信息不正确
    if(resultPage == ""):
        # 信息确实不正确，标记上
        rawList = [ksh, xm, '', '', '', '', '', '']
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    else:
        # 信息正确，洗数据吧。
        print("正在提取成绩信息...")
        # 数据放锅里洗出来再打碎
        soup = BeautifulSoup(resultPage, features='html.parser')
        tags = soup.find_all('li', class_='w35')
        k = 0
        rawList = ['', '', '', '', '', '', '', '']
        for tag in tags:
            if k == 8:
                break
            rawList[k] = tag.get_text().strip('\r\n').strip('\r').strip('\n').strip()
            k = k + 1
        print(rawList)  # 输出一下
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    print("延时1秒避免被封")
    time.sleep(1)  # 延时一下避免被封
    j = j + 1  # 别忘了自增开始下一名学生的处理。
