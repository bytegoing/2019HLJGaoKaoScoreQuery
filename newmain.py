import urllib.request
import urllib.error
import http.cookiejar
from PIL import ImageTk
from PIL import Image as Image
from tkinter import *
import PIL
import tkinter as tk
import io
import time
from bs4 import BeautifulSoup
import xlrd
import openpyxl

tobeCheckedCodeIMG = Image
finalResultRes = ""

# 验证码窗口相关代码
class GetCode(object):
    def __init__(self):
        global tobeCheckedCodeIMG
        global nowCode
        print("正在输入验证码...")
        self.data={}  # 存放返回值
        self.root = tk.Tk()
        self.root.geometry('135x100')
        self.root.resizable(width=False,height=False)   # 固定长宽不可拉伸

        self.textLabel=tk.Label(self.root,text="输入验证码后按下回车").pack()  # 标签
        self.textStr=StringVar()
        self.textEntry=tk.Entry(self.root,textvariable=self.textStr)  # 创建输入框
        self.textStr.set("")  # 清空输入框
        self.textEntry.pack()  # 输入框
        self.textEntry.bind('<Return>', self.return_code)  # 回车键按下自动递交
        self.textEntry.focus_set()  # 设置焦点

        im=PIL.Image.open(io.BytesIO(tobeCheckedCodeIMG))
        img=ImageTk.PhotoImage(im)
        tk.Label(self.root,image=img).pack() # 显示图片

        self.root.protocol('WM_DELETE_WINDOW', doNothing)  # 禁止关闭
        self.root.mainloop()

    def return_code(self, x):
        global nowCode
        # 返回输入框内容
        self.data["code"]=self.textStr.get()
        self.root.destroy()           # 关闭窗体
        nowCode = self.data["code"]


def doNothing():
        return
# 验证码窗口结束

# studentListExcel(studentList.xlsx)为学生信息表.该表结构如下：（第一行就应该开始放学生信息，不要有题头）
# |姓名|身份证号|准考证号|
studentListExcel = 'D:/studentList.xlsx'
# finalResult.xlsx为最终结果存储表，该表结构如下：(会自动生成题头)
# |准考证号|姓名|语文|数学|外语|综合|总分|听力|
finalResultExcel = 'D:/finalResult.xlsx'

# 读取xlsx表中信息并存入列表
sheet = xlrd.open_workbook(studentListExcel)
table = sheet.sheets()[0]  # 打开Sheet1
nrows = table.nrows  # 查询总行数
allStudents = nrows - 1

# 第一名学生插入List中为第0个
studentList = [[0, table.col_values(0)[1], table.col_values(1)[1], table.col_values(1)[1][-6:], table.col_values(2)[1]]]
# 从第二名(List中第1个)开始循环读取到最后一名学生插入List中
i = 2
while i < nrows - 1:
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
    coreUrls = ['http://xxcx.hljea.org.cn/JWWebGaokaoNew/index/']
    from random import choice
    coreUrl = choice(coreUrls)
    mainUrl = coreUrl + 'studentloginCheckScore'
    captchaUrl = coreUrl + 'getVerify'
    headers = {
        "Content-type": "application/x-www-form-urlencoded",
        'Accept-Language': 'zh-CN,zh;q=0.8',
        'User-Agent': "Mozilla/5.0 (Windows NT 6.1; rv:32.0) Gecko/20100101 Firefox/32.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Connection": "close",
        "Cache-Control": "no-cache"
    }
    print("本次使用url: " + mainUrl)
    # 开始爬取验证码
    reqCode = urllib.request.Request(url=captchaUrl, headers=headers)
    cjar = http.cookiejar.CookieJar()
    cookie = urllib.request.HTTPCookieProcessor(cjar)
    opener = urllib.request.build_opener(cookie)
    urllib.request.install_opener(opener)
    print("正在获取验证码...地址: " + captchaUrl)
    getCodeReady = False
    while (not getCodeReady):
        try:
            tobeCheckedCodeIMG = urllib.request.urlopen(reqCode).read()
        except Exception as e:
            print("遇到HTTP错误:" + str(e))
            print("暂停一秒后继续重试")
            time.sleep(1)
        else:
            getCodeReady = True
    GetCode()  # 弹出输入验证码窗口
    print("输入了验证码: "+nowCode)
    # 开始爬取成绩页面
    data = {
        'xm': xm,
        'ksh': ksh,
        'sfzh': sfz,
        'authCode': nowCode,
    }  # POST数据体
    getResultsReady = False
    # 构建请求体
    resultReq = urllib.request.Request(url=mainUrl, data=urllib.parse.urlencode(data).encode('utf-8'), headers=headers)
    while (not getResultsReady):
        try:
            finalResultRes = urllib.request.urlopen(resultReq).read().decode('utf-8')  # 发送请求并保存页面
        except Exception as e:
            print("遇到HTTP错误:" + str(e))
            print("等待一秒")
            time.sleep(1)  # 等待1秒
        else:
            # 开始检查返回页面正确性
            if(str(finalResultRes).find('请重新核实您的信息') != -1):
                # 信息不正确
                print("请重新核实您的信息!")
                finalResultRes = ""
                getResultsReady = True
            else:
                # 信息正确，开始检查验证码
                if(str(finalResultRes).find('验证码错误') == -1):
                    # 验证码正确
                    print("检查通过!")
                    getResultsReady = True
                else:
                    # 验证码错误
                    print("验证码错误!")
                    getResultsReady = False
    # 判断是否信息不正确
    if(str(finalResultRes) == ""):
        # 信息确实不正确，标记上
        rawList = [ksh, xm]
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    else:
        # 信息正确，洗数据吧。
        print("正在提取成绩信息...")
        # 数据放锅里洗出来再打碎
        rawList = BeautifulSoup(str(finalResultRes), features='html.parser').select_one('tr[bgcolor="#F5F5F5"]').get_text().split()
        print(rawList)  # 输出一下
        sheet.append(rawList)  # 追加一行
        print("正在保存...")
        workbook.save(finalResultExcel)  # 记得保存数据
    j = j + 1  # 别忘了自增开始下一名学生的处理。
