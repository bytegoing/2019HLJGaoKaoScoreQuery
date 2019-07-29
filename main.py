import urllib.request
import urllib.error
import xlrd
import http.cookiejar
from PIL import ImageTk
from tkinter import *
import PIL
import tkinter as tk
import io

tobeCheckedCodeIMG = PIL.Image.open('D:/test.jpg')
nowCode = ""


class GetCode(object):
    def __init__(self):
        global tobeCheckedCodeIMG
        global nowCode
        print("正在输入验证码...")
        self.data={}  # 存放返回值
        self.root = tk.Tk()
        self.root.geometry('108x130')
        self.root.resizable(width=False,height=False)   # 固定长宽不可拉伸

        self.textLabel=tk.Label(self.root,text="请输入验证码：").pack() # 标签
        self.textStr=StringVar()
        self.textEntry=tk.Entry(self.root,textvariable=self.textStr)
        self.textStr.set("")
        self.textEntry.pack()  # 输入框

        im=PIL.Image.open(io.BytesIO(tobeCheckedCodeIMG))
        img=ImageTk.PhotoImage(im)
        imLabel=tk.Label(self.root,image=img).pack() # 显示图片

        self.but = tk.Button(self.root,text="确认",command=self.return_code).pack(fill="x") # 按键
        self.root.mainloop()

    def return_code(self):
        global nowCode
        # 返回输入框内容
        self.data["code"]=self.textStr.get()
        self.root.destroy()           # 关闭窗体
        print("输入了: "+self.data["code"])
        nowCode = self.data["code"]


def spiderStart(xm, sfz, ksh):
    global tobeCheckedCodeIMG
    global nowCode
    global studentList
    # 爬虫内容
    coreUrls = ['http://gk.hljedu.gov.cn/']
    from random import choice
    coreUrl = choice(coreUrls)
    mainUrl = coreUrl + 'index/studentloginCheckScore'
    captchaUrl = coreUrl + 'index/getVerify'
    headers = {
        'User_Agnet': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36'
    }
    print("本次使用url: "+mainUrl)
    reqCode = urllib.request.Request(url=captchaUrl, headers=headers)
    cjar = http.cookiejar.CookieJar()
    cookie = urllib.request.HTTPCookieProcessor(cjar)
    opener = urllib.request.build_opener(cookie)
    urllib.request.install_opener(opener)
    print("正在获取验证码...")
    getCodeReady = False
    getCodeNum = 0
    while(not getCodeReady):
        try:
            tobeCheckedCodeIMG = urllib.request.urlopen(reqCode).read()
        except urllib.error.URLError as e:
            print("遇到HTTP错误:" + e.reason)
            getCodeNum = getCodeNum + 1
        else:
            getCodeReady = True

        if(getCodeNum >= 3):
            print("错误超过三次，跳过")
            return 0

    # 开始输入验证码
    gc = GetCode()
    print("获取到验证码: "+nowCode)
    data = {
        'xm': xm,
        'ksh': ksh,
        'sfzh': sfz,
        'authCode': nowCode,
    }
    print(data)
    getResultsReady = False
    retryTimes = 0
    postdata = urllib.parse.urlencode(data).encode('utf-8')
    print(postdata)
    resultReq = urllib.request.Request(url=mainUrl, data=postdata, headers=headers)
    print(resultReq)
    while(not getResultsReady):
        try:
            finalResultRes = urllib.request.urlopen(resultReq).read().decode('utf-8')
        except urllib.error.HTTPError as e:
            print("遇到HTTP错误:" + e.reason)
            retryTimes = retryTimes + 1
        else:
            getResultsReady = True

        if(retryTimes >= 5):
            print("错误超过五次，跳过")
            return 0

    if(getResultsReady):
        print("正在检查返回页面合法性")
        #if(finalResultRes.find('请重新核实您的信息')):
        #    print("请重新核实您的信息!")
        #    return 0
        #elif(finalResultRes.find('验证码错误')):
        #    print("验证码错误!")
        #    return 0
        #else:
        print("通过检查，正在保存...")
        with open("D:/GaokaoScoreHTML/"+ksh+".html", "wb") as f:
            f.write(bytes(finalResultRes, encoding="utf8"))
        # print(finalResultRes.read().decode('utf-8'))
        return 1

# 读取xlsx表中信息并存入列表
sheet = xlrd.open_workbook('D:/model.xlsx')
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
print("正在准备爬虫系统...")

j = 0
while j < allStudents:
    print("正在处理:")
    print(studentList[j])
    result = spiderStart(xm=studentList[j][1], ksh=str(studentList[j][4]), sfz=str(studentList[j][3]))
    if(result == 1):
        studentList[j][0] = 1
    j = j + 1

failList = ""

j = 0
failCount = 0
while j < allStudents:
    print("正在查找未成功...")
    if(studentList[j][0] != 1):
        failList = failList + '\n' + studentList[j][4]
        print(studentList[j][1]+" 未成功")
    j = j + 1