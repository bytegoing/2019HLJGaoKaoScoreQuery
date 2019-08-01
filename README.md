# 2019HLJGaokaoScoreQuery

该程序用来批量查询2019年黑龙江省地区普通高等学校招生全国统一考试成绩。

新版程序(newwechatmain.py)可以绕过验证码机制，而前一版本(newmain.py)需要手动输入验证码。


2019年6月23日首版(main.py+soup.py), 2019年7月30日终版(newmain.py).

2019年7月31日免验证码版本(newwechatmain.py)

### 使用方法(newmain.py/newwechatmain.py)
1. 下载python3并安装该程序所需的全部依赖库。
2. 下载newmain.py/newwechatmain.py
3. 修改newmain.py/newwechatmain.py中 ```studentListExcel``` 与 ```finalResultExcel``` 变量的内容，使其指向的文件有权限读写。
4. 按照程序注释中的说明，创建```StudentList.xlsx```文件。
5. 运行程序，手动输入验证码。
6. 查看finalResultExcel所指向的文件，进行统计。（可能需要进行去小数点操作后才可正常统计。）

BTW: 编程真的可以减少劳动量hhhh

BTW2: 如果遇到connection closed by end server这类的错误换一下header就好。被反爬虫措施限制了而已~

BTW3: 真的不要看原版...写的一团糟糕还需要手动再来soup.....年久失修。想要用或者参考直接用新版就好！！(newmain.py)

自己写个小东西，就酱。不提供技术支持。我用着爽就行。

对了，如果还想要一条龙把录取结果都查了的这里有个小工具
https://github.com/BYTEGOING/2019HLJGaokaoAdmissionQuery