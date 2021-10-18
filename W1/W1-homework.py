# !_*_coding:utf-8_*_
# @Time    : 2021/10/11 18:34
# @Author  : 'Rosalyn'

#录入成绩
print('''你好，欢迎使用成绩录入系统！
<********************************************************************************************>''')
name = input("请输入学员名称：")
while True:#录入学科成绩，并检查录入是否为数值型
    try:
        chineselesson = input("请输入%s的语文成绩：" % name)
        numchineselesson = float(chineselesson.replace(".",''))
        break   #若输入的正确，则退出，错误执行except下面代码
    except:
        print("您输入的内容不是数值，请重新输入")

while True:
    try:
        mathlesson = input("请输入%s的数学成绩：" % name)
        nummathlesson = float(mathlesson.replace(".",''))
        break
    except:
        print("您输入的内容不是数值，请重新输入")
while True:
    try:
        Englishlesson = input("请输入%s的英语成绩：" % name)
        numEnglishlesson = float(Englishlesson.replace(".",''))
        break
    except:
        print("您输入的内容不是数值，请重新输入")
print('''%s的成绩已经录入完成！
<********************************************************************************************>''' % name)

#判断成绩好坏
while True:
    try:
        alevel = input("请输入优秀分数线（比如：80）：")
        numalevel = float(alevel.replace(".",''))
        break   #若输入的正确，则退出，错误执行except下面代码
    except:
        print("您输入的内容不是数值，请重新输入")
totalscore = numchineselesson + nummathlesson + numEnglishlesson#录入学科进行汇总
averagescore = totalscore / 3 #总分除以学科数量
if averagescore >= numalevel:
    scoreresult = '表现太棒了，给%s点赞！！！' % name
else:
    scoreresult = '还需要再接再厉哦，加油~'


#打印成绩
print('''<********************************************************************************************>
%s在本次考试中的成绩为：
语文：%.2f 分
数学：%.2f 分
英语：%.2f 分
总分为：%.2f 分！
考试表现如何？    <%s>
<********************************************************************************************>
欢迎下次使用，谢谢！
'''%(name,numchineselesson,nummathlesson,numEnglishlesson,totalscore,scoreresult))


import xlrd,xlwt
from xlutils.copy import copy
xls_file = xlrd.open_workbook(r"w1operation-read book.xls")
print('''
<********************************************************************************************>
你好，成功打开表格%s'''%"w1operation-read book")

sheet_name = xls_file.sheet_names()  
sheet3 = xls_file.sheet_by_name('吃完一起躺板板')
print('''
文件内所有的表格名为：%s
其中sheet3的数据内容打印如下：
第一行：%s
第四列：%s
第四行第四列：%s
第四行第三列：%s
第三行第二列：%s
<********************************************************************************************>'''%(sheet_name,sheet3.row_values(0),sheet3.col_values(3),sheet3.cell_value(3,3),sheet3.cell_value(3,2),sheet3.cell_value(2,1)))  # 第4行 第4列
  

'''以下都是附加题
'''
print('''开始操作学习小组名单和需求列表''')
studyteam_file = xlrd.open_workbook(r"w1operation-studyteam.xls")  
new_file = copy(studyteam_file)
# 创建一个工作表
worksheet = new_file.add_sheet('需求统计表')
# 写入数据；（行, 列, 值）
cells = [(0,0),(0,1),(0,2),(0,3)]
values0 ={"第0行": ('实现需求','目前操作','预计时间','完成人'),
          "第1行": ('阿波罗基构云计数据后台导出并进行汇总分析','在Excel里面根据数据透视表完成数据的更新和操作','3小时每周','冉杨')}
row = 0          
for k in values0:
  for i in range(len(values0[k])):
    worksheet.write(row,i,values0[k][i])
  row += 1
  
# 新增数据到表格内
ws0 = new_file.get_sheet(0)
row = 10
studyteampart = ("1部","数据分析学习","冉杨")
for i in range(len(studyteampart)):
    ws0.write(row,i,studyteampart[i])
ws1 = new_file.get_sheet(1)
row = 10
teamage = ("冉杨",18)
for i in range(len(teamage)):
    ws1.write(row,i,teamage[i])
    
new_file.save("学习小组名单和需求统计表.xls")
print('''
数据已全部导入成功，欢迎下次使用，谢谢！
<********************************************************************************************>''')





