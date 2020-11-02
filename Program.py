import random
import xlrd
print('*****★☆★细胞生物学题库自主测试程序★☆★*****')
print('************************************************')
print('************★☆★制作：浅い冬★☆★************')
print('************************************************')
print('****************★☆★说明★☆★****************')
print('本程序所有题目均从“细胞生物学题库——集合版.xls”文件中导入，请保证该文件与本程序源文件在同一文件夹内。题型均为单选题，且每一小题有四到五个选项。在题目显示出来后，你需要输入你认为正确答案的标号（区分大小写），并且按下回车键。')
input('【按下回车开始运行】')
book = xlrd.open_workbook('细胞生物学题库——集合版.xls')
table = book.sheet_by_name('题目内容')
print('************************************************')
while True:
    try:
        count = eval(input('请输入你需要检测的题目数量：'))
        break
    except:
        print('【错误】请输入数字')
count_right = 0
count_wrong = 0
for i in range(count):
    j =random.randrange(1,1076)
    question = table.cell(j,1).value
    A = table.cell(j,2).value
    B = table.cell(j,3).value
    C = table.cell(j,4).value
    D = table.cell(j,5).value
    E = table.cell(j,6).value
    answer = table.cell(j,8).value
    print('************************************************')
    print(i + 1,'、',question,sep = '')
    print(' 【A】-',A,sep = '')
    print(' 【B】-',B,sep = '')
    print(' 【C】-',C,sep = '')
    print(' 【D】-',D,sep = '')
    if E != '':
        print(' 【E】-',E,sep = '')
    your_answer = input('▲请输入你的答案：')
    if your_answer == answer:
        print('【答案正确】')
        count_right += 1
    elif your_answer != answer:
        print('【答案错误】')
        print('正确答案为：【',answer,'】',sep = '')
        count_wrong += 1
print('************************************************')
print('********************答题结束********************')
print('************************************************')
accuracy = count_right / count
percentage = round(100 * accuracy,2)
print('总题数  ：',count,sep = '')
print('正确题数：',count_right,sep = '')
print('错误题数：',count_wrong,sep = '')
print('正确率  ：',percentage,'%',sep = '')
print('************************************************')
while True:
    input()
#Date last updated:11/2/2020
