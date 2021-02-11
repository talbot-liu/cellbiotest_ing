##导入random和xlrd库
import random
import xlrd

##说明
print('************************************************')
print('*******★☆★细胞生物学题库自测程序★☆★*******')
print('******************★浅い冬★********************')
print('************************************************')

##导入题库
input('即将导入题库，请确保"细胞生物学题库.xls"文件在本程序根目录下。\n[按下回车继续]')
while True:
    try:
        book = xlrd.open_workbook('细胞生物学题库.xls')
        table = book.sheet_by_name('题目内容')
        break
    except:
        input('【错误】未找到文件，请将"细胞生物学题库.xls"文件放在在本程序根目录下。\n[按下回车重试]')
print('导入成功')
print('************************************************')

##模式选择
print('*****************★模式选择★*******************')
print('************************************************')
print('【A】模拟考试：随机抽取100题(默认选择该模式)')
print('【B】遍历题库：一次性做完所有题目(乱序)')
print('【C】错题重做：将你的错题重做(需要错题标码)')
print('【D】自选题数：自己设定题目数(没啥用)')
mode = input('▲请输入你选择的模式(默认模式为A)：').upper()
print('************************************************')

##模式设定
if mode == 'D':
    print('*************★已选模式：自选题数★*************')
    print('************************************************')
    print('请输入你需测试的题目数量(1-928)：',end = '')
    while True:
        number = eval(input())
        if 1 <= number <= 928:
            break
        else:
            print('超出范围，请重新输入数量(1-928)：',end = '')
    count0 = list(range(1,929))
    random.shuffle(count0)
    count = count0[0:number]
elif mode == 'C':
    print('*************★已选模式：错题重做★*************')
    print('************************************************')
    print('请输入你的错题标码：',end = '')
    while True:
        try:
            count = eval(input())
            number = len(count)
            judge = 0
            for i in count:
                if 928<i or i<1:
                    judge = 1
            if judge == 0:
                break
            else:
                print('错题标码格式错误，请重新输入：',end = '')
        except:
            print('错题标码格式错误，请重新输入：',end = '')
    random.shuffle(count)
elif mode == 'B':
    print('*************★已选模式：遍历题库★*************')
    print('************************************************')
    number = 928
    count = list(range(1,929))
    random.shuffle(count)
else:
    print('*************★已选模式：模拟考试★*************')
    print('************************************************')
    number = 100
    count0 = list(range(1,929))
    random.shuffle(count0)
    count = count0[0:100]

##初始化
j = 0                   ##题号
code = []               ##错题集
random.shuffle(count)   ##乱序
count_right = 0         ##正确题数
count_wrong = 0         ##错误题数

##主程序
print('************************************************')
for i in count:
    j += 1
    question = table.cell(i,0).value
    A = table.cell(i,1).value
    B = table.cell(i,2).value
    C = table.cell(i,3).value
    D = table.cell(i,4).value
    E = table.cell(i,5).value
    answer = table.cell(i,6).value
    print(j,'、',question,sep = '')
    print(' 【A】',A,sep = '')
    print(' 【B】',B,sep = '')
    print(' 【C】',C,sep = '')
    print(' 【D】',D,sep = '')
    if E != '':
        print(' 【E】',E,sep = '')
    your_answer = input('▲请输入你的答案：').upper()
    if your_answer == answer:
        print('【答案正确】')
        count_right += 1
    elif your_answer != answer:
        print('【答案错误】')
        print('正确答案为：【',answer,'】',sep = '')
        count_wrong += 1
        code.append(i)
    print('************************************************')

##结算界面
print('********************答题结束********************')
print('************************************************')
accuracy = count_right / number
percentage = round(100 * accuracy,2)
print('总题数  ：',number,sep = '')
print('正确题数：',count_right,sep = '')
print('错误题数：',count_wrong,sep = '')
print('正确率  ：',percentage,'%',sep = '')
print('************************************************')

##生成错题汇总
input('即将生成"错题汇总.txt"，请注意：如该文件已存在，则将覆盖已有文件，若不需要生成错题汇总文件，请直接点击右上角叉号关闭程序\n[按下回车继续]')
file = open('错题汇总.txt','w',encoding='UTF-8')
file.write('本次错题标码：')
file.write(str(sorted(code)))
file.write('\n\n错题：\n')
for i in sorted(code):
    file.write(str(i))
    file.write('、')
    file.write(table.cell(i,0).value)
    A = table.cell(i,1).value
    B = table.cell(i,2).value
    C = table.cell(i,3).value
    D = table.cell(i,4).value
    E = table.cell(i,5).value
    F = table.cell(i,6).value
    file.write('\nA、')
    file.write(str(A))
    file.write('\nB、')
    file.write(str(B))
    file.write('\nC、')
    file.write(str(C))
    file.write('\nD、')
    file.write(str(D))
    if E != '':
        file.write('\nE、')
        file.write(str(E))
    file.write('\n答案：')
    file.write(str(F))
    file.write('\n\n')
file.close()
input('生成成功\n[按下回车结束程序]')
