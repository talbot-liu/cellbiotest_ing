import random
import xlrd
print('************************************************')
print('*****★☆★细胞生物学题库自主测试程序★☆★*****')
print('************************************************')
book = xlrd.open_workbook('细胞生物学题库——集合版.xls')
table = book.sheet_by_name('题目内容')
count_right = 0
count_wrong = 0
s = set()
count = list(range(1,686))
random.shuffle(count)
j = 0
for i in count:
    j += 1
    question = table.cell(i,0).value
    A = table.cell(i,1).value
    B = table.cell(i,2).value
    C = table.cell(i,3).value
    D = table.cell(i,4).value
    E = table.cell(i,5).value
    answer = table.cell(i,6).value
    print('************************************************')
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
        s.add(i)
print('************************************************')
print('********************答题结束********************')
print('************************************************')
accuracy = count_right / 685
percentage = round(100 * accuracy,2)
print('总题数  ：',685,sep = '')
print('正确题数：',count_right,sep = '')
print('错误题数：',count_wrong,sep = '')
print('正确率  ：',percentage,'%',sep = '')
print('************************************************')
input()
file = open('错题汇总.txt','w',encoding='UTF-8')
file.write(str(s))
file.write('\n')
for i in s:
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
