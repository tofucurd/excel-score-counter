import xlwings as xw

def addr(s,a,b,c=2):
    return s[a,b].address.replace('$','',c)

app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

bk=app.books.add()
print('your new xlsx name')
name=input()
bk.save(r''+name+'.xlsx')
sht=bk.sheets[0]

sht.clear()

print('number of student')
r=int(input())
print('all student name(one a line)')

sht[0,0].value='N/A'
sht[1,0].value='姓名'
for i in range(2,r+2):
    na=input()
    sht[i,0].value=na

sht[0,1].value='总成绩'
sht[1,1].value='平均占比'
for i in range(2,r+2):
    sht[i,1].formula='=ROUND(AVERAGE(0),2)'

bk.save()
bk.close()
app.quit()
"""
吴秋实
徐哲晨
刘晟林
纪卓然
康旻舟
张仡乐
蒲韦坤
夏知齐

"""