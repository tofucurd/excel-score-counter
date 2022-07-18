import os
import xlwings as xw

def addr(s,a,b,c=2):
    return s[a,b].address.replace('$','',c)

for i in os.listdir('./'):print(i)
print('choose two excel file(one a line,earlier first)')
ID1=input()
ID2=input()

app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

bk1=app.books.open(ID1)
sht1=bk1.sheets[0]

bk2=app.books.open(ID2)
sht2=bk2.sheets[0]

bk=app.books.add()
bk.save('result.xlsx')
sht=bk.sheets[0]
sht.clear()

sht[0,0].value=['姓名',ID1.replace('.xlsx', ''),ID2.replace('.xlsx', ''),'差']
s=sht2[2,0].expand('down').value
cnt=sht2[2,0].expand('down').count

for i in range(cnt):
    s1=0
    for j in range(cnt):
        if sht1[j+2,0].value==s[i]:
            s1=sht1[j+2,1].value
    s2=sht2[i+2,1].value
    sht[i+1,0].value=[s[i],s1,s2,s2-s1]

sht.range('a2').expand('table').api.sort(key1=sht.range('d2').api,order1=2)

sht[0,0].expand('table').api.VerticalAlignment=-4108

bk.save()
bk.close()
app.quit()