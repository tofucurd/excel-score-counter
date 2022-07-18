import os
import xlwings as xw
import matplotlib.pyplot as plt
import numpy as np

def addr(s,a,b,c=2):
    return s[a,b].address.replace('$','',c)

for i in os.listdir('./'):print(i)
print('choose one excel file')
ID=input()

app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False

bk=app.books.open(ID)
sht=bk.sheets[0]
plt.rcParams["font.family"]=["STHeiti"]

def work():
    r=sht.range('a1').expand('down').count
    c=sht.range('a3').expand('right').count

    sht[1,c].value=['排名','成绩','平均分占比']

    print('date of exam')
    date=input()
    sht.range(f'{addr(sht,0,c)}:{addr(sht,0,c+2)}').api.merge()
    sht[0,c].value=date

    bk.save()

    new=[]
    for i in range(2,r):
        print('scores for %s' %sht[i,0].value)
        new.append(input())
    for i in range(2,r):
        sht[i,c+1].value=new[i-2]

    bk.save()

    st=addr(sht,2,c+1,1)
    ed=addr(sht,r,c+1,1)
    for i in range(2,r):
        pos=addr(sht,i,c+1)
        sht[i,c].formula=f'=RANK({pos},{st}:{ed})'

    bk.save()

    c+=2
    st=addr(sht,2,c-1,1)
    ed=addr(sht,r,c-1,1)
    for i in range(2,r):
        pos=addr(sht,i,c-1)
        sht[i,c].formula=f'=ROUND({pos}/AVERAGE({st}:{ed}),4)*100'

    bk.save()

    for i in range(2,r):
        fo=sht[i,1].formula
        ls=list(fo)
        ls.insert(-4, f',{addr(sht,i,c,2)}')
        fo=''.join(ls).replace('(0,','(')
        # print(fo)
        sht[i,1].formula=fo

    bk.save()

    sht.range('a3').expand('table').api.sort(key1=sht.range('b3').api,order1=2)

    bk.save()

    sht.autofit()

    bk.save()

    graph()

def graph():
    r=sht.range('a1').expand('down').count
    c=sht.range('a3').expand('right').count

    sht.range('a3').expand('table').api.sort(key1=sht.range('b3').api,order1=2)

    fig=plt.figure(num=1,figsize=(4*r,18))

    if sht.pictures.count>0:
        sht.pictures[0].delete()

    bk.save()

    for i in range(2,r):
        name=sht[i,0].value
        data=[]
        cnt=0
        for j in range(2,c+1):
            fo=sht[i,j].formula
            if fo.find('/AVERAGE')!=-1:
                cnt+=1
                data.append(sht[i,j].value)
        ax=fig.add_subplot(8,4,i-1)
        ax.set_title(name)
        ax.tick_params(labelbottom=False)
        ax.tick_params('both',direction='in')
        ax.set_ylim([0,300])
        ax.plot(np.arange(0,cnt),data,"c-d")
        ax.plot(np.arange(0,cnt),np.linspace(100,100,cnt))
    if r-2==8:sht.pictures.add(fig,name="picture",update=True,left=sht.range((r+3,1)).left,top=sht.range((r+3,1)).top,width=3000,height=4500//r)
    else:sht.pictures.add(fig,name="picture",update=True,left=sht.range((r+3,1)).left,top=sht.range((r+3,1)).top,width=1000,height=1000//r)
    bk.save()

while True:
    print("add all data/just update graph?(1/2)")
    res=input()
    if res=='1':work()
    else:graph()
    print('do you want to continue adding data?(y/n)')
    res=input()
    if res=='n':break
    
bk.close()
app.quit()