# 使用python实现的excel成绩统计器

## 环境要求

- python3

- Excel

- python第三方库：``xlwings`` ``matplotlib`` ``numpy``

## Important

python程序运行时均会打开excel软件，属于正常现象，请不要对打开的excel做任何操作，只需要在程序中操作即可，程序执行完后均会关闭正在运行的excel。

## 使用

使用命令``python3 -u <name>.py``运行，注意把要操作的成绩单与成绩放在同一个目录下

### 新建

使用``init.py``新建成绩单

### 添加

使用``excel.py``向一个已有的成绩单添加成绩（注：必须为``init.py``创建）

使用方法：

程序会事先检索其所在目录，并列出excel文件，输入将要更改的文件的**全名**（即包含后缀名）后，会有两个选项，添加成绩选1即可，2则是更新折线时用。添加完毕后会询问是否继续添加，按需选择即可。

统计方式：

单次测试共统计3项：

- 单次排名

- 单次成绩

- 单次成绩相对于当次测试平均分的占比（为了消除每次考试难易程度以及总分不一样的影响），反应了这场测试中每个人相对于集体的水平

除成绩外均使用excel公式

对于一个成绩单，统计一个总成绩，代表这个成绩单中每个人单次占比的平均，反映了这一些考试中每个人相对于集体的水平

在这个表的最下面，还使用``matplotlib``生成了一个折线图，代表每个人在成绩单中考试占比的变化，其中水平的蓝线代表平均水平，顺序按照表中的顺序从左到右。

注意：倘若成绩有误，只需手动将错误的成绩改过来，再运行``excel.py``并选择``just update graph``即可更新折线。


### 比较

使用``comp.py``对两个**成员相同**的成绩单的总成绩一栏做成绩波动比较，并输出到``result.xlsx``中，常用于两个阶段的比较