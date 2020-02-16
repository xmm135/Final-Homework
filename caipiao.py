#一、导入爬虫第三方库
import requests
from bs4 import BeautifulSoup
import time
import xlwt
#二、分析网页爬取数据
#构造浏览器模拟（应对403禁止访问）
headers = {'User-Agent':'Mozilla/5.0(Wimdows NT 6.1; WOW64) AppleWebkit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'}
#构造网页结构
all_lists = []
for k in range(1,74):
    url = 'http://tubiao.17mcp.com/Qxc/Haoma.aspx?page={}'.format(str(k))
    #请求HTML
    res = requests.get(url, headers=headers,timeout=30)
    #对HTML进行解析
    soup = BeautifulSoup(res.text, 'lxml')

    for item in soup.select('tr')[1:-1]:
        dict = {}
        #先寻找标识tr
        #try异常错误机制
        try:
            #寻找tr下面的标识
            qihao = item.select('td strong')[0].text
            onewei = item.select('td div')[0].text
            twowei = item.select('td div')[1].text
            threewei = item.select('td div')[2].text
            fourwei = item.select('td div')[3].text
            fivewei = item.select('td div')[4].text
            sixwei = item.select('td div')[5].text
            sevenwei = item.select('td div')[6].text
            list = [qihao,onewei,twowei,threewei,fourwei,fivewei,sixwei,sevenwei]
            all_lists.append(list)
            print(list)
        except IndexError:
            pass
#三、把数据存入Excel中
row0 = ['期数','个位','十位','千位','万位','第5位','第6位','第7位']#定义表头，即Excel中第一行标题
book = xlwt.Workbook(encoding='utf-8')#创建工作簿
sheet = book.add_sheet('七星彩',cell_overwrite_ok=True)
#创建表名,cell_overwrite_ok=True很重要，用于确认同一个cell单元是否可以重设值
for row in range(len(row0)):
    sheet.write(0,row,row0[row])#写入表头
i = 1 #第一行开始
for all_list in all_lists:
    j = 0#迭代行
    for data in all_list:
        sheet.write(i,j,data)#迭代列，并写入数据，#重新设置，需要cell_overwrite_ok=True
        j += 1
    i += 1
book.save('qixingcai.xls')