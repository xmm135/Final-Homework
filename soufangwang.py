#导入需要用到的模块
import requests
import pymysql
import time
from bs4 import BeautifulSoup
import tkinter as tk
import xlwt
import importlib,sys
importlib.reload(sys)
from PIL import Image,ImageTk
#背景图片
def resize( w_box, h_box, pil_image):
    """调整图片大小,适应窗体大小"""
    """arg:: w_box:new width h_box:new height pil_image:img"""
    w, h = pil_image.size #获取图像的原始大小
    f1 = 1.0*w_box/w
    f2 = 1.0*h_box/h
    factor = min([f1, f2])
    width = int(w*factor)
    height = int(h*factor)
    return pil_image.resize((width, height), Image.ANTIALIAS)

#获取url下的页面内容，返回soup对象
def get_page(url):
    responce = requests.get(url)
    soup = BeautifulSoup(responce.text,'html.parser')
    return soup
#封装成函数，作用是获取列表页下面的所有租房页面的链接，返回一个连接列表
def get_links(link_url):
    soup = get_page(link_url)
    links_div = soup.find_all('div',class_="pic-panel")
    links = [div.a.get('href') for div in links_div]
    return links
def get_house_info(house_url):
    soup = get_page(house_url)
    price = soup.find('span',class_='total').text #价格
    unit = soup.find('span',class_='unit').text.strip() #单位 strip()函数去空格
    house_info = soup.find_all('p')
    area = house_info[0].text[3:] #面积
    layout = house_info[1].text[5:] #户型
    floor = house_info[2].text[3:] #楼层
    towards = house_info[3].text[5:] #朝向
    subway = house_info[4].text[3:] #地铁
    uptown = house_info[5].text[3:-8].strip() #小区
    location = house_info[6].text[3:] #位置
    info ={
        '价格':price,
        '单位':unit,
        '面积':area,
        '户型':layout,
        '楼层':floor,
        '朝向':towards,
        '地铁':subway,
        '小区':uptown,
        '位置':location
    }
    return info
DATABASE = {
    'host':'localhost',#如果是远程数据库，此处为远程服务器的ip地址
    'database':'examination',
    'user':'root',
    'password':'123456',
    'charset':'utf8mb4'
}
def get_db(setting):
    return pymysql.connect(**setting)
def insert(db,house):
    table_name=cityEntry.get()+'_'+localEntry.get()
    values = "'{}',"* 8 +"'{}'"
    sql_values = values.format(house['价格'],house['单位'],house['面积'],house['户型'],
                               house['楼层'],house['朝向'],house['地铁'],house['小区'],
                              house['位置'])
    
    sql = """
        insert into {0}(price,unit,area,layout,floor,towards,subway,uptown,location)
        values({1})
    """.format(table_name,sql_values)
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()
def creatTable(db):
    table_name=cityEntry.get()+'_'+localEntry.get()
    sql = """
        CREATE TABLE `{}` (
       `price` varchar(80) DEFAULT NULL,
       `unit` varchar(80) DEFAULT NULL,
       `area` varchar(80) DEFAULT NULL,
       `layout` varchar(80) DEFAULT NULL,
       `floor` varchar(80) DEFAULT NULL,
       `towards` varchar(80) DEFAULT NULL,
       `subway` varchar(80) DEFAULT NULL,
       `uptown` varchar(80) DEFAULT NULL,
       `location` varchar(80) DEFAULT NULL
     );""".format(table_name)
    cursor = db.cursor()
    cursor.execute(sql)
    db.commit()
    
def main():
    db = get_db(DATABASE)
    try:
        creatTable(db)
    except:
        print("数据库已存在")
        pass
    num = int(numberEntry.get())
    for i in range(num):
        links = get_links("https://"+dict_loc['{}'.format(cityEntry.get())]+".lianjia.com/zufang/"+dict_loc['{}'.format(localEntry.get())]+"/pg{}/".format(i))
        for link in links:
            time.sleep(0.1)
            house = get_house_info(link)
            insert(db,house)
    lableInit.config(text="{}市{}区数据获取成功".format(cityEntry.get(),localEntry.get()))
    print('DONE')

def quitw():
    top.destroy()

def export():
    db = get_db(DATABASE)
    cursor = db.cursor()
    table_name=cityEntry.get()+'_'+localEntry.get()
    count = cursor.execute('select * from '+table_name)
    # 重置游标的位置
    cursor.scroll(0,mode='absolute')
    # 搜取所有结果
    results = cursor.fetchall()
    # 获取MYSQL里面的数据字段名称
    fields = cursor.description
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('table_'+table_name,cell_overwrite_ok=True)
    # 写上字段信息
    for field in range(0,len(fields)):
        sheet.write(0,field,fields[field][0])
    # 获取并写入数据段信息
    row = 1
    col = 0
    for row in range(1,len(results)+1):
        for col in range(0,len(fields)):
            sheet.write(row,col,u'%s'%results[row-1][col])
    workbook.save(r'C:\Users\Lenovo\Desktop\{}.xls'.format(table_name))
    lableInit.config(text="共{}条数据导出成功!".format(count))

#构造字典
dict_loc = {
    '北京':'bj',
    '东城':'dongcheng',
    '西城':'xicheng',
    '朝阳':'chaoyang',
    '海淀':'haidian',
    '丰台':'fengtai',
    '上海':'sh',
    '浦东':'pudong',
    '宝山':'baoshan',
    '杭州':'hz',
    '西湖':'xihu',
    '下城':'xiacheng',
    '余杭':'yuhang',
    '富阳':'fuyang',
    '郑州':'zz',
    '金水':'jinshui',
    '中原':'zhongyuan',
    '二七':'erqi',
    '高新':'gaoxin',
    '新郑市':'xinzhengshi',
    '洛阳':'luoyang',
    '嵩县':'songxian',
    '新乡':'xinxiang',
    '牧野':'muye'
    }
if __name__ == "__main__":
    top = tk.Tk()
    top.title("链家")
    top.resizable(width=False,height=False)#设置不可拉伸
    top.geometry("410x510") #设置窗口大小
         
    #添加背景图片
    canvas = tk.Canvas(top) #设置canvas
    pil_image = Image.open('lianjiabg2.png') #打开背景图片
    pil_image_resize = resize(410,510,pil_image) #将它放大保存
    im = ImageTk.PhotoImage(pil_image_resize)
    canvas.create_image(205,255,image = im) #将图片加载到canvas来
    canvas.place(x=0,y=0,height=510,width=410,anchor='nw')#放到屏幕当中
    
    #图片
    photo = tk.PhotoImage(file="F:\CodeWorkspace\lianjia.png")
    imgLabel = tk.Label(top,image=photo,bg='#fbfbfb',width=410)
    imgLabel.grid(row=0,column=0,columnspan=2)
    #lable
    Label = tk.Label(top,fg='#589e6e',bg='#f9f7ba',font = '隶书 -20 ', text = "请输入您要查询的地区")
    Label.grid(row=1,column=0,columnspan=2,pady=5)
    #市
    cityEntry = tk.Entry(top,width=12)
    cityEntry.grid(row=2,column=0,padx=5,pady=10,sticky="E")
    cityLabel = tk.Label(top,fg='#589e6e',font = '隶书 -20 ',bg='#d9f3e1', text = "市")
    cityLabel.grid(row=2,column=1,sticky="W")
    #区
    localEntry = tk.Entry(top,width=12)
    localEntry.grid(row=3,column=0,padx=5,pady=10,sticky="E")
    localLabel = tk.Label(top,fg='#589e6e',font = '隶书 -20 ',bg='#d9f3e1', text = "区")
    localLabel.grid(row=3,column=1,sticky="W")
    #lable2
    Label2 = tk.Label(top,fg='#589e6e',bg='#f9f7ba',font = '隶书 -20 ', text = "请输入您要查询的页数\n(每页30条数据)")
    Label2.grid(row=4,column=0,columnspan=2)
    #信息数
    numberEntry = tk.Entry(top,width=12)
    numberEntry.grid(row=5,column=0,padx=5,pady=10,sticky="E")
    numberLabel = tk.Label(top,fg='#589e6e',font = '隶书 -20 ',bg='#abe1c1', text = "页")
    numberLabel.grid(row=5,column=1,sticky="W")
    #提交
    submit = tk.Button(top,bg='#589e6e',fg='white',width=12,height=1,font = 'Helvetica -15 bold', text="数据获取",command=main)
    submit.grid(row=6,column=0,columnspan=2,padx=3,pady=5)
    #lable3
    Label3 = tk.Label(top,fg='#589e6e',bg='#f9f7ba',font = '隶书 -20 ', text = "将数据导出为Excel格式")
    Label3.grid(row=7,column=0,columnspan=2)
    #导出excel
    export = tk.Button(top,bg='#f9a33f',fg='white',width=12,height=1,font = 'Helvetica -15 bold', text="导出数据",command=export)
    export.grid(row=8,column=0,columnspan=2,padx=3,pady=5)
    #退出
    quitB = tk.Button(top,bg='#ff5757',fg='white',width=12,height=1,font = 'Helvetica -15 bold', text="退出",command=quitw)
    quitB.grid(row=9,column=0,columnspan=2,padx=3,pady=3)
    #反馈
    lableInitTitle = tk.Label(top,font = '正楷 -12',text="* * * 提 示 信 息 * * *",width=40,fg="#f9a33f")
    lableInitTitle.grid(row=10,column=0,columnspan=2,ipady=5)

    lableInit = tk.Label(top,bg='#d9f3e1',font = '正楷 -12 ',text="请在上方输入您要查询的信息",width=40,fg="red")
    lableInit.grid(row=11,column=0,columnspan=2,ipady=5)