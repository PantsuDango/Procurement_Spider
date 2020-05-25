import requests
import re
import xlwt
from mypinyin import Pinyin
import time
import random
import traceback
import os


def respone(url):

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36"}
    res = requests.get(url, headers=headers)
    res.encoding = 'gb18030'
    html = res.text
    return html


def re_html(html):
    
    # 厂家名
    company = re.findall(r'data-comname="(.+?)"', html)[0]
    # 产品名称
    name = re.findall(r'target="_blank" title="(.+?)">', html)[0]
    # 联系方式
    tel = ' '.join(list(re.findall(r'data-name="(.+?)" data-gender="(.+?)" data-tel="(.*?)" data-mobile="(.*?)"', html)[0]))
    # 起订量
    minimum_quantity = ' / '.join(re.findall(r'<td> \d+\u002d\d+|<td> ?\u2265\d+', html)).replace('<td>','').replace(' ','') 
    # 价格
    price = ' / '.join(re.findall(r'class="red">.+?</span>|<td>\u9762\u8bae', html)).replace('<td>','').replace('class="red">','').replace('</span>','')
    # 单位
    unit = re.findall(r'<th>\u8ba2\u8d27\u91cf（(.+?)）</th>', html)[0]  
    # 供货总量
    try:
        counts = re.findall(r'\u4f9b\u8d27\u603b\u91cf.+?<td> (.+?)件 </td>', html, re.S)[0]  
    except IndexError:
        counts = '未知'
    # 产地
    area = re.findall(r'\u4ea7&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;\u5730.+?<td>(.+?)</td>|\u4ea7\u5730.+?<td>(.+?)</td>', html, re.S)[0]
    # 发货期
    send_time = re.findall(r'\u53d1&nbsp;\u8d27&nbsp;&nbsp;\u671f</th>.+?<td> (.+?) </td>|\u53d1\u8d27\u671f.+?<td> (.+?) </td>', html, re.S)[0]  
    # 是否有现货
    try:
        have = re.findall(r'\u662f\u5426\u6709\u73b0\u8d27.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        have = '未知'
    # 型号
    try:
        model = re.findall(r'\u578b\u53f7.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        model = '未知'
    # 材质
    try:
        material = re.findall(r'<td>\u6750\u8d28.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        material = '未知'
    # 尺寸
    try:
        size = re.findall(r'<td>\u89c4\u683c.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        size = '未知'
    # 包装
    try:
        pack = re.findall(r'\u5305\u88c5.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        pack = '未知'
    # 产量
    try:
        production = re.findall(r'\u4ea7\u91cf.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        production = '未知'
    # 颜色
    try:
        color = re.findall(r'\u989c\u8272.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        color = '未知'
    # 品牌
    try:
        brand = re.findall(r'<td>\u54c1\u724c.+?<td>(.+?)</td>', html, re.S)[0]  
    except IndexError:
        brand = '未知'

    data = []
    data.append(company)
    data.append(name)
    data.append(tel)
    data.append(minimum_quantity)
    data.append(price)
    data.append(unit)
    data.append(counts)
    data.append(area)
    data.append(send_time)
    data.append(have)
    data.append(model)
    data.append(material)
    data.append(size)
    data.append(pack)
    data.append(production)
    data.append(color)
    data.append(brand)
    return data,name


def write_file(datas,thing):

    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet(thing)

    title = ['行业','子目录','商品链接','厂家名','产品名称','联系方式','起订量','价格（元）','单位','供货总量','产地','发货期','是否有现货','型号','材质','尺寸','包装','产量','颜色','品牌']
    for index,col_name in enumerate(title):
        worksheet.write(0,index,label=col_name)

    line = 1
    for data in datas:
        worksheet.write(line,0,label='休闲箱包')
        worksheet.write(line,1,label=thing)
        col = 2
        for word in data:
            worksheet.write(line,col,label=word)
            col += 1
        line += 1

    workbook.save('%s.xls'%thing)


def image_download(html,name,index,thing):

    try:
        os.mkdir('%s商品图片'%thing)
    except Exception:
        pass
    intab = r'?*/\|".:><\n'
    outtab = "            "
    trantab = str.maketrans(intab, outtab)
    path1 = os.getcwd() + '\\' + '%s商品图片'%thing + '\\' 
    name = name.translate(trantab)
    name = '%d.%s'%(index+1, name)
    os.mkdir(path1+name)
    
    # 图片
    images_url = re.findall(r'''rel="{gallery: 'gal1',smallimage: '(.+?)',largeimage|<img class="imgborderdetails" src="(.+?)"''', html)
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36"}
    path2 = path1 + name + '\\'
    number = 1
    for image_url in images_url:
        if image_url == images_url[0]:
            image_url = image_url[1]
        else:
            image_url = image_url[0]
        image = requests.get(image_url, headers=headers).content
        with open(path2+'%d.png'%number, 'wb') as file:
            file.write(image)
        number += 1


def main():

    thing = input('\n>>> 请输入要爬取的商品类型：')
    counts = int(input('>>> 输入要爬取的商品数：'))
    p = Pinyin()
    thing_pinyin = p.get_pinyin(u"%s"%thing,'')
    pages = counts // 60 + 1

    print('\n>>> 开始获取所有商品的详情页...')
    url_list = []
    for page in range(1,pages+1):
        url = 'https://cn.made-in-china.com/market/%s-%d.html'%(thing_pinyin,page)
        html = respone(url)
        url_list += re.findall(r'<div class="tit js-tit">.+?<a href="(.+?)"',html,re.S)
        sec = random.uniform(3,5)
        sec = round(sec,2)
        time.sleep(sec)
    
    url_list = url_list[:counts]
    print('>>> 详情页获取完毕，共成功获取 %d 条商品详情页url\n'%len(url_list))

    print('>>> 开始爬取数据...预计耗时 %d 秒\n'%(counts*4))
    datas = []
    failure = 0
    for index,url in enumerate(url_list):
        try:
            html = respone(url)
            data,name = re_html(html)
            image_download(html,name,index,thing)
        except:
            print('>>> %d. %s ---> 爬取失败，原因如下：'%(index+1,url))
            traceback.print_exc()
            failure += 1
        else:
            data.insert(0,url)
            datas.append(data)
            print('>>> %d. %s ---> 爬取成功'%(index+1,url))
        
        sec = random.uniform(3,5)
        sec = round(sec,2)
        time.sleep(sec)

    print('\n>>> 全部数据爬取完毕，成功爬取 %d 个商品，失败 %d 个'%(counts-failure,failure))
       
    print('\n>>> 开始将所有爬取到的数据写入文件中...')
    write_file(datas,thing)
    print('>>> 文件写入完毕，保存至 ---> %s.xls\n'%thing)


if __name__ == '__main__':
    
    main()