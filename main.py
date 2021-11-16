import xlwt
from datetime import datetime, date, timedelta
import time
import urllib3
from bs4 import BeautifulSoup
import re
import sys

# 关于样式
style_head = xlwt.XFStyle() # 初始化样式
font = xlwt.Font() # 初始化字体相关
font.name = "微软雅黑"
font.bold = True
font.colour_index = 1 # TODO 必须是数字索

bg = xlwt.Pattern() # 初始背景图案
bg.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
bg.pattern_fore_colour = 19 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray

# 设置字体
style_head.font = font
# 设置背景
style_head.pattern = bg

# 创建一个excel
excel = xlwt.Workbook()
# 添加工作区
sheet = excel.add_sheet("list")
sheet.col(0).width = 1000

# 标题信息
head = ["标题","编辑","来源","链接","日期",'涉及关键词']
for index,value in enumerate(head):
    sheet.write(0,index,value,style_head)


# 初始化链接字典
url = {
    '要闻': "http://chisa.edu.cn/rmtnews1/ssyl/",
    '时评': "http://chisa.edu.cn/rmtnews1/guandian/",
    '海外': "http://chisa.edu.cn/rmtnews1/haiwai/",
    '人才': "http://chisa.edu.cn/rmtnews1/rencai/",
    '综合': "http://chisa.edu.cn/rmtnews1/zonghe/",
    '原创': "http://chisa.edu.cn/rmtycgj/",
    '创业': "http://chisa.edu.cn/rmtnews1/chuangye/",
    '留学': "http://chisa.edu.cn/rmtnews1/chuguo/"
}
# 全局变量 && 初始化参数
KEYWORD = ['习近平', '近平']  # 该元组的条件为&&    读取本地关键词文件
COLUMN = ['title','content']#默认只检索title，新增content或者from

BEGINDATE = (date.today() + timedelta(days = -7)).strftime("%Y%m%d")  #开始日期
ENDDATE = (date.today()).strftime("%Y%m%d")    #结束日期
ROWLINE = 0             #excel行号

# 定义批次写入excel程序：行号自增
def writeInFile(content):
    global ROWLINE
    row = ROWLINE

    
    for item in content:
        if item:
            for val in item:
                global sheet
                row = row + 1
                sheet.write(row, 0, str(val['title']))
                sheet.write(row, 1, str(val['editor']))
                sheet.write(row, 2, str(val['source']))
                sheet.write(row, 3, str(val['href']))
                sheet.write(row, 4, str(val['date']))
                sheet.write(row, 5, str(val['word']))
    ROWLINE = row + 1

# 在参数content中，匹配元组KEYWORD中的每一个元素，如有某个元素匹配不到，则返回False
def matchByKeyword(content, keyword = KEYWORD):
    result = {}
    result["status"] = False
    result["word"] = ''
    i = 0
    for item in keyword:
        if item in content:
            result["status"] = True
            result["word"] = result["word"],'|',item
            i = i + 1
    return result

# 此函数接收文章url，抓取内容并根据条件进行返回
def getContent(fatherUrl, url):
    http = urllib3.PoolManager()

    # 拼写完整url，抓取内容
    url = (fatherUrl+url)
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        print('article link faild')
        sys.exit()
        return None

    # 格式化抓取数据
    content = response.data.decode()
    html = BeautifulSoup(content, features='html.parser')

    # 单独抓取到PC模板的leftpart内容
    #divList = html.find_all("div", {"class", "leftpart"})
    divList = html.find_all("html")
    try:
        # 格式化数据
        item = divList[0]
        item = str(item)
        item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
        item = re.sub('\n', '', item)
        item = re.sub('\s', '', item)
        
        title = re.findall(r'<h1class="content_title">(.*?)</h1>', item)
        title = re.findall(r'<title>(.*?)</title>', item)
        title = ''.join(title)
        
        # title条件筛选
        if 'title' in COLUMN:
            titleIsLegal = matchByKeyword(title)

        # content条件筛选
        if 'content' in COLUMN:
            content = re.findall(r'<divclass="detail"id="js_content">(.*?)</div>', item)
            content = ''.join(content)
            contentIsLegal = matchByKeyword(content)
        
        if titleIsLegal['status']==False and contentIsLegal['status']==False:
            return None

        # 定义result字典，用于返回给调用函数
        result = {}
        result['title'] = title

        result['source'] = re.findall(r'<divclass="from">.*?</script>来源：(.*?)<script>.*?</div>', item)
        result['source'] = ''.join(result['source'])

        result['editor'] = re.findall(r'<pclass="more">责任编辑：(.*?)</p>', item)
        result['editor'] = ''.join(result['editor'])

        result['href'] = url

        articleDate = re.findall(r'.*?t(\d+)_.*?', (url))
        result['date'] = ''.join(articleDate)

        result['word'] = titleIsLegal['word']

        return result

    except BaseException as err:
        print(err)
        print('find content faild')
        sys.exit()
        return None

# 分离每页文章，判断链接日期后，发送链接至getContent函数
def processData(url, row=0, rootUrl=False):

    # 创建http连接池
    http = urllib3.PoolManager()

    # 抓取一级目录列表
    try:
        response = http.request('GET', url)
    except BaseException as err:
        print(err)
        print('url item link faild')
        sys.exit()
        return None
    # 获取状态码，如果是200表示获取成功
    code = response.status

    # 读取内容
    if 200 == code:
        content = response.data.decode()
        html = BeautifulSoup(content, features='html.parser')

        # 将每个文章列表分离为单独的元素
        result = []
        divList = html.find_all("div", {"class", "hnews block nopic"})
        for item in divList:

            # 格式化数据
            item = str(item)
            item = item.replace('\r', '').replace('\r\n', '').replace('\t', '')
            item = re.sub('\n', '', item)
            item = re.sub('\s', '', item)

            # 获取文章链接日期进行比对
            articleDate = re.findall(r'<divclass="txtconthline">.*?<ahref=".*?t(.*?)_.*?.html".*?>.*?</a>.*?</div>', (item))
            articleDate = ''.join(articleDate)
            if articleDate > ENDDATE or articleDate < BEGINDATE:
                continue

            # 匹配获取到文章的url
            articleContentHerf = re.findall(
                r'<divclass="txtconthline">.*?<ahref="(.*?.html)".*?>.*?</a>.*?</div>', (item))
            articleContentHerf = ''.join(articleContentHerf)

            # url传递给getContent函数
            articleContent = getContent(
                rootUrl if rootUrl else url, articleContentHerf)
            if articleContent != None:
                result.append(articleContent)
        return result


# 执行程序入口：循环读取url元组内的链接地址，拼接读取50页
for (key, item) in url.items():
    writeIn = []
    writeIn.append(processData(item))

    i = 1
    while i <= 50:
        writeIn.append(processData(item + "index_" + str(i) + ".html", 0, item))
        i = i + 1
    # 写入excel
    writeInFile(writeIn)


# 保存excel
Today = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
excel.save("./logs/{filename}.xlsx".format(filename = Today))