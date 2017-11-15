import requests
from bs4 import BeautifulSoup
import urllib
from collections import OrderedDict  
from pyexcel_xls import get_data  
from pyexcel_xls import save_data
import os
import random

def openAndclean_web(website_link):
    headers = {
    'Referer':'http://pubsonline.informs.org/toc/mnsc/63/11',
    'Host':'pubsonline.informs.org',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36',
    'Accept':'*/*'
    }

    proxies = {"https": "http://61.135.217.7:80" }   
    s = requests.session()
    s = BeautifulSoup(s.get(website_link,proxies=proxies,headers=headers).content, "lxml")
    return s

def get_issues_list(open_link):
    s = openAndclean_web(open_link)
    volume_list=s.find_all('div',{'class':"row js_issue"})
    issue_list=[]
    for i in volume_list:
        issue_list.append(i.find('a')['href'])
        
    return issue_list
def get_paperlinks(open_link):
    s = openAndclean_web(open_link)
    main_root = s.find_all('a',{'class':"ref nowrap"})
    paper_links=[]
    #paper_names=[]
    for i in main_root:
        #print('文章名称：'+i.get_text()+'\n'+'网页地址：http://pubsonline.informs.org'+i['href']+'\n')
        paper_links.append('http://pubsonline.informs.org'+i['href'])
        #paper_names.append(i.get_text())
    return paper_links

def down_paperinfo(paper_link):
    authors=[]
    #title
    attributes=[]
    #abstract
    s = openAndclean_web(paper_link)
    #提取文章标题信息
    title=s.find('h1',{'class':"chaptertitle"}).get_text().strip().replace('\n','')
    #提取文章作者名
    Findauthor=s.find_all('div',{'class':'header'})
    for i in Findauthor:
        authors.append(i.get_text())
    #提取关键词
    Findkeywords=s.find_all('a',{'class':'attributes'})
    for i in Findkeywords:
        attributes.append(i.get_text())
    #提取文章摘要
    abstract_root = s.find('div',{'class':'abstractSection abstractInFull'})
    try:
        abstract=abstract_root.get_text()
    except TypeError and AttributeError:
        abstract = []
    return [title,authors,attributes,abstract]
    
 # 写Excel数据, xls格式  
def save_xls_file(m1,m2,m3,m4,namestr):  
    data = OrderedDict()  
    # sheet表的数据  
    #sheet_1 = []  
    #row_1_data = [u"标题", u"作者", u"关键词", u"摘要"]   # 每一行的数据  
    row_2_data = [m1, m2, m3,m4]  
    # 逐条添加数据
    #sheet_1.append(row_1_data)  
    sheet_1.append(row_2_data)  
    # 添加sheet表  
    data.update({u"这是信息": sheet_1})  
  
    # 保存成xls文件  
    save_data(namestr, data)     
    


#(testa,testb)=get_paperlinks('http://pubsonline.informs.org/toc/mnsc/63/11')
if __name__ == '__main__':
    issue_list= get_issues_list('http://pubsonline.informs.org/loi/mnsc')
    issue_list = issue_list[:3]
    for each_list in issue_list:
        vol_name = each_list.split('/')[-2]
        issue_name = each_list.split('/')[-1]
        print(vol_name+' '+issue_name)
        if str(vol_name) not in os.listdir(r'/home/caesarhtx/PycharmProjects/paper_crawler'):
            os.mkdir(str(vol_name))
            os.chdir(r'/home/caesarhtx/PycharmProjects/paper_crawler/%s' % vol_name)
        else:
            os.chdir(r'/home/caesarhtx/PycharmProjects/paper_crawler/%s' % vol_name)
        try:
            paper_link_list = get_paperlinks(each_list)
        except TimeoutError:
            print('IssueErrorOccur:'+each_list)
            continue
        data = OrderedDict()
        sheet_1 = []
        row_1_data = [u"标题", u"作者", u"关键词", u"摘要"]
        sheet_1.append(row_1_data)
        data.update({u"这是标题": sheet_1})
        save_data("%s.xls" % issue_name, data)
        for link in paper_link_list:
            try:
                [ti, au, at, ab] = down_paperinfo(link)
            except TimeoutError:
                print('PaperErrorOccur:'+link)
                continue
            save_xls_file(ti, au, at, ab, "%s.xls" % issue_name)
