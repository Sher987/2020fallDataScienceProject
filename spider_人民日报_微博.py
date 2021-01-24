# coding=utf-8

import sys
from lxml import etree
from bs4 import BeautifulSoup
import re
import urllib
import urllib.request
import requests
import json
import xlwt
import socket


def main():
    baseurl1="https://weibo.com/rmrb?is_all=1&stat_date=202003&page={}&display=0&retcode=6102"  #  人民日报微博2019年12月微博
    datalist1=getData(baseurl1)
    savepath1 = ".\\人民日报微博2020年3月10日至31日.xls"
    saveData(datalist1,savepath1)
    baseurl2=""

head = {
    # 模拟浏览器访问网页
        # "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        # "accept-encoding": "gzip, deflate, br",
        # "accept-language": "zh-CN,zh;q=0.9",
        # "cache-control": "max-age=0",
        # "host" : "weibo.com",
        "Cookie": "login_sid_t=569dd6f5549e26e915e2276a64217d12; cross_origin_proto=SSL; _s_tentry=passport.weibo.com; Apache=329628730611.3022.1611066632287; SINAGLOBAL=329628730611.3022.1611066632287; ULV=1611066632299:1:1:1:329628730611.3022.1611066632287:; wvr=6; ALF=1642774499; SCF=AnEz52KI7VvC4EkuT6Zxvr2nAuSQ5ost8RvH6fMPOlDQ81Q7v8niO0uM4rQhX0ZWyXT_ruBsZqOQXxGWuD6qLUI.; wb_view_log_6546793530=1366*7681; WBtopGlobal_register_version=2021012222; crossidccode=CODE-yf-1L2XtW-3mKDw0-qTh1Bu8erTqopuJc26334; SSOLoginState=1611325939; SUB=_2A25NDpGjDeThGeBL71QW-S3JyDyIHXVu8D_rrDV8PUJbkNANLUGnkW1NRx55TC2jsE5G6DJp57bwAtM9UWUDrOfx; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WhqqhpQ6vTE-lTD1IBfxZML5NHD95QcSKBcS0.0SKe7Ws4DqcjDi--ciK.4i-zXi--fi-2Xi-24i--fi-2RiKn0i--fi-2RiKn0S0-0eKnt; UOR=,,graph.qq.com; webim_unReadCount=%7B%22time%22%3A1611326326539%2C%22dm_pub_total%22%3A0%2C%22chat_group_client%22%3A0%2C%22chat_group_notice%22%3A0%2C%22allcountNum%22%3A67%2C%22msgbox%22%3A0%7D",
        # "referer": "https://weibo.com/rmrb?topnav=1&wvr=6&topsug=1",
        # "upgrade-insecure-requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"
    }



findcontent=re.compile(r'<divclass="WB_textW_f14"node-type="feed_list_content"nick-name="人民日报">n(.*?)</div>n',re.S)
# findtext=re.compile(r'【(.*?)】')
findcommentnum=re.compile(r'<em>([0-9]+)</em>',re.S)

def getData(baseurl):
    datalist = []

    for i in range(0,36):  # 2019年12月8日到2019年12月31日
        html = askURL(baseurl.format(i))
        every_id = re.compile('name=(\d+)', re.S).findall(str(html))  # 获取评论页面需要的id
        # print(html)
        # print(every_id)
        news=[]
        comment=[]
        i=0
        for id in every_id:
            comment.append(get_comments(id))
            i=i+1
        print(comment)
        soup = BeautifulSoup(html, "html.parser")

        content=re.findall(findcontent,str(soup).replace(" ",""))
        # commentnum=re.findall(findcommentnum,str(soup))
        # print(commentnum)
        # print(soup)
        j=0
        for item in content:
            item=re.findall('[\u4e00-\u9fa5。，!#【】0-9]+',str(item))
            item=re.sub(r"'[0-9]+'","",str(item))
            item=str(item).replace(",","").replace(" ","").replace("'","").replace("[","").replace("]","")
            if(str(item).__contains__("疫") or str(item).__contains__("新冠") or str(item).__contains__("湖北" or str(item).__contains__("武汉"))):
                news.append(item)
                news.append(comment[j])
                datalist.append(news)
                j=j+1

                news=[]
            else:
                j+=1
        # print(datalist)
    return datalist

def saveData(datalist,savepath):
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)  # 创建workbook对象 样式压缩
    sheet = book.add_sheet('人民日报微博', cell_overwrite_ok=False)  # 创建工作表  单元可覆盖
    col=("微博内容","微博热门评论")
    for i in range(0,2):
        sheet.write(0,i,col[i])  #列名
    for i in range(0, len(datalist)):
        data = datalist[i]
        for j in range(0, 2):
            sheet.write(i + 1, j, data[j])

    book.save(savepath)




def get_comments(id):
    link = "https://weibo.com/aj/v6/comment/small?ajwvr=6&act=list&mid={}".format(id)
    # html=askURL(link)
    info = []
    try:
        res = requests.get(link, headers=head,timeout=10)
        response = res.json()
        count = response['data']['count']
        html = etree.HTML(response['data']['html'])
        info = html.xpath("//div[@node-type='replywrap']/div[@class='WB_text']/text()")  # 评论信息
        info = "".join(info).replace(" ", "").replace("：", "").replace("\xa0", " ").split("\n")
        info.pop(0)
    except requests.exceptions.SSLError as e:
        print(e)
    except socket.timeout:
        print('Time Out!')
    return info



def askURL(url):

    request = urllib.request.Request(url=url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8","ignore").replace("\\", "")
        # res = requests.get(url, headers=head)
        # response1 = res.content.decode().replace("\\", "")

        # html=requests.get(url,headers=head)
        #print(html)
    except urllib.error.URLError as e:
        if (hasattr(e, "code")):
            print(e.code)
        if (hasattr(e, "reason")):
            print(e.reason)
    return html


if __name__ == '__main__':
    main()
    print("succeed!")