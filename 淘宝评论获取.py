# 导入需要的库
import requests
import json
import re,pyexcel
from pyexcel_xlsx import save_data
from collections import OrderedDict

import time

# 宏变量存储目标js的URL列表
COMMENT_PAGE_URL = []

# 生成链接列表
def Get_Url(num):
    COMMENT_PAGE_URL.clear()

    urlFront = 'https://rate.taobao.com/feedRateList.htm?auctionNumId=566232197899&userNumId=1692310221&currentPageNum='
    urlRear = '&pageSize=20&rateType=&orderType=sort_weight&attribute=70071004-11&sku=&hasSku=false&folded=0&ua=098%23E1hvEQvRvPOvUpCkvvvvvjiWP2Shljr2nLsZ6jljPmPpzjrPP2MWQjnhP2cUgjlPRF9CvvpvvhCv9vhv2KMNzjQx7rMNz64rzbA%2FRvhvCvvvphvRvpvhMMGvvvvCvvOv9hCvvvmgvpvIvvCvpvvvvvvvvhNjvvmCKvvvBGwvvvUwvvCj1Qvvv99vvhNjvvvmmU9CvvOCvhE2gnkIvpvUvvCCnGokn1yUvpCW9PA9ICz6%2Bu0Oe163D70Oe8gcbhet5LI65ti6pYFIeEyXezEJ0f06W3vOJ1kHsfUpeE3TmEcBKFyzhmx%2F1WmK5d8rwZXl%2Bb8reE9aU8OCvvpvvhHh39hvCvmvphm%2BvpvEvvUXjQQvvUPWC9hvCYMNzn14Pa%2BgvpvhvvCvpv%3D%3D&_ksTS=1607150953112_1293&callback=jsonp_tbcrate_reviews_list'
    for i in range(0, num):
        COMMENT_PAGE_URL.append(urlFront + str(1 + i) + urlRear)    #分别代表着每一页的评论

# 获取评论数据
def GetInfo(num):
    dataexcel = OrderedDict()
    sheet_1 = []
    row_title = [u"用户昵称", u"评价时间", u"颜色分类", u"评价", u"促销", u"图片"]
    sheet_1.append(row_title)  # 添加标题
    # 循环获取每一页评论
    for i in range(num):
        # 头文件，没有头文件会返回错误的js

        headers = {
            'cookie': 't=1f733697e6a938afb2fa9da995aaf41e; cna=1ygGF61+CBwCAT2vhg6yobYK; xlly_s=1; hng=CN%7Czh-CN%7CCNY%7C156; _m_h5_tk=ab1580791d71f4847a0fe17c32c08312_1607163177796; _m_h5_tk_enc=8c0909db7a2f57a672d6cb10f1a51120; thw=cn; cookie2=1b93b9aa7958ac34e1bfc253e9467c48; _tb_token_=eb6e5e8dbbb1e; v=0; _samesite_flag_=true; unb=863198918; lgc=mrs%5Cu795E%5Cu4ED9%5Cu59D0%5Cu59D0; cookie17=W89PQor2srOF; dnk=mrs%5Cu795E%5Cu4ED9%5Cu59D0%5Cu59D0; tracknick=mrs%5Cu795E%5Cu4ED9%5Cu59D0%5Cu59D0; _l_g_=Ug%3D%3D; sg=%E5%A7%9082; _nk_=mrs%5Cu795E%5Cu4ED9%5Cu59D0%5Cu59D0; cookie1=W8slpuGg2cAvxSV%2FMPG3iZzNuuc31%2FsDm9LqtrxJmyI%3D; sgcookie=E100UUJ8hJ1LUUGxIxjCPiTdvjF%2Fhp4ZHlj5adaOcE44h8V67Ly10Gfagy7AMZSvfowRjvaRzyKR5hzq81MokJe9TA%3D%3D; uc3=lg2=Vq8l%2BKCLz3%2F65A%3D%3D&vt3=F8dCuf2ADm0F1%2FWx7%2Bs%3D&nk2=DkLfUfFa%2FoYhSFs%3D&id2=W89PQor2srOF; csg=108a94b9; skt=92e4cbe3b490df76; existShop=MTYwNzE1ODUzNw%3D%3D; uc4=nk4=0%40DCdu2SkLfC%2F7OVazLRBLW3KkDnQOKw%3D%3D&id4=0%40Wey%2BQg2K6Zg8MYGGsq9mCXriUfM%3D; _cc_=V32FPkk%2Fhw%3D%3D; enc=c4WRuAAZ1U2yUD24HK%2FzcV9AYz1R8RIf5wCC4vG9ys7%2BsxFJcLei9RLUAPIGzviZc9OtGeDTpPGrHWBQtkjvLg%3D%3D; mt=ci=67_1; uc1=existShop=false&cookie21=UIHiLt3xTIkz&cookie14=Uoe0almLgR8bfw%3D%3D&cookie16=VFC%2FuZ9az08KUQ56dCrZDlbNdA%3D%3D&cookie15=VFC%2FuZ9ayeYq2g%3D%3D&pas=0; x5sec=7b22726174656d616e616765723b32223a223366313662653134323866643134616566313761653266363964326666313964434a535772663446454e6e4d696553453674446a50526f4c4f44597a4d546b344f5445344f7a453d227d; tfstk=cQu5BOYGdTX7SFKFzbO4Tk3rGf4fa71_O3woN0fFnxgfv1h7HsmP7R94RDXdklFf.; l=eBIKwXbgv06TC1y3BO5alurza77TNIObzsPzaNbMiInca1lPML4lhNQ2ssI9ydtjgt5DLetrJmGnARFvSmUU-xTjGO0qOC0eQtJ68e1..; isg=BMDAsrmNi0i9OHRVzY_FKyyYkU6SSaQTBzAUHTpQpVtMtWLf4lpLolDDzR11BVzr',
            'user-agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36',
            'referer': 'https://item.taobao.com/item.htm?spm=a21ag.11815245.0.0.710a50a54LehUI&id=566232197899',
            'accept': '*/*',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'zh-CN,zh;q=0.9'
        }#伪装成浏览器访问，防止乱码或者防止访问失败
        # 解析JS文件内容
        print(i)
        print(COMMENT_PAGE_URL[i])
        content = requests.get(COMMENT_PAGE_URL[i], headers=headers).text  # 调用http接口并获取他的文字
        # time.sleep(5)
        # 筛选json格式数据
        print(content)
        first=content.find('"comments":') +11
        end=content.find(',"currentPageNum"')
        content=content[first: end]
        # 用json加载数据
        content = json.loads(content)
        print(content)
        # content = content['rateDetail']['rateList']

        # print(content)
        for comment in content:
            nickname = comment['user']['nick']

            ratedate = comment['date']
            text = comment['content']  # 正则表达式匹配存入列表
            auctionPic = comment['user']['creditPic']
            sku = comment['auction']['sku']
            promotionType = comment['promotionType']  # 正则表达式匹配存入列表
            text = text.replace('<em>', ' ')
            text = text.replace('</em>', ' ')

            # 将数据写入TEXT文件中
            makemain =[nickname,ratedate,  sku, text,promotionType,auctionPic]
            dataexcel.update({u"天猫评论": sheet_1})
            sheet_1.append(makemain)

    save_data("D:\淘宝评论.xlsx", sheet_1)

# 主函数
if __name__ == "__main__":
    Page_Num = 1 #从lastpage参数中可以看到这个值是多少。
    Get_Url(Page_Num)
    GetInfo(Page_Num)