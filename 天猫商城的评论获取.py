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
    urlFront = 'https://rate.tmall.com/list_detail_rate.htm?itemId=577452915178&spuId=1057170418&sellerId=1131487699&order=3&currentPage='
    urlRear = '&append=0&content=1&tagId=70131008&posi=-1&picture=0&groupId=&ua=098%23E1hvApvWvPgvU9CkvvvvvjiWP2qUAjDWPFswAjthPmPvsjrPRsqOAj38PLLWsj1UPsuvvpvi9RGmcCH4zYMNJJbG7lf5grYFzEqWKz%2B2J%2FbUvIuWOa8B7pEmkgiXeJu%2BvpvEvv9qvSPsvbF8vvhvC9vhvvCvp89Cvv9vvUmToTsD7I9CvvOUvvVvayRgvpvIvvvvK6CvvvvvvUHFphvWspvv96CvpC29vvm2phCvhC9vvUnvphvWsp9CvhQW%2BXvvClsh6jc6%2BulgE4AxfwkKHkx%2Fgjc60fJ6EvLv%2BExrV4tYVVzh6jZ7%2B3%2Bu6jc6k24zU6sBSw5ChBODN%2B1lYE7rejpiKWzC%2BfpT29hvCvvvMMGgvpvhvvvvv8OCvvpvvUmm39hvCvvhvvvgvpvhvvvvvv%3D%3D&itemPropertyId=&itemPropertyIndex=&userPropertyId=&userPropertyIndex=&rateQuery=&location=&needFold=0&_ksTS=1602222534254_1093&callback=jsonp1094'
    for i in range(0, num):
        COMMENT_PAGE_URL.append(urlFront + str(1 + i) + urlRear)    #分别代表着每一页的评论

# 获取评论数据
def GetInfo(num):
    dataexcel = OrderedDict()
    sheet_1 = []
    row_title = [u"用户昵称", u"评价时间", u"颜色分类", u"初次评价内容", u"追加评论", u"商家回复", u"收货当天追加", u"初次评价"]
    sheet_1.append(row_title)  # 添加标题
    # 循环获取每一页评论
    for i in range(num):
        # 头文件，没有头文件会返回错误的js
        headers = {
            'cookie': 'hng=CN%7Czh-CN%7CCNY%7C156; cna=A4qEFlahYQ8CAT2vhg52cU5b; UM_distinctid=173190d6c5598a-049a0e8d88e0f9-7d7f582e-1fa400-173190d6c568ed; enc=GvdRXzzrToyLd%2BPWjeX9BwoxixdEI7XT1P%2FhjUOVPLUVHVD0da5nqBUfz0rzf8wSVMBpFdr0dObvCWGKOSGsQw%3D%3D; sm4=330200; lid=ume%E7%85%A7%E6%98%8E%E6%97%97%E8%88%B0%E5%BA%97%3A%E5%8F%AF%E5%8F%AF; xlly_s=2; _m_h5_tk=29508a036385007d58c3886d2ee4d03d_1602154501332; _m_h5_tk_enc=57554ee1a80a72650fbdcd223fcff78d; Hm_lvt_96bc309cbb9c6a6b838dd38a00162b96=1602054535,1602121812,1602145680,1602146983; t=e80e1b7ee309812a040d80ce90994a5c; _tb_token_=837d81e5773; cookie2=1f23e3d7b3656f2c50a2636c760c0f2e; l=eB_3cAHnQO9obEtMBO5Clurza77tFhO3fkPzaNbMiIncC6Y5Ej9pKStQKmE9xIxRR8XcMwtX4grm8kwTCFZU5PHfoTB7K9cdvdeXCef..; tfstk=c2FPBR0I4_CzuocCBbGeFykiTaW5Cbzu-IusEJzCHp0R2TfSvz5xwkEDeG_Cbvpd.; isg=BEdHtvTg1KGARlFR9sp-fLAa2PsRTBsuruWC6BkhRVoYiHhKPR6NfI0OKkjWe_Om',
            'user-agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36',
            'referer': 'https://detail.tmall.com/item.htm?spm=a1z10.5-b-s.w4011-17205939323.51.30156440Aer569&id=41212119204&rn=06f66c024f3726f8520bb678398053d8&abbucket=19&on_comment=1&sku_properties=134942334:3226348',
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
        jsondata = re.search('^[^(]*?\((.*)\)[^)]*$', content).group(1)
        # 用json加载数据
        content = json.loads(jsondata.text[26:-2])
        print(content)
        # content = content['rateDetail']['rateList']

        # print(content)
        for i in range(0, len(content['rateDetail']['rateList'])):
            nickname = content['rateDetail']['rateList'][i]['displayUserNick']  # 正则表达式匹配存入列表
            auctionSku = content['rateDetail']['rateList'][i]['auctionSku']
            ratecontent = content['rateDetail']['rateList'][i]['rateContent']
            ratedate = content['rateDetail']['rateList'][i]['rateDate']
            serviceRateContent = content['rateDetail']['rateList'][i]['serviceRateContent']
            position = content['rateDetail']['rateList'][i]['position']
            try:
                contentsouce = content['rateDetail']['rateList'][i]['appendComment']['content']
            except:
                contentsouce = ''
            print(contentsouce)
            reply = content['rateDetail']['rateList'][i]['reply']
            ratecontent=ratecontent.replace('<b>','')
            ratecontent=ratecontent.replace('</b>','')
            # 将数据写入TEXT文件中
            makemain =[nickname,ratedate, auctionSku,  ratecontent, serviceRateContent, reply,position,contentsouce]
            dataexcel.update({u"天猫评论": sheet_1})
            sheet_1.append(makemain)

    save_data("D:\天猫评论.xlsx", sheet_1)

# 主函数
if __name__ == "__main__":
    Page_Num = 28 #从lastpage参数中可以看到这个值是多少。
    Get_Url(Page_Num)
    GetInfo(Page_Num)