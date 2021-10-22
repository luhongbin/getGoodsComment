'''Scraping reviews of 10 smartphones from the amazon.in
saving them in different csv files.
'''

import requests
from bs4 import BeautifulSoup
import pandas as pd


#links of the 10 mobile phones
mobiles = ['Redmi-Note-Pebble-Grey-Storage/dp/B086977TR6/ref=zg_mg_1389432031_1?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'Test-Exclusive-614/dp/B07HGJJ559/ref=zg_mg_1389432031_2?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'Samsung-Galaxy-Ocean-128GB-Storage/dp/B07HGGYWL6/ref=zg_mg_1389432031_3?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'OnePlus-Nord-Gray-128GB-Storage/dp/B08695ZSP6/ref=zg_mg_1389432031_5?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'Test-Exclusive-712/dp/B07DJCJBB3/ref=zg_mg_1389432031_7?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'Redmi-Prime-Storage-Display-Camera/dp/B08696XM8J/ref=zg_mg_1389432031_9?_encoding=UTF8&psc=1&refRID=7BXWY1FD028BTNASWPK0',
           'Vivo-Storage-Additional-Exchange-Offers/dp/B07KXCKPZZ/ref=zg_mg_1389432031_79?_encoding=UTF8&psc=1&refRID=FKT3SBY1ZM8TESTC1GNX',
           'Samsung-Galaxy-Storage-Additional-Exchange/dp/B086KCDGTQ/ref=zg_mg_1389432031_17?_encoding=UTF8&psc=1&refRID=BNDYYCE3WJS3GZ4M2QGZ',
           'Test-Exclusive-749/dp/B07DJ8K2KT/ref=zg_mg_1389432031_66?_encoding=UTF8&psc=1&refRID=FKT3SBY1ZM8TESTC1GNX',
           'Magic-Storage-Additional-Exchange-Offers/dp/B089MTDPKN/ref=zg_mg_1389432031_85?_encoding=UTF8&psc=1&refRID=FKT3SBY1ZM8TESTC1GNX']


for mob in mobiles:
    #getting the link of product page
    link ='https://www.amazon.in/'+mob

    #converting the page content to soup using BeautifulSoup
    header = {'user-agent': 'Mozilla/4.0'}
    res = requests.get(link, headers=header)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, 'lxml')
    print(res.text)
    #getting the names of the customers
    names = soup.find_all('span' , class_='a-profile-name')

    #cust_name is the list of names of the customers
    cust_name = []
    for name in names:
        cust_name.append(name.get_text())

    #list cust_review will contains the reviews of customers
    review = soup.find_all("span",{"data-hook":"review-body"})
    cust_review=[]
    for rev in review:
        cust_review.append(rev.get_text())

    #removing all the \n
    cust_review[:] = [review.lstrip('\n') for review in cust_review]

    #storing the reviews in the csv file
    df = pd.DataFrame()
    df['Customer Name'] = cust_name
    df['Customer Review'] = cust_review
    title = soup.find("span", class_="a-size-large product-title-word-break")

    #we will use title of the product as individual file name
    #each csv file with mobile title will have reviews of it
    #we will use 20 characters of title as a file name
    newname = title.get_text().strip()[:30]

    df.to_csv(r'D:\MOBILE_' + newname + '.csv')