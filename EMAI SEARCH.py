from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import numpy as nu
import io
from bs4 import BeautifulSoup as soup

import pandas as pd
import re
import locale


driver = webdriver.Chrome("D:\webdriver/chromedriver.exe")
   # 打开 Chrome 浏览器
driver.get('https://login.huawei.com/login/?redirect=http%3A%2F%2Fapp.huawei.com%2Ftcsmadmin%2Findex.do%3Ftimestamp%3D1563869144028%26catalogId%3D-1&lang=en&msg=1&v=v3.33')

time.sleep(8)
driver.find_element_by_xpath("//a[@title='Logout']").click()
#mouse=driver.find_element_by_class_name("psi-name")
#ActionChains(driver).move_to_element(mouse).perform()
#driver.findElement(By.linkText("Exit")).click();
time.sleep(2)
#driver.find_element_by_id("psi_logout").click()

login = driver.find_element_by_class_name("user")
login.send_keys("")

password=driver.find_element_by_class_name("psw")
password.send_keys("")
time.sleep(1)
driver.find_element_by_class_name("btn").click()

#driver.find_element_by_xpath('//a[@href="'http://app.huawei.com/tcsmadmin/logout.do'"]')

s1 = Select(driver.find_element_by_id('currentrole')) 
s1.select_by_index(0) 

driver.find_element_by_id("leftMenuTree_6_a").click()




list=[]
clist=[]
flist=[]
with open('D:/test1.txt', 'r') as myfile:
  data = myfile.read()

emai=data.split('\n')


for d in range(0,len(emai)):
    datou=emai[d]+","+"865473025229694"+","+"865473025231963"+","+"865473025233605"+","+"865473025233720"+","+"865473025233845"
    print(emai[d])
    driver.find_element_by_xpath("//input[@name='serialId']").send_keys(datou)
    driver.find_element_by_xpath("//input[@name='search']").click()
    
    page = soup(driver.page_source, "html.parser")
    links=page.find_all("tr",align="center")
    if len(links)<6:
        list.append(emai[d])
        list.append("product_offering")
        list.append("contract_no")
        list.append("packing")
        list.append("shipping_date")
        list.append("arrive")
        list.append("destination")
  
        list.append("signed_customer")
        list.append("active_date")
        list.append("activation_country")
        list.append("quantity")
        
    else:
        for i in range(0,len(links)):
            if links[i].contents[24].contents[1].contents[0]!="CN":
                item=links[i].contents[2].contents[1].contents[0]
                
                product_offering=links[i].contents[4].contents[1].contents[0]
                
                contract_no=links[i].contents[6].contents[1].contents[0]
                
                packing=links[i].contents[8].contents[1].contents[0]
                
                shipping_date=links[i].contents[10].contents[1].contents[0]
                
                arrive=links[i].contents[12].contents[1].contents[0]
                
                destination=links[i].contents[14].contents[1].contents[0]
                

                
                signed_customer=links[i].contents[18].contents[1].contents[0]
                
                active_date=links[i].contents[22].contents[1].contents[0]
                
                activation_country=links[i].contents[24].contents[1].contents[0]
                
                quantity=links[i].contents[30].contents[1].contents[0]
                
                list.append(emai[d])
                list.append(product_offering)
                list.append(contract_no)
                list.append(packing)
                list.append(shipping_date)
                list.append(arrive)
                list.append(destination)
            
                list.append(signed_customer)
                list.append(active_date)
                list.append(activation_country)
                list.append(quantity)
                break
    driver.find_element_by_xpath("//input[@name='serialId']").clear()

name=['EMAI','product_offering','contract_no','packing','shipping_date','arrive','destination','signed_customer','active_date','activation_country','quantity',]
#outcome=pd.DataFrame(columns=name)
flist.clear()
k=0;
for i in range(0,int(len(list)/11)):
    #e=list[i]
    print(k)
  

    flist.append([list[k],list[k+1],list[k+2],list[k+3],list[k+4],list[k+5],list[k+6],list[k+7],list[k+8],list[k+9],list[k+10]])
   
    
    k+=11

df = pd.DataFrame(columns=name,data=flist)     
df.to_excel("D:/out.xlsx",encoding="utf-8")

#for link in links:
 #   print link.name, link['title'], link.get_text()

