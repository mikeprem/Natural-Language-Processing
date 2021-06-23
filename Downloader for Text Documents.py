# -*- coding: utf-8 -*-
"""
Created on Tue Jul 10 16:27:02 2018
@author: WB511527
#Input should be an excel with two columns
#Column 1: Project Id Column 2: link to the document
"""
from selenium import webdriver
options = webdriver.ChromeOptions()
options.add_experimental_option('useAutomationExtension', False)
browser = webdriver.Chrome(options=options,executable_path=r"C:\\WBG\\chromedriver.exe")
import pandas as pd
import os
directory=r'C:\Users\wb511527\OneDrive - WBG\Desktop\New folder (2)'
os.chdir(os.fsdecode(directory))
df=pd.read_excel('ImageBank_URLs.xlsx',encoding='latin1')
m=0
misslis=[]
i=0
for i in range(0,len(df)):
  try:
     browser.get("http://ifcintranet.ifc.org/wps/wcm/connect/dept_int_content/ims/ifc+board/resources+-+board+paper+2")
     f=browser.find_element_by_xpath("html").text

     with open(df.iloc[i][0]+".txt", "w",encoding='utf-8') as text_file:
       text_file.write(f)
     print(i)
  except:
     m=m+1 
     misslis.append(df.iloc[i][0])
df=df[df['Project ID'].isin(misslis)]
df.to_excel('Faillist.xlsx',index=False)







