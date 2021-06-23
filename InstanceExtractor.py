import pandas as pd
import os
from tika import parser
import re
import numpy as np
from nltk.stem import PorterStemmer
ps = PorterStemmer()
directory=os.fsdecode(r'C:\Users\wb511527\OneDrive - WBG\IEG-Nutrition\Additional\PJPR')
os.chdir(directory)
tax=pd.read_excel('Taxonomy.xlsx',encoding='latin1')
df=pd.DataFrame(np.empty((100000,len(tax)+2)))
lis=list(tax[['Taxonomy']].values.flatten())
lis.insert(0,'Docname')
lis.insert(1,'Instance')
df.columns=lis
j=0
listt=list(tax[['Taxonomy']].values.flatten())
#listtt=[' '.join([ps.stem(it).lower() for it in item.split()]) for item in listt]
n=0
for file in os.listdir(directory):
   n=n+1
   filename=os.fsdecode(file)
   if filename.endswith('.txt'):   
      f=open(filename,'r',encoding='latin1').read()
      f=str(f)
      f=re.sub('[ ]+',' ',f)
      #f=' '.join(temlis)
      df=df.astype(str)
      for i in range(0,len(listt)):
         m=0
         pattern=re.compile('(?<=[^0-9][.][^0-9])(?:.|\n){0,350} '+listt[i]+' (?:.|\n){0,1500}(?:[^0-9A-Z][.])')
         result = [word for word in  pattern.findall(f)]
         if len(result)>=0:
          for m in range(0,len(result)):
            df.iloc[j+m][0]=filename
            df.iloc[j+m][1]='('+listt[i]+') '+result[m]
          j=j+m 
          print(n,i,j)  
#df=df.dropna(how='all')
df=df[:j]
l=len(df)
df.to_excel('first.xlsx',index=False)
df=pd.read_excel('first.xlsx',encoding='latin1')
i=0
while i!=l-1:
    m=0
    for j in range(i+1,l-1):
       if df.iloc[i][0]==df.iloc[j][0]:
           m=m+1
           df.iloc[i,1]+='\n'+str(m+1)+'\n==>>'+df.iloc[j][1]   
           print(i)
    i=i+m+1          
          
def lenre(row):
     row['len']=len(row['Instance']) 
     return(row)     
df=df.apply(lenre,axis=1)     
df=df.sort_values('len').drop_duplicates(subset='Docname', keep='last')
for i in range(0,len(df)):
    for j in range(2,len(df.columns)):
       if df.columns[j] in df.iloc[i][1]:
          df.iloc[i,j]=df.iloc[i][1].count(df.columns[j])
       else:
          continue
df=df.replace('(?:\n){2,100}','',regex=True)
df['Instance']=df['Instance'].apply(lambda x: '1\n==>'+x)
df.to_excel('PJPR_instance.xlsx',index=False)           

    
    