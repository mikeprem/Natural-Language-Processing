#step 1
from nltk import RegexpParser;
import re
import nltk
import pandas as pd
import numpy as np
import os
from collections import OrderedDict
directory=os.fsdecode(r'C:\Users\wb511527\OneDrive - WBG\Desktop\New folder (2)\NLP_Skills_Test_Data\Articles_of_incorporations')
os.chdir(directory)
l=len(os.listdir(directory))
df=pd.DataFrame(np.zeros((l,2)))
df=df.astype(str)
df.columns=['PIDS','Text Section']
i=0
from tika import parser
import re

for file in os.listdir(directory):
    filename=os.fsdecode(file)
    if len(filename)>=1:
        stringg=parser.from_file(filename)
        stringg=str(stringg['content'])
        stringg=re.sub('[^a-zA-Z0-9\' ]',' ',stringg)
        stringg=re.sub('[ ]+',' ',stringg)
        df.iloc[i,0]=filename
        df.iloc[i,1]=stringg
    i+=1
    print(i)
#df=df.replace("[ ]+"," ",regex=True)    
df=df[df["Text Section"]!="0.0"]
key = "3b0cce824fd04bbcb9c9c7867f643b7c"
endpoint = "https://spach-ifc.cognitiveservices.azure.com/"

from azure.ai.textanalytics import TextAnalyticsClient
from azure.core.credentials import AzureKeyCredential

def authenticate_client():
    ta_credential = AzureKeyCredential(key)
    text_analytics_client = TextAnalyticsClient(
            endpoint=endpoint, credential=ta_credential)
    return text_analytics_client

client = authenticate_client()
strr=".".join(list(df["Text Section"]))
df=df.replace("[.]+",".",regex=True)  
lis=[]
m=0
while m<len(strr[:]):
    
        documents = [strr[m:min(m+5120,len(strr))]]
        response = client.extract_key_phrases(documents = documents)[0]

        if not response.is_error:
            
           for phrase in response.key_phrases:
                #print("\t\t", phrase)
                lis.append(phrase)
        else:
            print(response.id, response.error)
        m=m+5120
        print(m)
lis=pd.DataFrame(lis)
lis.to_excel('temp.xlsx')

lis=lis.drop_duplicates(subset=[0],keep='last')

#
liss=list(lis[0])
lisres=[str(item).replace(" ","_") for item in liss]
df=df.astype(str)
for i in range(0,len(lisres)):
    df=df.replace(liss[i],lisres[i],regex=True)

df.to_excel('tempdf.xlsx')
for i in range(0,len(lisres)):
    df=df.replace(lisres[i],liss[i],regex=True)
#word2vec
from gensim.models import Word2Vec
import gensim
import pandas as pd
import os
import re
from nltk.stem import PorterStemmer

ps = PorterStemmer()
from nltk.corpus import stopwords
listt=list(stopwords.words('english'))
for t in listt:
    df=df.replace(" "+t+" "," ",regex=True)

#df["Extract"]=df["Extract"].apply(lambda x: " ".join(ps.stem(item) for item in x.split()))

sentences=list(df["Text Section"])
sentences=[item.split() for item in sentences]
#gmodel=gensim.models.KeyedVectors.load_word2vec_format(sentences[:300])
gmodel=Word2Vec(sentences,size=100, window=20, min_count=1)
os.chdir(r'C:\Users\wb511527\OneDrive - WBG\Desktop\New folder (2)\NLP_Skills_Test_Data')
refdf=pd.read_excel("CG_domain_and_topics.xlsx",encoding="latin1")
lis=[]
lis1=[]
lis2=[]
for i in refdf["CG methodology "]:
    try:
         ms=gmodel.wv.most_similar(positive=list(str(i).split()))
         for x in ms[:10]:
           lis2.append(i)
           lis.append(x[0])
           #lis1.append(x[1])
           print('success')
    except:
       print("Hard Luck!")
#lis=list(set(lis))
cv=pd.DataFrame(lis)
cv['val']=lis1
cv['ref']=lis2
cv.to_excel("Keywords-Word2vec1.xlsx",index=False)
lis['flag']=['']*len(lis)

lis.columns=['Key Phrases','flag']
for i in range(0,len(lis)-1):
  if lis.iloc[i][1]!="1":  
    for j in range(i+1,len(lis)):
        if lis.iloc[j][1]!="1":
          print(i,j)
          if len(set(lis.iloc[i][0].lower().split()).intersection(set(lis.iloc[j][0].lower().split())))>=2:
             lis.iloc[j,1]="1"
lis.to_excel('outa.xlsx',sheet_name='Sheet1',index=False)
directory=os.fsdecode(r'C:\Users\wb511527\OneDrive - WBG\Desktop\syn\Process')
os.chdir(directory)
lis=pd.read_excel('outa.xlsx',encoding="latin1")
lis['len']=[len(item) for item in lis['Key Phrases']]
lis=lis[lis["len"]<=40]
del lis['len']
lis=lis[lis['flag']!=1]
directory=os.fsdecode(r'C:\Users\wb511527\OneDrive - WBG\Desktop\syn\sample')
os.chdir(directory)
l=len(os.listdir(directory))
df=pd.DataFrame(np.zeros((l+5,2)))
df=df.astype(str)
df.columns=['PIDS','Extract']
i=0
for file in os.listdir(directory):
    filename=os.fsdecode(file)
    if len(filename)>=1:
        stringg=parser.from_file(filename)
        stringg=str(stringg['content'])
        stringg=re.sub('\n',' ',stringg)
        stringg=re.sub('-',' ',stringg)
        stringg=re.sub('[^a-zA-Z0-9\' ]',' ',stringg)
        stringg=re.sub('[ ]+',' ',stringg)
        df.iloc[i,0]=filename[:7]
        df.iloc[i,1]=stringg
    i+=1
    print(i)
#df=pd.read_excel("Copy of PADS_DATABASE.xlsx",encoding='latin1')
df["Extract"]=df["Extract"].apply(lambda x: re.sub("[ ]+"," ",re.sub("[^a-zA-Z ]"," ",x)))
sentences=list(df["Extract"])
#sentences=[item.split() for item in sentences]
strr=" ".join([item for item in sentences])

li=lis.copy()
lis=" ".join(lis['Key Phrases'])
liss=lis.split()
liss=list(set(list(liss)))
strr=strr.lower()
liss2=['']*len(liss)
for i in range(0,len(liss)):
     liss2[i]=strr.count(liss[i].lower())
     print(i)
dff=pd.DataFrame(liss)
dff['count']=liss2
mer=dff.copy()
mer=mer.sort_values(by=["count"],ascending=False)
mer=mer[:int(len(mer)/10)]
mer['len']=[len(item) for item in mer[0]]
mer=mer[mer["len"]>=4]
lis=li.copy()
lis=[item for item in lis['Key Phrases'] if len(set(str(item).lower().split()).intersection(set(list(mer[0]))))>=1]
lis=pd.DataFrame(lis)

#lis=pd.read_excel('Taxonomy-Augmentation-CTC-Intranet-Design-IT-20201030.xlsx',encoding="latin1")
lis=lis.iloc[:,:1]
reftt=pd.read_excel('metadata.xlsx',encoding="latin1")
refbt=pd.read_excel('metadata-BT.xlsx',encoding="latin1")
lis.columns=['tax']
lis['TT']=['']*len(lis)
lis['BT']=['']*len(lis)
for i in range(0,len(lis)):
    for j in range(0,len(reftt)):
       print(i,j)
       if len(set(lis.iloc[i][0].lower().split()).intersection(set(reftt.iloc[j][0].lower().split())))>=2:
          lis.iloc[i,1]="Yes"
       else:
          lis.iloc[i,1]="No"
for i in range(0,len(lis)):
    for j in range(0,len(refbt)):
       print(i,j)
       if len(set(lis.iloc[i][0].lower().split()).intersection(set(refbt.iloc[j][0].lower().split())))>=2:
          lis.iloc[i,2]="Yes"
       else:
          lis.iloc[i,2]="No"
lis['Keep/Delete']=[""]*len(lis)
lis.columns=['Taxonomy','Match with Topical Taxonomy?','Match with Business Taxonomy?','Keep/Delete']
lis.to_excel("Taxonomy-Augmentation-Syndications.xlsx",index=False)




            

