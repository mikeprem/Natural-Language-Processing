# -*- coding: utf-8 -*-
"""
Created on Wed Dec 16 10:07:28 2020

@author: wb511527
"""
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
        stringg=re.sub('[^a-zA-Z. ]',' ',stringg)
        stringg=re.sub('[ ]+',' ',stringg)
        stringg=stringg.lower()
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
lis.to_excel('temp.xlsx',index=False)


lisd=pd.read_excel("temp.xlsx",encoding='latin1')
lisd['word count']=[str(item).count(" ") for item in lisd[0]]
lisd=lisd[(lisd["word count"]==1)]
lisd['len']=[min([len(str(item1)) for item1 in item.split()]) for item in lisd[0]]
lisd=lisd[lisd["len"]>=4]

lisd=lisd.drop_duplicates(subset=[0],keep="last")


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

#df["Text Section"]=df["Text Section"].apply(lambda x: " ".join(ps.stem(item) for item in x.split()))

sentences=list(df["Text Section"])
sentences_clubbed=" ".join(sentences)
sentences_joined=sentences_clubbed.split(".")

sentences_split=[item.split() for item in sentences_joined]
#gmodel=gensim.models.KeyedVectors.load_word2vec_format(sentences[:300])
gmodel=Word2Vec(sentences_split,size=3000, window=20, min_count=1)
os.chdir(r'C:\Users\wb511527\OneDrive - WBG\Desktop\New folder (2)\NLP_Skills_Test_Data')
refdf=pd.read_excel("CG_domain_and_topics.xlsx",encoding="latin1")
lis1=[]
lis2=[]
lis3=[]
refdf=refdf[refdf["CG methodology "].map(str)!="nan"]
refdf=refdf.drop_duplicates(subset=["CG methodology "],keep="last")
refdf=refdf.replace('[^a-zA-Z ]',' ',regex=True)
refdf=refdf.replace('[ ]+',' ',regex=True)
cv=pd.DataFrame(np.zeros((len(lisd[0]),len(refdf))))
cv.index=lisd[0]
cv.columns=list(refdf["CG methodology "])

n=0
for i in [str(item).lower() for item in list(refdf["CG methodology "])]:
    m=0
    for j in list(lisd[0]):
      try:
         cv.iloc[m,n]=gmodel.wv.n_similarity(list(str(i).lower().split()),list(str(j).lower().split()))
         m=m+1
      except:
         cv.iloc[m,n]=0
         m=m+1
    n=n+1
    print(n)
#lis=list(set(lis))
pdf=cv.copy()

topics=[]
Keywords=[]
pdf=pdf.astype(str)
for i in pdf.columns:
    tempdf=pdf[[i]].copy()
    tempdf=tempdf.sort_values(by=[i],ascending=False)
    tempdf=tempdf[:15]
    tempstr=",".join(list(tempdf.index))
    topics.append(i)
    Keywords.append(tempstr)
    
extractedandmapped=pd.DataFrame(topics)
extractedandmapped['keywords']=Keywords

extractedandmapped.to_excel("results2.xlsx",index=False)
    

gmodel.wv.n_similarity(['board','diversity'],['female','executive'])

ms=gmodel.wv.most_similar(positive=list(str("board diversity").split()))

