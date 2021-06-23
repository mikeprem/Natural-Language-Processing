# -*- coding: utf-8 -*-
"""
Created on Thu Aug 17 17:40:02 2017

@author: WB511527
"""
from nltk import RegexpParser
import re
import nltk
import pandas as pd
import numpy as np


df = pd.DataFrame(np.zeros((50000,1)))
df.columns=['KeyWords']

patterns = """ P: {<NNP><VBZ>}
               A: {<NN.><NN.>}
               B: {<NN><NN>} 
               C: {<NNP><NN><NN>}
               D: {<JJ><NN>}
               E: {<JJ.><NN.>}
               F: {<JJ><NN.><NN.>}
               """

PChunker = RegexpParser(patterns)
f=open('1.txt','r',encoding='latin1').read()
words = nltk.word_tokenize(f)
tags=nltk.pos_tag(words)
result=PChunker.parse(tags)
n=0

for j in range(0,len(result)):
  if len(result[j][0][0])>1: 
     stringg=''   
     for i in range(0,len(result[j])):
      stringg=stringg+result[j][i][0]+' '
     df.iloc[n,0]=stringg
     n=n+1
     print(j,n)       
           
df=df.astype(str)
df


writer = pd.ExcelWriter('NounPhrases.xlsx', engine='xlsxwriter')
df.to_excel(writer,sheet_name='Sheet1',index=False)

writer.save()
