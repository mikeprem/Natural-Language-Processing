# -*- coding: utf-8 -*-
"""
Created on Sun Jan 10 22:22:13 2021
@author: wb511527
"""

from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes
from azure.cognitiveservices.vision.computervision.models import VisualFeatureTypes
from msrest.authentication import CognitiveServicesCredentials
from array import array
import os
from PIL import Image
import sys
import time
subscription_key = "your_key"
endpoint = "your_endpoint"
computervision_client = ComputerVisionClient(endpoint, CognitiveServicesCredentials(subscription_key))
print("===== Batch Read File - local =====")
# Get image of handwriting
directory="C:\\Users\\wb511527\\OneDrive - WBG\\Desktop\\New folder (3)"
answer=""
for image in os.listdir(directory):
     if image.endswith(".png"):

      local_image_handwritten_path = directory+"\\"+image
      print(local_image_handwritten_path)
      # Open the image
      local_image_handwritten = open(local_image_handwritten_path, "rb")

# Call API with image and raw response (allows you to get the operation location)
      recognize_handwriting_results = computervision_client.read_in_stream(local_image_handwritten, raw=True)
# Get the operation location (URL with ID as last appendage)
      operation_location_local = recognize_handwriting_results.headers["Operation-Location"]
# Take the ID off and use to get results
      operation_id_local = operation_location_local.split("/")[-1]

# Call the "GET" API and wait for the retrieval of the results
      while True:
          recognize_handwriting_result = computervision_client.get_read_result(operation_id_local)
          if recognize_handwriting_result.status not in ['notStarted', 'running']:
              break
          time.sleep(1)
      
# Print results, line by line
      if recognize_handwriting_result.status == OperationStatusCodes.succeeded:
        for text_result in recognize_handwriting_result.analyze_result.read_results:
            for line in text_result.lines:
                print(line.text)
                answer=answer+" "+line.text
            #print(line.bounding_box)
print()
import pandas as pd
os.chdir(directory)
df=pd.read_excel("base.xlsx")
pattern="(?:"+")|(?:".join(df["Ques"])+")"
import re
#pattern="(?:Enumerate the issues associated with functioning of tribunals in India)|(?:Examine the significance of Gram Sabhas, as mentioned in Article 243A of the Indian constitution)"
answers=re.split(pattern,answer)

lines=answer.split(".")
from fpdf import FPDF 
pdf = FPDF()   
for i in range(0,len(answers)-1):
       pdf.add_page()  
       pdf.set_font(family="Times",style ='B',size=10)
       pdf.cell(100, 6, txt = list(df["Ques"])[i], ln = 1, align = 'L') 
       pdf.set_font("Arial", size = 6) 
       lines=(answers[i+1]).split(".")
       for x in lines: 
          pdf.set_fill_color(175, 238, 238)
          if len(set(list(df["Answer Key"])[i].split("|")).intersection(set(x.split())))>=1:
             pdf.set_font(family="Times",style = '',size=8)
             pdf.cell(len(x)-3, 6, txt = x, ln = 1, align = 'L',fill=True) 
          else:
             pdf.set_font(family="Times",style = '',size=8)
             pdf.cell(200, 4, txt = x, ln = 1, align = 'L') 
       pdf.set_font(family="Times",style = '')
       pdf.cell(200, 4, txt = "Word Length = "+str(len((answers[i+1]).split(" "))), ln = 1, align = 'L')
       pdf.set_font(family="Times",style ='B',size=8)
       pdf.cell(100, 6, txt = "Expected Keywords :", ln = 1, align = 'L') 
       for m in list(df["Answer Key"])[i].split("|"):
           pdf.set_font(family="Times",style = '')
           pdf.cell(200, 4, txt = m, ln = 1, align = 'L')
       pdf.set_font(family="Times",style ='B',size=8)    
       pdf.cell(100, 6, txt = "Hit Keywords :", ln = 1, align = 'L') 
       for n in list(set(list(df["Answer Key"])[i].split("|")).intersection(set((answers[i+1]).split(" ")))):
           pdf.set_font(family="Times",style = '')
           pdf.cell(200, 4, txt = n, ln = 1, align = 'L')
       
pdf.output("myg.pdf")    


