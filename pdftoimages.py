# -*- coding: utf-8 -*-
"""
Created on Mon Jan 25 17:37:06 2021

@author: wb511527
"""

import fitz
import os
os.chdir(r"C:\Users\wb511527\OneDrive - WBG\Desktop\New folder (3)")

doc = fitz.open('be159-183643_1049_jatin-kishore_rank_2.pdf')

doc.pageCount
for page in doc:
   pix = page.getPixmap()
   pix.writeImage("%i.png" % page.number)
