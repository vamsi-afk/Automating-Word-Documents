import docx
import os
from docx.enum.text import WD_COLOR_INDEX
import win32com.client as win32
import re

import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\oof.docx")
words = ["one","two","three","four","five","six","seven","eight","nine","ten","eleven","twelve","thirteen","fourteen","teen","fifteen","sixteen","seventeen","eighteen","nineteen","twenty","thirty","fourty","fifty","sixty","seventy","eighty","ninty","hundred","thousand","thousands","lakh","lakhs","crore","crores","million","millions","billion", "billions", "tens","twentys","zero","zeros"]
words_1 = ["thousands","lakhs","crores","millions", "billions","trillion","trillions", "tens","twentys","zero","zeros"]

for word in doc.Words:
    ele = str(word)
    ele = ele.lower()
    ele = ele.strip()
    if str(ele) in words:
        word.HighlightColorIndex = 7
    if ele in words_1:
        word.HighlightColorIndex = 7