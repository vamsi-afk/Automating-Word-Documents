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
words = []

for word in doc.Words:
    if str(word) == "century" or str(word) == "Century":
        word.HighlightColorIndex = 7
