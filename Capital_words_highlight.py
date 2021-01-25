import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\oof.docx")
words = []

for word in doc.Words:
     t = re.findall('([A-Z][a-z]+)', str(word))
     if(t):
         word.HighlightColorIndex = 7
