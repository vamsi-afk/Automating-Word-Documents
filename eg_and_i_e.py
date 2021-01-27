import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\oof.docx")
words = []

for word in doc.Paragraphs:
     t = re.findall("\w[.]\w[.][,]", str(word))
     if(t):
          print(t)
          # word.HighlightColorIndex = 7
for word in doc.Sentences:
     t = re.findall('for example',str(word))
     t1 = re.findall('that is',str(word))
     if t:
          print(t)
          word.HighlightColorIndex = 7
     if t1:
          print(t1)
          word.HighlightColorIndex = 7