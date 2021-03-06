import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\test01.docx")
words = []

for sentence in doc.Paragraphs:
     t = re.findall("\w[.]\w[.][,]", str(sentence))
     if(t):
          print(t)
for sentence in doc.Sentences:
     x = re.findall("e[.]g",str(sentence))
     if x:
          print(x)
          # word.HighlightColorIndex = 7
for word in doc.Paragraphs:
     x = re.findall("e[.]g[.]",str(word))
     if x:
          print(x)
for word in doc.Sentences:
     t = re.findall('for example',str(word))
     t1 = re.findall('that is',str(word))
     if t:
          print(t)
          # word.HighlightColorIndex = 7
     if t1:
          print(t1)
          # word.HighlightColorIndex = 7
