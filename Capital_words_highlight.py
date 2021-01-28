import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\oof.docx")
words = []

for word in doc.Sentences:
    x = (str(word).partition(' ')[0])
    words.append(x)
for word_1 in doc.Words:
     t = re.findall('([A-Z][a-z]+)', str(word_1))
     if(t):
          if (t[0] not in words):
               print(t)
               word_1.HighlightColorIndex = 7
