import win32com.client as win32
from colorama import init, Fore
import regex as re
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\test1.docx")


for word in doc.Sentences:
    x = re.findall(r'[A-Za-z]+[%]', str(word))
    y = re.findall(r'[A-Za-z]+\s+[%]', str(word))
    z = re.findall(r'[0-9]+[%]', str(word))
    a = re.findall(r'[0-9]\s[%]', str(word))
    if x:
        print(x)
        word.HighlightColorIndex = 7
    if y:
        print(y)
        word.HighlightColorIndex = 7
    if z:
        print(z)
        word.HighlightColorIndex = 7
    if a:
        print(a)
        word.HighlightColorIndex = 7
for word in doc.Words:
    s = re.findall(r'Percent', str(word))
    d = re.findall(r'percent', str(word))
    if d:
        print(d)
        word.HighlightColorIndex = 7
    if s:
        print(s)
        word.HighlightColorIndex = 7