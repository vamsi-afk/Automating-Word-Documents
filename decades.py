import win32com.client as win32
import re

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\test1.docx")
for word in doc.Words:
    a = re.findall(r'twenties',str(word))
    b = re.findall(r'thirties', str(word))
    c = re.findall(r'fourties', str(word))
    d = re.findall(r'fifties', str(word))
    e = re.findall(r'sixties', str(word))
    f = re.findall(r'seventies', str(word))
    g = re.findall(r'eighties', str(word))
    h = re.findall(r'ninties', str(word))
    i = re.findall(r'hundreds', str(word))
    j = re.findall(r'twenties', str(word))
    k = re.findall(r'[0-9]+s', str(word))
    if a:
        print(a)
        word.HighlightColorIndex = 7
    if b:
        print(b)
        word.HighlightColorIndex = 7
    if c:
        print(c)
        word.HighlightColorIndex = 7
    if d:
        print(d)
        word.HighlightColorIndex = 7
    if e:
        print(e)
        word.HighlightColorIndex = 7
    if f:
        print(f)
        word.HighlightColorIndex = 7
    if g:
        print(g)
        word.HighlightColorIndex = 7
    if h:
        print(h)
        word.HighlightColorIndex = 7
    if i:
        print(i)
        word.HighlightColorIndex = 7
    if j:
        print(j)
        word.HighlightColorIndex = 7
    if k:
        print(k)
        word.HighlightColorIndex = 7