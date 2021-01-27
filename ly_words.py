import win32com.client as win32
import re
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\krupa\\Desktop\\20101.docx")


result = []
result2 = []
for word in doc.Sentences:
    result1 = re.findall(r'\w+(?:ly)+', str(word))
    result = result + result1
print(result)
