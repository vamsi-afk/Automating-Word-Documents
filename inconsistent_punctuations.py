import win32com.client as win32
import regex as re
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\test01.docx")
prob = ",,"
a = ["..", ",.", "--", "//", ",,", "/!", "??", "?!", "!!", "::", ":;", ";;", "?:", ":?", ",?", "?,", ".?", "?.", ".!", "__", "-!", "_!", "-?", "-.", "-,"]
b = [",:", ",;", "./", ".;", ";/", "/?", "/.", "/:", "/;", ";.", ";:", ".:", ":.", "?;"]
c = ["!:", "!;", "!.", ".!", ";!", ":!", "/!"]
# for word in doc.Words:
#     d = re.findall('r[”][’]', str(word))
#     if d:
#         print(d)
for words in doc.Sentences:
    d = re.findall(r'[”][’]', str(words))
    if d:
        print(d)
        words.HighlightColorIndex = 7
for words in doc.Words:
    x = str(words)
    x = x.strip()
    if x in a:
        print(words)
        words.HighlightColorIndex = 7
    if x in b:
        print(words)
        words.HighlightColorIndex = 7
    if x in c:
        print(words)
        words.HighlightColorIndex = 7
    if x in d:
        print(words)
        words.HighlightColorIndex = 7