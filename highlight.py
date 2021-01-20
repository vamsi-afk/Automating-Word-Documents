import win32com.client as win32

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True
doc = word.Documents.Open("fullpathofyourdoc.docx")

Arr=[] //insert the words u want to highlight in yellow.

for word_t in doc.Words:
	ele = str(word_t)
	ele = ele.strip()
	ele = ele.lower()
	if ele in Arr:
		word_t.HighlightColorIndex=7



		
