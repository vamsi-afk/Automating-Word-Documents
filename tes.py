import win32com.client as win32
import regex as re

wdFindContinue = 1
wdReplaceAll = 2

word = win32.gencache.EnsureDispatch('Word.Application')
word.DisplayAlerts = 0
word.Visible = True
password= input("Please enter the password: ")
#
doc = word.Documents.Open("C:\\Users\\VAMSI\\Desktop\\as\\te.docx")
# doc.Protect(0, True, password, False, False)
# protecttype = doc.ProtectionType
z = "\""      # Straight quotes
a = '“'       # Open curly quotes
b = '”'       # Close curly quotes
for sentence in doc.Sentences:  # Going through each sentence
      x = re.findall(r'["]', str(sentence)) # Finding all the occurences of straight quotes
      if x:
          print(sentence)
          for i in x:
           word.Selection.Find.Execute(str(i), False, False, False, False, False, True, wdFindContinue, False,
                                      b, wdReplaceAll)
           exit(0)
