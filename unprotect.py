import win32com.client as win32

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
doc = word.Documents.Open("C:\\Users\\pc\\Desktop\\Word_Automations\\protected.docx")

password= input("Please enter the password: ")

doc.Unprotect(password)                      
print("Protection Removed!")

doc.SaveAs("C:\\Users\\pc\\Desktop\\Word_Automations\\unprotected.docx")