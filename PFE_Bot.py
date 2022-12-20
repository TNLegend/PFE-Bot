from tkinter import filedialog
import os
import win32com.client
import comtypes.client
from email.message import EmailMessage
import mimetypes
import smtplib
print('''

██████  ███████ ███████     ██████   ██████  ████████ 
██   ██ ██      ██          ██   ██ ██    ██    ██    
██████  █████   █████       ██████  ██    ██    ██    
██      ██      ██          ██   ██ ██    ██    ██    
██      ██      ███████     ██████   ██████     ██    
                                                      
Created with love By TNLegend
''')

answer = int(input("Type 1 to import your resume and follow up letter\nType 2 to exit\n>>"))
if answer == 2:
    exit(0)
else:
    #preparing cv and resume file
    path = str(filedialog.askopenfilename(title="select your cv and resume",filetypes=[("doc file",".doc*")]))
    path = path.replace("/","\\")
    print(path)
    toReplace = input("please type the text you want to replace >>")
    wd_replace = 2  # 2=replace all occurences, 1=replace one occurence, 0=replace no occurences
    wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached
    path2 = str(filedialog.askopenfilename(title="select emails list",filetypes=[("text file", ".txt")]))
    print(path2)
    sender = input("type your email >>")
    appPass = input("type your email password >>")
    subject = input("type email subject >>")
    body = input("type what you want to send in all mails body >>")
    with open(path2,mode="r",encoding="utf-8") as mails:
        lines = mails.readlines()
    for line in lines:
        myList = line.split(":")
        receiver = myList[0]
        replaced = myList[1]

        # Open Word
        word_app = win32com.client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        word_app.Documents.Open(fr"{path}")
        # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
        word_app.Selection.Find.Execute(
            FindText=toReplace,
            ReplaceWith=replaced,
            Replace=wd_replace,
            Forward=True,
            MatchCase=True,
            MatchWholeWord=False,
            MatchWildcards=True,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Wrap=wd_find_wrap,
            Format=True,
        )
        # Save the new file
        word_app.ActiveDocument.SaveAs(str(f"{os.getcwd()}/final.docx"))
        word_app.ActiveDocument.Close(SaveChanges=False)
        word = comtypes.client.CreateObject('Word.Application')
        word.Quit()
        word_app.Application.Quit()
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(str(f"{os.getcwd()}/final.docx"))
        wdFormatPDF = 17
        doc.SaveAs(str(f"{os.getcwd()}/final.pdf"), FileFormat=wdFormatPDF)
        doc.Close()
        os.remove(f"{os.getcwd()}/final.docx")
        message = EmailMessage()
        message['from'] = sender
        message['to'] = receiver
        message['subject'] = subject
        message.set_content(body)
        mime_type, _ = mimetypes.guess_type(f"{os.getcwd()}/final.pdf")
        mime_type, mime_subtype = mime_type.split('/', 1)
        with open("final.pdf", 'rb') as ap:
            message.add_attachment(ap.read(), maintype=mime_type, subtype=mime_subtype,filename=os.path.basename(f"{os.getcwd()}/final.pdf"))
        with smtplib.SMTP(host="smtp.gmail.com",port=587) as smtp:
            smtp.ehlo() #start communiction with smtp server
            smtp.starttls() #tls = transport layer security
            smtp.login(sender,appPass)
            smtp.send_message(message)
            print(f"email sent to {receiver}")
