import win32com.client as win32
import os, sys

def word2pdf(filedoc):
    print(filedoc)
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = 0
        test = word.Documents.Open(filedoc)
        if filedoc.endswith('.doc'):
            test.ExportAsFixedFormat(filedoc[:-3]+'pdf',17)
        elif filedoc.endswith('.docx'):
            test.ExportAsFixedFormat(filedoc[:-4]+'pdf',17)
        test.Close()
        word.Quit()
    except:
        print('open failed')
    
if __name__ == '__main__':
    for filetuple in os.walk('.\\'):
        for datefile in filetuple[2]:
            if datefile.endswidth('.doc'):
                word2pdf(str(os.getcwd())+'\\'+datefile)
            elif datefile.endswidth('.docx'):
                word2pdf(str(os.getcwd())+'\\'+datefile)