from pathlib import Path

import win32com.client as win32

def word2pdf(filedoc):
    """ convert the word *.doc, *.docx to *.pdf (save to same location)
    Args:
        filedoc (str): the file path for doc files
    """
    try:
        word = None
        doc = None
        if Path(filedoc).suffix in ['.doc', '.docx']:
            output = Path(filedoc).with_suffix('.pdf')
            if output.exists():
                return
            print("Convert WORD into PDF: "+filedoc)
            word = win32.DispatchEx('Word.Application')
            word.Visible = 0
            doc = word.Documents.Open(filedoc, False, False, True)
            # 'OutputFileName', 'ExportFormat', 'OpenAfterExport', 'OptimizeFor', 'Range',
            # 'From', 'To', 'Item', 'IncludeDocProps', 'KeepIRM', 'CreateBookmarks', 'DocStructureTags',
            # 'BitmapMissingFonts', 'UseISO19005_1', 'FixedFormatExtClassPtr'
            doc.ExportAsFixedFormat(
                str(output),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=1)
    except Exception as e:
        print('open failed due to \n' + str(e))
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()
    
if __name__ == '__main__':
    base_dir = '.'
    for doc_file in Path(base_dir).glob('*.doc?'):
        word2pdf(str(doc_file))
