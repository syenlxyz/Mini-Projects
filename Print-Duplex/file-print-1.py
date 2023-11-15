from win32com.client import Dispatch

class FilePrinter():
    def printPdf(self, path):
        app = Dispatch('AcroExch.App')
        app.Hide()
        avDoc = Dispatch('AcroExch.AVDoc')
        avDoc.Open(str(path), '')
        pdDoc = avDoc.GetPDDoc()
        nFirstPage = 0
        nLastPage = pdDoc.GetNumPages() - 1
        nPSLevel = 2
        bBinaryOk = 0
        bShrinkToFit = 0
        avDoc.PrintPages(nFirstPage, nLastPage, nPSLevel, bBinaryOk, bShrinkToFit)
        app.MenuItemExecute('Quit')

    def printWord(self, path):
        app = Dispatch('Word.Application')
        app.Visible = False
        doc = app.Documents.Open(str(path), ReadOnly=True)
        Background = True
        Append = False
        Range = 3
        OutputFileName = ''
        From = str(1)
        To = str(doc.ComputeStatistics(2))
        doc.PrintOut(Background, Append, Range, OutputFileName, From, To)
        doc.Close()

    def printFiles(self, input_path):
        target_ext = ['.doc', '.docx', '.pdf']
        print('====================================================================')
        print(f'Supported file formats: {", ".join(target_ext)}')
        for ext in target_ext:
            print('--------------------------------------------------------------------')
            file_list = list(input_path.glob(f'**/*{ext}'))
            if file_list:
                print(f'Working on {ext} file(s)')
                print('--------------------------------------------------------------------')
                for index, file_path in enumerate(file_list):
                    print(f'Printing {file_path.name} ({index + 1}/{len(file_list)})')
                    if ext in ['.doc', '.docx']:
                        self.printWord(file_path)
                    elif ext == '.pdf':
                        self.printPdf(file_path)
            else:
                print(f'There is no {ext} file(s) available')