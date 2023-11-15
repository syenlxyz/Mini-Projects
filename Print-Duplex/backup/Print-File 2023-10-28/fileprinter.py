from pyoffice.filesystem import FileSystem
from pyoffice.fileoptimizer import getExecutionTime
from win32com.client import Dispatch

class FilePrinter():
    def __init__(self):
        file_system = FileSystem()
        input_path = file_system.getInputPath()
        file_system.getFileCounter(input_path)
        self.printFiles(input_path)

    def printFiles(self, input_path):
        doc_ext = ['.docx',  '.doc']
        other_ext = ['.htm',  '.html',  '.pdf',  '.txt',  '.jpeg',  '.jpg',  '.png']
        print('====================================================================')
        file_list = [path for path in input_path.glob('**/*') if path.is_file()]
        for index, file_path in enumerate(file_list):
            print(f'Printing {file_path.name} ({index + 1}/{len(file_list)})')
            try:
                if file_path.suffix in doc_ext:
                    exec_time = getExecutionTime(self.doc2Print, path=file_path)
                    print(f'Execution time: {exec_time}')
                elif file_path.suffix in other_ext:
                    exec_time = getExecutionTime(self.other2Print, path=file_path)
                    print(f'Execution time: {exec_time}')
                else:
                    print(f'{file_path.suffix} is not supported')
            except:
                try:
                    exec_time = getExecutionTime(self.other2Print, path=file_path)
                    print(f'Execution time: {exec_time}')
                except:
                    print('Conversion failed')
    
    def doc2Print(self, path):
        app = Dispatch('Word.Application')
        app.Visible = False
        doc = app.Documents.open(str(path), ReadOnly=True)
        doc.PrintOut()
        doc.Close()
        app.Quit()
    
    def other2Print(self, path):
        app = Dispatch('AcroExch.App')
        app.Hide()
        avDoc = Dispatch('AcroExch.AVDoc')
        avDoc.Open(str(path), '')
        pdDoc = avDoc.GetPDDoc()
        avDoc.PrintPages(0, pdDoc.GetNumPages() - 1, nPSLevel=2, bBinaryOk=0, bShrinkToFit=0)
        avDoc.Close(True)
        app.MenuItemExecute('Quit')