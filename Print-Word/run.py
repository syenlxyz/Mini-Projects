from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

def run():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        input_path.mkdir()
    
    file_list = list(input_path.glob('*.docx'))
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }
    
    results = alive_it(
        file_list, 
        len(file_list), 
        finalize=lambda bar: bar.text('Printing Word Document: done'),
        **options
    )
    
    app = Dispatch('Word.Application')
    app.Visible = False
    for file_path in results:
        results.text(f'Printing Word Document: {file_path.name}')
        doc = app.Documents.Open(str(file_path), ReadOnly=True)
        doc.PrintOut()
        doc.Close()
    app.Quit()

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')