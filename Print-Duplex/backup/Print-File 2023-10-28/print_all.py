import os
import win32api
import numpy as np
from pathlib import Path
from win32com import client
from datetime import datetime

def check_dir(input_path):
    folder_list = []
    file_list = []
    ext_list = []
    for path in input_path.glob('**/*'):
        if path.is_dir():
            folder_list.append(path)
        if path.is_file():
            file_list.append(path)
            ext_list.append(path.suffix)
    print(f'Total number of folders: {len(folder_list)}')
    print(f'Total number of files: {len(file_list)}')
    print('--------------------------------------------------------------------')
    ext_name, ext_count = np.unique(ext_list, return_counts=True)
    ext_counter = dict(zip(list(ext_name), list(ext_count)))
    for ext in ext_counter:
        print(f'Number of {ext} files: {ext_counter[ext]}')

def print_word(file_path):
    word = client.Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(str(file_path), ReadOnly=True)
    doc.PrintOut()
    doc.Close()

def print_pdf(file_path):
    win32api.ShellExecute(0, 'print', str(file_path), None, '.', 0)

def main():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        raise Exception('Sorry, input path does not exist.')
    if not list(input_path.iterdir()):
        raise Exception('Sorry, input path is empty.')
    print('--------------------------------------------------------------------')
    print('Checking input folder...')
    check_dir(input_path)
    target_ext = ['.doc', '.docx', '.pdf']
    print(f'Supported File Formats: {", ".join(target_ext)}')
    for ext in target_ext:
        print('--------------------------------------------------------------------')
        print(f'Working on {ext} files...')
        file_list = list(input_path.glob(f'**/*{ext}'))
        for index, file_path in enumerate(file_list):
            print(f'Printing {file_path.name} ({index + 1}/{len(file_list)})')
        if ext in ['.doc', '.docx']:
            print_word(file_path)
        elif ext == '.pdf':
            print_pdf(file_path)
            
if __name__ == '__main__':
    print('Program Started')
    start_time = datetime.now()
    main()
    end_time = datetime.now()
    print('==========================End of Program============================')
    exec_time = end_time - start_time
    print(f'Total execution time: {exec_time}')
    input('Press Enter to continue...')
    os.system('taskkill /f /im Acrobat.exe')
