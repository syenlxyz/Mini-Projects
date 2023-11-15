import os
import win32api
from pathlib import Path
from datetime import datetime

def print_pdf(file_path):
    win32api.ShellExecute(0, 'print', str(file_path), None, '.', 0)
    
def main():
    input_path = Path.cwd() / 'input'
    if not input_path.is_dir():
        raise Exception('Sorry, the input path does not exist.')
    if not list(input_path.iterdir()):
        raise Exception('Sorry, the input path is empty.')
    print('--------------------------------------------------------------------')
    ext = 'pdf'
    print(f'Supported File Format: {ext}')
    file_list = list(input_path.glob(f'**/*.{ext}'))
    num_file = len(file_list)
    print(f'Number of .{ext} files found: {num_file}')
    for index, file_path in enumerate(file_list):
        print(f'Printing {file_path.name} ({index + 1}/{num_file})')
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