from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
import subprocess

def run():
    input_path = Path.cwd() / 'input'
    file_list = list(input_path.glob('**/*.mp4'))
    
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
        finalize=lambda bar: bar.text('Renaming metadata: done'),
        **options
    )
    
    for file_path in results:
        results.text(f'Renaming metadata: {file_path.name}')
        subprocess.run(['exiftool', f'-title={file_path.stem}', '-quiet', '-overwrite_original', file_path.as_posix()])
        subprocess.run(['exiftool', f'-comment={file_path.stem}', '-quiet', '-overwrite_original', file_path.as_posix()])
    return file_path

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')