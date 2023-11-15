from alive_progress import alive_it
from datetime import datetime
from pathlib import Path
import re
import shutil

def run():
    input_path = Path.cwd() / 'input'
    output_path = Path.cwd() / 'output'
    
    if not input_path.is_dir():
        input_path.mkdir()
    
    if not output_path.is_dir():
        output_path.mkdir()
    
    file_list = list(output_path.iterdir())
    if file_list:
        shutil.rmtree(output_path)
        output_path.mkdir()
    file_list = list(input_path.glob('*.mp4'))
    
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
        finalize=lambda bar: bar.text('Copying file: done'),
        **options
    )
    
    for file_path in results:
        pattern = r'\d+'
        text = file_path.name
        
        result = re.search(pattern, text)
        index = result.start()
        
        new_name = text[:index] + 'S1E' + text[index:]
        new_path = output_path / new_name
        
        results.text(f'Copying file: {new_name}')
        shutil.copy2(file_path, new_path)

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')