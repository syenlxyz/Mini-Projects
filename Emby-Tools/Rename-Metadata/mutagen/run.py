from alive_progress import alive_it
from datetime import datetime
from mutagen.mp4 import MP4
from pathlib import Path

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
        file = MP4(file_path)
        file['©nam'] = file_path.stem
        file['©cmt'] = file_path.stem
        file.save()
        
        results.text(f'Renaming metadata: {file_path.name}')
    return file_path

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')