from alive_progress import alive_it
from datetime import datetime
from mutagen.mp4 import MP4
from pathlib import Path
import re

def run():
    input_path = Path.cwd() / 'input'

    folder_list = [path for path in input_path.iterdir() if path.is_dir()]
    options = {
        'length': 70,
        'spinner': 'classic',
        'bar': 'classic2',
        'receipt_text': True,
        'dual_line': True
    }

    results = alive_it(
        folder_list,
        len(folder_list),
        finalize=lambda bar: bar.text('Renaming files: done'),
        **options
    )
    
    for folder_path in folder_list:
        folder_name = folder_path.name
        results.text(f'Renaming files: {folder_name}')
        
        file_list = list(folder_path.glob('*.mp4'))
        for file_path in file_list:
            update_title(file_path)
        
        pattern = r'\[(.*?)\]'
        result = re.findall(pattern, folder_path.name)
        english = result[0]
        chinese = result[1]
        subtitle = result[-3]
        resolution = result[-2]
        
        season = get_season(folder_path)
        episodes = get_episodes(file_list)
        
        num_digit = len(episodes[0])
        season = season.zfill(num_digit)
        
        for episode in episodes:
            if episode:
                file_name = f'{english}.{chinese}.S{season}E{episode}.{subtitle}.{resolution}.mp4'
                target = file_path.parent / file_name
                file_path.rename(target)

def update_title(file_path):
    file = MP4(file_path)
    title = ''.join(file['©nam'])
    
    if not title:
        file['©nam'] = file_path.stem
        file['©cmt'] = file_path.stem
        file.save()

def get_season(folder_path):
    pattern = r'\[S\d+\]'
    result = re.search(pattern, folder_path.name)
    if result:
        text = result.group()
        search = re.search(r'\d+', text)
        season = search.group()
    else:
        season = '1'
    return season

def get_episodes(file_list):
    episodes = []
    for file_path in file_list:
        keywords = ['.5', 'OVA']
        for keyword in keywords:
            if keyword in file_path.stem:
                episode = None
                break
        else:
            episode = get_episode(file_path)
        episodes.append(episode)
    
    start = min([int(episode) for episode in episodes if episode])
    if start > 1:
        episodes = [ifelse(episode, str(int(episode) - start + 1), episode) for episode in episodes]
    
    num_digit = max([len(episode) for episode in episodes if episode])
    episodes = [ifelse(episode, episode.zfill(num_digit), episode) for episode in episodes]
    return episodes

def get_episode(file_path):
    patterns = [
        r'\[E\d+\]',
        r'\[\d+(v\d)?\]',
        r'\[\d+\s?END\]',
        r'\[\d+_\d+\]',
        r'-\s?\d+',
        r'\.\d+',
        r'第\d+话',
        r'第\d+話'
    ]
    
    for pattern in patterns:
        result = re.search(pattern, file_path.stem)
        if result:
            text = result.group()
            search = re.search(r'\d+', text)
            episode = search.group()
            return episode

def ifelse(test_expression, x, y):
    if test_expression:
        return x
    else:
        return y

if __name__ == '__main__':
    start_time = datetime.now()
    run()
    end_time = datetime.now()
    run_time = end_time - start_time
    print(f'Execution time: {run_time}')