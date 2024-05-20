from pathlib import Path
import os
from datetime import datetime
import pandas as pd

def get_dir_df(root_path: str) -> pd.DataFrame:
    
    ROOT_PATH = Path(root_path)

    dir_info = {
        'name'          : [],
        'path'          : [],
        'created_at'    : [],
        'last_accessed' : [],
        'last_modified' : [],
        'filetype'      : []
    }

    with os.scandir(ROOT_PATH) as root_directory:
        for item in root_directory:
            
            item_stats = item.stat()
            dir_info['name'].append(item.name)
            dir_info['path'].append(item.path)
            dir_info['created_at'].append(datetime.fromtimestamp(item_stats.st_ctime))
            dir_info['last_accessed'].append(datetime.fromtimestamp(item_stats.st_atime))
            dir_info['last_modified'].append(datetime.fromtimestamp(item_stats.st_mtime))
            
            filetype = os.path.splitext(item.path)[-1]
            dir_info['filetype'].append('dir') if item.is_dir() else dir_info['filetype'].append(filetype)
    
    dir_df = pd.DataFrame(dir_info)
    return dir_df
    
def get_subdir_df(dir_df: pd.DataFrame) -> pd.DataFrame:
    
    subdir_info = {
        'name'          : [],
        'dir'           : [],
        'path'          : [],
        'created_at'    : [],
        'last_accessed' : [],
        'last_modified' : [],
        'filetype'      : []
        }

    for path in dir_df['path']:
        
        ITEM_PATH = Path(path)
        
        with os.scandir(ITEM_PATH) as item_directory:
            for item in item_directory:
                
                item_stats = item.stat()
                subdir_info['name'].append(item.name)
                subdir_info['dir'].append(ITEM_PATH)
                subdir_info['path'].append(ITEM_PATH.name)
                subdir_info['created_at'].append(datetime.fromtimestamp(item_stats.st_ctime))
                subdir_info['last_accessed'].append(datetime.fromtimestamp(item_stats.st_atime))
                subdir_info['last_modified'].append(datetime.fromtimestamp(item_stats.st_mtime))
                
                filetype = os.path.splitext(item.path)[-1]
                subdir_info['filetype'].append('dir') if item.is_dir() else subdir_info['filetype'].append(filetype)
                
    
    proj_df = pd.DataFrame(subdir_info)
    return proj_df

if __name__ == '__main__':
    
    root_path = input('Please input the path to analise: ')
    
    try:
        
        directory = get_dir_df(root_path = root_path)
        subdirectories = get_subdir_df(dir_df = directory)
            
        with pd.ExcelWriter('files.xlsx', engine = 'xlsxwriter') as writer:
        
            directory.to_excel(
                excel_writer = writer,
                header = directory.columns,
                sheet_name = 'items',
                index_label = 'id')
            
            subdirectories.to_excel(
                excel_writer = writer,
                header = subdirectories.columns,
                sheet_name = 'subitems',
                index_label = 'id')
            
        print('Excel generated succesfully')
    
    except Exception as e:
        print(e)