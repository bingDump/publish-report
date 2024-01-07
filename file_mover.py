import shutil
import os
import Publish_report

path = Publish_report.path
file_channel = Publish_report.Channels.keys()


class file_move:
    def __init__(self,path) -> None:
        self.path = path
        
    def folder(self,path_file):
        folder_list = []
        folder = os.listdir(path_file)
        for fold in folder:
            full_path = os.path.join(path_file,fold)
            if os.path.isdir(full_path):
                folder_list.append(fold)
        return folder_list
    
    def move(self):
        folder_list = file_mover.folder(self.path)
        for folder in folder_list:
            sub_folders = file_mover.folder(folder)
            for sub_fold in sub_folders:
                if len(sub_fold) == 0:
                    continue
                else:
                    file_path = f'{path}/{folder}/{sub_fold}'
                    files = os.listdir(file_path)
                    for file in files:
                        for chanel in file_channel:
                            if chanel in file:
                                path_to =f'{path}/{folder}'
                                shutil.move(f'{file_path}/{file}',f'{path_to}/{file}')
                            else:
                                continue 
        return None                   

file_mover = file_move(path)
