import shutil
import os

'''
    #Need file template
    it gonna copy the template as much as tickets/folder 
'''

folder_path = ('../COPY_RENAMEFILE')
path = os.listdir(folder_path)

Ticket_folder = []
Ticket_to_copy = 'AOS-00000 Publish Report.xlsx'

#read just folder and ignoring file
for name in path:
    # Join the folder path and the name to get the full path
    full_path = os.path.join(folder_path, name)
    # Check if the full path is a directory
    if os.path.isdir(full_path):
        Ticket_folder.append(name)
        

for channel in Ticket_folder:
    shutil.copy(Ticket_to_copy,f'AOS-{channel} Publish Report.xlsx')
