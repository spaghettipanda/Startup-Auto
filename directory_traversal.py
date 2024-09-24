import sys
import os
import ctypes
import wget
import zipfile
from interface_input import message_box

def open_file(file_path):
    os.startfile(file_path) # Comment out if testing #
    print(f'File: {os.path.basename(file_path)}')

def create_folder(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)

# Function for traversing directory
# Requires directory as parameter
def traverse_dir(dir):
    try:
        path = dir
        print(f'\nTraversing {os.path.basename(dir)} Folder...\n')
        l_files = os.listdir(path)
        for file in l_files:
            # Variable for full file path
            file_path = f'{path}\\{file}'
            
            if os.path.isfile(file_path):
                try:
                    # Open the file pertaining to file_path
                    open_file(file_path)
                except:
                    # Catching if any error occurs and alerting the user
                    print(f'ALERT: {file} could not be printed! Please check the associated softwares, or the file type.')
            elif(os.path.isdir(file_path)):
                 print('\n--------------------------------------')
                 print(f'Folder \"{file}\" found! Entering folder...\n')
                 traverse_dir(f'{dir}\\{file}')
            else:
                print(f'ALERT: {file} is not a file, so can not be printed!')
    
    except FileNotFoundError as err:
        message_box('warning', f"""***************************************************************
            \n{os.path.basename(dir)} could not be found! 
            \n
            \n\"{dir}\"
            \n***************************************************************""", f'\"{os.path.basename(dir)}\" folder not found!')
        
        create_dir = message_box('Y/N', f"""Would you like a folder to be created for you?
        \n
        \n({dir})""", f'\"{os.path.basename(dir)}\" folder')
        
        if(create_dir == True):
            create_folder(dir)
            print(f'\nCreating folder at {dir}...\n')
        elif(create_dir == False):
            message_box('error', f"""Please try again after creating a folder at: 
            \n\"{dir}\"
            \nClick OK or press enter to exit...""", f'\"{os.path.basename(dir)}\" folder not found!')
            print('\nEnding execution...')
            sys.exit()
        else:
            print('\nEnding execution...')
            sys.exit()
    
    except Exception as err:
        print(f'\nAn unexpected error occurred! \nType: {sys.exc_info()} \nError:{err.args}')
        sys.exit()

def execute_command(cmd):
    try:
        print(f'\n--------------------------------------\nExecuting command: \'{cmd}\'\n--------------------------------------\n')
        os.system(f"{cmd}")  
    except Exception as err:
        print(f'\nCommand didn\'t work! \nType: {sys.exc_info()} \nError:{err.args}')
        sys.exit()

def get_data(EXTENDED_NAME_FORMAT: int):
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    data = EXTENDED_NAME_FORMAT
 
    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(data, None, size)
 
    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(data, nameBuffer, size)
    return nameBuffer.value

def get_username():
    return os.getlogin()

def download_to_dir(url, path):
    wget.download(url,out = path)

def extract_zip(zip_path, dir, file=None):
    with zipfile.ZipFile(zip_path, 'r') as zip:
        zip.extractall(dir)

def extract_zip(zip_path , dir, file_name):
    print(f'\nExtracting {file_name} from {dir}...\n')
    found = 0
    with zipfile.ZipFile(zip_path, 'r') as zip:
        contents = zip_contents(zip_path)
        print(f'Zip folder contains: {contents}')
        
        # Iterate over the file names
        for fileName in contents:
            if(fileName == file_name):
               print(f'\nExtracting {file_name} from {zip_path}...\n')
               zip.extract(fileName, dir)
               found = found+1
    if(found == 0):
        print(f'\nCould not find {file_name} in {dir}!...\n')
               

def zip_contents(file):
    with zipfile.ZipFile(file, 'r') as zip:
        return zip.namelist()

def remove_file(path):
    print(f'\nCleaning up {path}...\n')
    os.remove(path)