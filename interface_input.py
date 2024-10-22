# %%
import ctypes
import time
from tkinter import *
from termcolor import colored
import tkinter as tk
import tkinter.messagebox as tkmsg
import tkinter.font as font
import win32gui, win32con

# Message Box
def message_box(type, message, title):
    try:
        root = Tk()
        root.withdraw()

        # Yes/No/Cancel Box
        if(type=='Y/N/C'):
            user_input = tkmsg.askyesnocancel(title, message)
            if(user_input != None):
                print(colored(f'{message}:', 'light_magenta'), colored(f'{user_input}', f'{bool_color(user_input)}'))
                return user_input
            else:
                print(colored(f'{message}:', 'light_magenta'), colored(f'Cancel', 'light_yellow'))
                return None


        # Yes/No Box  
        elif(type=='Y/N'):
            user_input = tkmsg.askyesno(title, message)
            if(user_input != None):
                print(colored(f'{message}:', 'light_magenta'), colored(f'{user_input}', f'{bool_color(user_input)}'))
                return user_input
            else:
                print(colored(f'{message}:', 'light_magenta'), colored(f'{user_input}', 'light_yellow'))
                return None
        

        # OK/Cancel Box  
        elif(type=='O/C'):
            user_input = tkmsg.askokcancel(title, message)
            if(user_input==True):
                print(colored(f'{message}:', 'light_magenta'), colored(f'OK', 'light_green'))
                return True
            elif(user_input==False):
                print(colored(f'{message}:', 'light_magenta'), colored(f'Cancel', 'light_red'))
                return None
            else:
                raise ValueError('Unknown User Input! ', user_input)
        
        # OK
        elif(type=='O'):
            user_input = tkmsg.showinfo(title, message)
            if(user_input=='ok'):
                return True
            else:
                raise ValueError('Unknown User Input! ', message + ': ' + user_input)
            
        elif(type=='warning'):
            tkmsg.showwarning(title, message)
            print(colored(f'\n{message}\n', 'light_yellow', attrs=["bold", "reverse"]))
            
        elif(type=='error'):
            tkmsg.showerror(title, message)
            print(colored(f'\n{message}\n', 'red', attrs=["bold", "reverse"]))
            
        else:
            raise ValueError('Unknown MessageBox Type Selected: ', type)
        

    except ValueError as err:
        print(err.args)

def get_console():
    return win32gui.GetForegroundWindow()

def hide_console(console):
    win32gui.ShowWindow(console, win32con.SW_HIDE)

def show_console(console):
    win32gui.ShowWindow(console, win32con.SW_SHOW)
    win32gui.SetForegroundWindow(console)


def top_console(console):
    win32gui.SetWindowPos(console, win32con.HWND_TOPMOST, 0,0,0,0, 
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def get_perf_time():
    return time.perf_counter()

def bool_color(bool):
    if(bool):
        return 'green'
    else:
        return 'red'