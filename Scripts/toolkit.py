from ctypes import windll
import re
import win32com.client
from tkinter import ttk, Tk, HORIZONTAL
import logging
    
def mbox(title, text, style=0x40000):
    return windll.user32.MessageBoxW(0, text, title, style)
    
def integerize_row(lst, param):
    return([(k, v) if (k != param) else (param, int(v)) for (k, v) in lst])
    
def integerize_tbl(tbl, param):
    #converts the specified field to an integer
    return([integerize_row(row, param) for row in tbl])
    
def alphanum(str):
    #Converts a string to only alphanumeric characters without spaces
    return(re.sub(r'[^A-Za-z0-9]+', '', str))

def create_window(title):
    root = Tk()
    root.title(title)

    return root

def create_progressbar(root, label, rownum, determinate=True):

    ttk.Label(root,text=label).grid(row=rownum, column=0, padx=20, pady=10)
    if determinate:
        progress = ttk.Progressbar(root, orient=HORIZONTAL,mode='determinate',length = 400)
    else:
        progress = ttk.Progressbar(root, orient=HORIZONTAL,mode='indeterminate',length = 400)
    progress.grid(row=rownum,column=1, padx=(0,10),pady=10)
    root.update()
    
    return progress

def step_progressbar(progressbar, step):

    progressbar['value'] += step
    progressbar._root().update()

def activate_window(root):
    root.lift()

def error_occur(e):
    from os.path import dirname, abspath
    logpath = dirname(abspath(__file__)) + '\\Error.log'
    
    logging.basicConfig(filename=logpath, format='\n%(levelname)s - %(message)s')
    logging.exception(e)
    mbox('Error', 'Python error occured: {error}'.format(error=e))