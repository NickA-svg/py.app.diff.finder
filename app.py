#!/usr/bin/env python
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.constants import TRUE
from PIL import ImageTk
from main import main_gui

class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self) # create window
        self.title("Diff Finder Desktop")
        self.geometry('600x400')
        self.style = ttk.Style(self)

        filename = ImageTk.PhotoImage(file = "/Users/nick.aristidou@convexin.com/Documents/Projects/Python/py.diff.finder.gui/assets/background.png")
        background_label = tk.Label(self, image=filename)
        background_label.pack(fill='both',expand='yes')

        self.filename = "" # variable to store filename
        self.filename2 = "" # variable to store comparison filename

        ttk.Button(self, text='Browse',style="C.TButton", command= lambda: self.openfile(comparison= False),width=15).place(relx=0.5, rely=0.4, anchor='center')
        ttk.Button(self, text='Browse Comparison',style="C.TButton",command= lambda: self.openfile(comparison = True),width=15).place(relx=0.5, rely=0.5, anchor='center')
        ttk.Button(self, text='Compare',style="C.TButton",command= lambda: self.saveFile(path1=self.filename,path2=self.filename2),width=15).place(relx=0.5, rely=0.6, anchor='center')
        
        self.mainloop()

    def openfile(self,comparison):
        if comparison:
            self.filename2 = filedialog.askopenfilename(title="Open file")
        else:
            self.filename = filedialog.askopenfilename(title="Open file")

    def saveFile(self,path1,path2):
        self.files = [('Excel Workbook', '*.xlsx')]
        self.file = filedialog.asksaveasfile(filetypes = self.files, defaultextension = self.files)
        main_gui(path1,path2,self.file.name)

if __name__ == '__main__':
    App()

# in venv in terminal run python3 -m PyInstaller --onefile --windowed --icon=assets/app.ico app.py to deploy app
#ref https://mborgerson.com/creating-an-executable-from-a-python-script/