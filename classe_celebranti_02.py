##############################################################
##########TOP WINDOW CELEBRANTI###############################
##############################################################

import tkinter as tk
from tkinter import ttk

class SettingsWindows:
    def __init__(self):
        self.win = tk.Toplevel() # simile a root = tk.Tk()
        self.frame = tk.Frame(self.win)
        self.frame.pack(padx=5, pady=5)

        self.label = tk.Label(self.frame, text= 'This is a new window')
        self.label.pack(padx=5, pady=5)

        self.radio = tk.Radiobutton(self.frame, text='Option 1', value=1)
        self.radio.pack(padx=20, pady=20)

        self.radio2 = tk.Radiobutton(self.frame, text='Option 2', value=2)
        self.radio2.pack(padx=20, pady=20)


class MainWindow:
    def __init__(self, master):
        self.master = master

        self.frame = tk.Frame(self.master, width=300, height=200, background='lightgrey')
        self.frame.pack()

        self.botton = tk.Button(self.frame, text='Click_me', command=self.c1)
        self.botton.place(x=50, y=50)

        self.botton2 = tk.Button(self.frame, text='Min_Max Settings', command=self.c2)
        self.botton2.place(x=50, y=100)

        self.flag = 0

    def c1(self):
        self.settings = SettingsWindows()
        self.settings.win.mainloop()

        self.flag = 1

    def c2(self):
        if  self.flag == 1: #if the window is showing
                self.settings.win.withdraw()
                self.flag = 0
        elif self.flag == 0:
                self.settings.win.deiconify()
                self.flag = 1

root = tk.Tk()
window = MainWindow(root)
root.mainloop()