

from tkinter import *
from pandastable import Table

root = Tk()
frame = Frame(root)
frame.pack()

table = Table(frame, showtoolbar=True, showstatusbar=True)
table.show()



root.mainloop()

