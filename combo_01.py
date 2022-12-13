import tkinter as tk
from tkinter import ttk
my_w = tk.Tk()
my_w.geometry("300x150")  # Size of the window
my_w.title("www.plus2net.com")  # Adding a title
def my_upd1():
    cb1.set('Apr') # update selection to Apr
    l1.config(text=cb1.get()+':'+ str(cb1.current())) # value & index
###
'''
def my_upd1():
    cb1.set('') # Clear the selection 
    cb1.delete(0,'end') # clear the selection
    l1.config(text=cb1.get()+':'+ str(cb1.current())) # value & index
'''
###
months=['Jan','Feb','Mar','Apr','May','Jun']
cb1 = ttk.Combobox(my_w, values=months,width=7)
cb1.grid(row=1, column=1,padx=10,pady=20)

b1=tk.Button(my_w, text="set('Apr')", command=lambda: my_upd1())
b1.grid(row=1, column=2)

l1=tk.Label(my_w, text='Month')
l1.grid(row=1, column=3)
print(cb1.get())
my_w.mainloop()  # Keep the window open