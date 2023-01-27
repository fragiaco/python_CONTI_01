from tkinter import *

def up():
    number.set(number.get()+1)

def down():
    number.set(number.get()-1)

window = Tk()
window.title("Programme")
window.geometry('350x250')
number = IntVar()

frame = Frame(window)
frame.pack()

entry = Entry(frame, textvariable=number, justify='center')
entry.pack(side=LEFT, ipadx=15)

buttonframe = Frame(entry)
buttonframe.pack(side=RIGHT)

buttonup = Button(buttonframe, text="▲", font="none 5", command=up)
buttonup.pack(side=TOP)

buttondown = Button(buttonframe, text="▼", font="none 5", command=down)
buttondown.pack(side=BOTTOM)

window.mainloop()

def main():
    root = tk.Tk()
    root.title('ButtonEntryCombo')
    root.resizable(width=tk.NO, height=tk.NO)
    app = App(root)
    root.mainloop()

main()