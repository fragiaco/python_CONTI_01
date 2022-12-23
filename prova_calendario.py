from tkinter import *
from tkcalendar import *

screen=Tk()
screen.minsize(800, 600)
screen.configure(background='blue')

def selectDate():
    myDate=myCal.get_date()
    selectedDate =Label(text=myDate)
    selectedDate.place(x=425, y=350)

myCal = Calendar(screen, setmode='day', date_pattern='d/m/yy')
myCal.place(x=360, y=100)
openCal = Button(screen, text='select date', command=selectDate)
openCal.place(x=425, y=300)
screen.mainloop()