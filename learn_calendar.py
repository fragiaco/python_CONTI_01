import calendar
from tkinter import *

c = calendar.TextCalendar(calendar.MONDAY)
c.prmonth(2009, 2)

print(calendar.monthrange(2009, 2)) #(6, 28) (0 è lunedì, il 6 è la domenica)
#print(calendar.calendar(2009))

#print (calendar.TextCalendar(calendar.SUNDAY).formatyear(2007, 2, 1, 1, 2))

root = Tk()
frame = Frame(root).pack()
cal_content = calendar.monthrange(2009,2)



# Create a label for showing the content of the calendar
cal_year = Label(frame, text=cal_content[0], font="Consolas 10 bold").pack()

root.mainloop()

# Import Required Library
from tkinter import *
from tkcalendar import Calendar

# Create Object
root = Tk()

# Set geometry
root.geometry("400x400")

# Add Calendar
cal = Calendar(root, selectmode='day',
               year=2020, month=5,
               day=22)

cal.pack(pady=20)


def grad_date():
    date.config(text="Selected Date is: " + cal.get_date())


# Add Button and Label
Button(root, text="Get Date",
       command=grad_date).pack(pady=20)

date = Label(root, text="")
date.pack(pady=20)

# Execute Tkinter
root.mainloop()