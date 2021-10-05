from tkcalendar import DateEntry
from datetime import date
import tkinter as tk
today = date.today()
root = tk.Tk()
d = DateEntry(root, mindate=today)
d.pack()

e = DateEntry(root, mindate=today)
e.pack()
root.mainloop()
