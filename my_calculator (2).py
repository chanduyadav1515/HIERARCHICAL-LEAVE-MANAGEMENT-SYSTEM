from tkinter import *
import tkinter
import tkinter.messagebox
def proces():
    number1=Entry.get(E1)
    number2=Entry.get(E2)
    operator=Entry.get(E3)
    number1=int(number1)
    number2=int(number2)
    E4["state"] = "normal"
    if operator =="+":
        answer=number1+number2
    if operator =="-":
        answer=number1-number2
    if operator=="*":
        answer=number1*number2
    if operator=="/":
        answer=number1/number2
    Entry.insert(E4,0,answer)
    b.grid_forget()
    print(answer)

def CloseWindow():
    calculator.destroy()

calculator = tkinter.Tk()
L1 = Label(calculator, text="My calculator",).grid(row=0,column=1)
L2 = Label(calculator, text="Number 1",).grid(row=1,column=0)
L3 = Label(calculator, text="Number 2",).grid(row=2,column=0)
L4 = Label(calculator, text="Operator",).grid(row=3,column=0)
L4 = Label(calculator, text="Answer",).grid(row=4,column=0)
E1 = Entry(calculator, bd =5)
E1.grid(row=1,column=1)
E2 = Entry(calculator, bd =5)
E2.grid(row=2,column=1)
E3 = Entry(calculator, bd =5)
E3.grid(row=3,column=1)
E4 = Entry(calculator, bd =5,state=DISABLED)
E4.grid(row=4,column=1)

b=Button(calculator, text ="Submit",command = proces)
b.grid(row=5,column=1,)



cl = Button(calculator, text = "Close", command = CloseWindow, state=NORMAL, anchor=W, justify=LEFT, padx=5,height = 100, width = 100)
cl.grid(row=6,column=1,)

calculator.mainloop()
