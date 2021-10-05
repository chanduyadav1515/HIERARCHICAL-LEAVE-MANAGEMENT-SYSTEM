# working in coordination with the file : C:\Users\sasi\Desktop\learning python\basics\button_led\src\main.cpp

from tkinter import *
import tkinter
import tkinter.messagebox
import serial  # add Serial library for Serial communication


def cload1_1():
    Arduino_Serial.write(str.encode('0'))  # send 1 to arduino


def cload1_0():
    Arduino_Serial.write(str.encode('1'))


def cload2_1():
    Arduino_Serial.write(str.encode('2'))


def cload2_0():
    Arduino_Serial.write(str.encode('3'))


def cload3_1():
    Arduino_Serial.write(str.encode('4'))


def cload3_0():
    Arduino_Serial.write(str.encode('5'))

def cload4_1():
    Arduino_Serial.write(str.encode('6'))


def cload4_0():
    Arduino_Serial.write(str.encode('7'))


def conc():
    str1 = str(e1.get())
    com_port = "com"+str1
    print(com_port)
    global Arduino_Serial
    Arduino_Serial = serial.Serial(com_port, 9600)  # Create Serial port object
    b2.destroy()
    lable.destroy()
    e1.destroy()
    call_number.geometry('1366x588')


call_number = tkinter.Tk()
call_number.title("LOADS ON / OFF")
call_number.geometry('1366x768+0+0')
call_number.configure(bg='WHITE')


loadoff = PhotoImage(file = r"C:\Users\sasi\Desktop\learning python\red_off2.png")
loadon = PhotoImage(file = r"C:\Users\sasi\Desktop\learning python\green_on2.png")

#LABELS FOR DISPLAYING LOADS

lable1=Label(call_number, text=" LOAD - 1", justify=LEFT, compound = LEFT,padx = 10,background = "white")
lable1.place(x = 50, y = 10, width=200, height=25)
lable1.config(font=("Courier", 20, 'bold'))

lable2=Label(call_number, text=" LOAD - 2", justify=LEFT, compound = LEFT,padx = 10,background = "white")
lable2.place(x = 350, y = 10, width=200, height=25)
lable2.config(font=("Courier", 20, 'bold'))

lable3=Label(call_number, text=" LOAD - 3", justify=LEFT, compound = LEFT,padx = 10,background = "white")
lable3.place(x = 645, y = 10, width=200, height=25)
lable3.config(font=("Courier", 20, 'bold'))


lable4=Label(call_number, text=" LOAD - 4", justify=LEFT, compound = LEFT,padx = 10,background = "white")
lable4.place(x = 950, y = 10, width=200, height=25)
lable4.config(font=("Courier", 20, 'bold'))

# for COM port number
lable=Label(call_number, text=" Enter Com Port Number : ", justify=LEFT, compound = LEFT,padx = 10,background = "white")
lable.place(x = 250, y = 650, width=450, height=50)
lable.config(font=("Courier", 20, 'bold'))


# BUTTONS FOR ON AND OFF
b1=Button(call_number, image = loadon, borderwidth = 0, command = cload1_1,background = "white", fg="white", width = 200, height = 200)
b1.grid(row=1,column=1,padx=50, pady=60)

b2=Button(call_number, command =cload1_0,  borderwidth = 0, image = loadoff, background = "white", fg="white", width = 200, height = 200)
b2.grid(row=5,column=1,padx=50, pady=60)

b3=Button(call_number, image = loadon, command = cload2_1, borderwidth = 0, background = "white", fg="white", width = 200, height = 200)
b3.grid(row=1,column=2,padx = 50,pady=20)

b4=Button(call_number, command =cload2_0, image = loadoff,  borderwidth = 0, background = "white", fg="white", width = 200, height = 200)
b4.grid(row=5,column=2, padx = 50, pady=20)

b5=Button(call_number, image = loadon, command = cload3_1, borderwidth = 0, background = "white", fg="white", width = 200, height = 200)
b5.grid(row=1,column=3,padx = 50, pady=20)

b6=Button(call_number, command =cload3_0, image = loadoff, borderwidth = 0, background = "white", fg="white", width = 200, height = 200)
b6.grid(row=5,column=3, padx = 50, pady=20)

b7=Button(call_number, image = loadon, command = cload4_1, borderwidth = 0, background = "white", fg="white", width = 200, height = 200)
b7.grid(row=1, column=4, padx=50, pady=20)

b8 = Button(call_number, command=cload4_0, image=loadoff, borderwidth=0, background="white", fg="white", width=200, height=200)
b8.grid(row=5, column=4, padx=50, pady=20)

# to get com port details-------------------
e1 = Entry(call_number)
e1.place(x=720, y=650, width=50, height=50)
e1.config(font=("Courier", 20, 'bold'))
e1.focus_set()

b2=Button(call_number,text ="SUBMIT", command =conc, borderwidth = 2, background = "white", fg="black", width = 20, height = 20)
b2.place(x = 800, y = 650, width=100, height=50)

call_number.mainloop()
