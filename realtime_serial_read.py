import serial
import tkinter
from tkinter import *
class SerialViewer:
    def __init__(self):
        self.win = Tk()
        self.ser = serial.Serial('com5',9600)

    def bt1 (self):
        self.ser.write('on')

    def bt2 (self):
        self.ser.write('off')

    def bt3 (self):
        self.ser.write(self.v.get())

    def bt4 (self):
        self.ser.write(self.v.get())

    def makewindow (self):
        frame1 = Frame(self.win)
        frame1.pack(side = LEFT)
        b1 = Button(frame1, text = "ON", command = self.bt1)
        b2 = Button(frame1, text = "OFF", command = self.bt2)
        b1.grid(row = 0, column = 0)
        b2.grid(row = 0, column = 1)

        frame2 = Frame(self.win)
        frame2.pack()
        self.v = StringVar()
        r1 = Radiobutton(frame2,text = 'on', variable = self.v, value = 'on')
        r2 = Radiobutton(frame2,text = 'off', variable = self.v, value = 'off')
        r1.select()
        b3 = Button(frame2, text = 'send', command = self.bt4)
        b3.pack(sid = RIGHT, padx = 5)
        r1.pack(side = LEFT)
        r2.pack(side = LEFT)

        frame3 = Frame(self.win)
        frame3.pack()
        self.d = StringVar()
        self.d.set('default')
        label = Label(frame3, textvariable = self.d, relief = RAISED)
        label.pack(side = RIGHT)

    def update(self):
        data = self.ser.readline(self.ser.inWaiting())
        self.d.set(data)
        self.win.after(100,self.update)

    def run(self):
        self.makewindow()
        self.update()
        self.win.mainloop()

SerialViewer().run()
