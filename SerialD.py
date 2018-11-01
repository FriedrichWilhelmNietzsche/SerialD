#! /usr/bin/env python  
# -*- coding: utf-8 -*-  
 
  
import threading
from tkinter import *
import tkinter as tk  
import serial.tools.list_ports  
from tkinter import ttk  
from tkinter import scrolledtext
import time
import os
import string  
import threading
import tkinter.filedialog
import matplotlib.pyplot as plt

from bs4 import  BeautifulSoup
import binascii
import re
import random
from datetime import datetime
from xlwt import *

 
SerialPort = serial.Serial()
GUI = tk.Tk()  # main frame
#GUI.iconbitmap(r'c:\\Users\Administrator\lf1128.ico')
GUI.title("Serial_D")  # frame title  
GUI.geometry("860x519")  # frame size



Receive = tk.LabelFrame( GUI,  padx = 0, pady = 0 ,relief = RAISED, bg = "white")     # log frame
Receive.place(x = 0, y = 0)
Receive.pack(side = LEFT, expand = YES, fill = BOTH)  



Receive_Window = scrolledtext.ScrolledText(Receive, bg = "white", width=88, height = 38, pady = 5,wrap = tk.WORD)
# Receive_Window.grid()
Receive_Window.pack( expand = YES, fill = BOTH)
#canvas = tkinter.Canvas(Receive_Window)
#canvas.configure(bg = "blue")



Receive_Window.port_list = list(serial.tools.list_ports_windows.comports()) # serial port information
serial_num=len(Receive_Window.port_list)  
serial_list=[] 
for i in range(0,serial_num):   
   serial_list.append(list(Receive_Window.port_list[i])[0])                 # serial port scan, print“COM n”  
   Receive_Window.insert("end", list(Receive_Window.port_list[i])[1]+"\n")  
   serial_list.sort()
   #Receive_Window.config(state=DISABLED)

  
option = tk.LabelFrame( GUI,  bg = "white",padx = 31,pady = 10)        #serial port choose frame
#option = Canvas(GUI,   bg = "white")
option.place(x = 640, y = 0, width = 222)
option.pack()

Send = tk.LabelFrame( GUI, text = "", padx = 0, pady = 0 ,bg = "white")  # send ascii frame
#Send = Canvas(GUI,   bg = "white" )
Send.place(x = 640, y = 190, width = "")
Send.pack()  
ButtonFrame1 = tk.LabelFrame( GUI, text = "", padx = 0, pady = 30, bg = "white" )
ButtonFrame1.place(x = 640, y = 222, width = 222)
ButtonFrame1.pack()  


# serial port 
lb1 = Label( option, text = "SerialPort   :"   )
lb1 .config(bg = "white" )
lb1.grid( column = 0, row = 0, sticky=W)                            
lb2 = Label( option, text = "BaudRate   :"    )
lb2.config(bg = "white" )
lb2.grid( column = 0, row = 1, sticky=W)  
lb3 = Label( option, text = "ByteSize     :"  )
lb3.config(bg = "white")
lb3.grid( column = 0, row = 2, sticky=W)
lb4 = Label( option, text = "StopBits     :"  )
lb4.config(bg ="white" )
lb4.grid( column = 0, row = 3, sticky=W)
lb5 = Label( option, text = "Parity         :")
lb5.config(bg = "white" )
lb5.grid( column = 0, row = 4, sticky=W)
lb6 = Label( option, text = "Flowcontrol:"    )
lb6.config(bg = "white" )
lb6.grid( column = 0, row = 5, sticky=W)
  
Port = tk.StringVar()                                
Port_list = ttk.Combobox( option, width = 8, textvariable = Port , state = 'readonly' )  
ListPorts = list(serial.tools.list_ports.comports())  
Port_list['values'] = [i[0] for i in ListPorts]  
Port_list.current(0)  
Port_list.grid(column=1, row=0)    
  
BaudRate = tk.StringVar()          
BaudRate_list = ttk.Combobox( option, width = 8, textvariable = BaudRate, state = 'readonly' )  
BaudRate_list['values'] = (1200, 2400, 4800, 9600, 14400, 19200, 38400, 43000, 57600, 76800, 115200)  
BaudRate_list.current(10)  
BaudRate_list.grid(column=1, row=1)   

Bytesize = tk.StringVar()         # undefinition
Bytesize_list = ttk.Combobox( option, width = 8, textvariable = Bytesize, state = 'readonly' )  
Bytesize_list['values'] = (5,6,7,8)  
Bytesize_list.current(3)  
Bytesize_list.grid(column=1, row=2) 

Stopbits = tk.StringVar()         # undefinition
Stopbits_list = ttk.Combobox( option, width = 8, textvariable = Stopbits, state = 'readonly' )  
Stopbits_list['values'] = (1, 1.5, 2)  
Stopbits_list.current(0)  
Stopbits_list.grid(column=1, row=3) 

Parity = tk.StringVar()           # undefinition
Parity_list = ttk.Combobox( option, width = 8, textvariable = Parity, state = 'readonly' )  
Parity_list['values'] = (None, "Odd", "Even", "Mark", "Space")  
Parity_list.current(0)  
Parity_list.grid(column=1, row=4) 

Flowcontrol = tk.StringVar()         # undefinition
Flowcontrol_list = ttk.Combobox( option, width = 8, textvariable = Flowcontrol, state = 'readonly' )
Flowcontrol_list['values'] = ("Hardware", "Software", None, "Custom")  
Flowcontrol_list.current(2)  
Flowcontrol_list.grid(column=1, row=5)



# add button frame definition
Open     = tk.StringVar(GUI,'Connect') 
isOpened = threading.Event()

 
EntrySend = tk.StringVar()                                                  
Send_Window = ttk.Entry(Send, textvariable = EntrySend, width= 19)  
Send_Window.grid()

      



 
def ReceiveData():                                             
    while SerialPort.isOpen():
        Receive_Window.config(state=NORMAL)
        Receive_Window.insert("end", str(SerialPort.readline())[2:-5] + '\n')
        Receive_Window.see("end")    
        Receive_Window.config(state=DISABLED)
  
def Close_Serial():  
    SerialPort.close()  
  
  
def SerialPortOpen():  
    if not isOpened.isSet():  
        try:
            SerialPort.port     = Port_list.get()  
            SerialPort.baudrate = BaudRate_list.get()
            SerialPort.timeout = None
        
            SerialPort.open()
            t = threading.Thread(target=ReceiveData)  
            t.setDaemon(True)  
            t.start()
            Receive_Window.insert("end","\nStart...\n")
        except Exception:
            Receive_Window.insert("end","\nSerial communications error!\n")
        else:
            isOpened.set()
            Open.set('Disconnect')                   
             
    else:  
        SerialPort.close()
        isOpened.clear()
        Open.set('Connect')
        Receive_Window.insert("end","\nStop...\n")

def Clear_Serial():
    Receive_Window.config(state=NORMAL)
    Receive_Window.delete(0.0,"end")

    

def Save_log():   
    Receive_Window.time_now=time.strftime("%Y%m%d_%H_%M_%S", time.localtime())   # log time  
    # print(Receive_Window.time_now)   
    with open (str(tkinter.filedialog.askdirectory())+"/log_"+ Receive_Window.time_now+".txt",'w+') as fb:   # os.getcwd()
        fb.write(Receive_Window.get(0.0,'end')) 
    Receive_Window.insert("end","\nLog has been saved!\n")


def String_Send():                                                    
    if  SerialPort.isOpen():  
        global DataSend  
        DataSend = EntrySend.get()  
        Receive_Window.insert("end", 'Sending the command ：' + str(DataSend) + '\n')  
        Receive_Window.see("end")  
        SerialPort.write(bytes(DataSend, encoding='utf8'))
    else:
        Receive_Window.insert("end","\nSerial communications error!\n")


def Read_file():
        filetypes = [  
                ("All Files", '*'),  
                ("Python Files", '*.py', 'TEXT'),  
                ("Text Files", '*.txt', 'TEXT'),  
                ("Config Files", '*.conf', 'TEXT')]  
        fobj = tkinter.filedialog.askopenfile( mode='r', filetypes=filetypes)
        if fobj:  
            Receive_Window.delete(0.0,"end") 
            Receive_Window.insert("end", "\n" + fobj.read()  + "\n")        
            Receive_Window.insert("end", "\n"+ fobj.name +"\n")  
       
        
                                 
def Demo_analysis():
    filetypes =  [("Text Files", '*.txt', 'TEXT')]
    Nu_Demo = tkinter.filedialog.askopenfile( mode='r', filetypes = filetypes )
    if Nu_Demo:          
       r = [ ] # original
       for line in Nu_Demo:
          line = line.split()
          if 'skt_im_op' in line:
             
              r.append( (float(str(line [-1]) [8:-1]) ) /1000 )

              
              """
              data = str( (float(str(line [-1]) [8:-1]) ) /1000 )
              for row, rowData in enumerate(data):
                 for col, value in enumerate(rowData):
                    ws.cell(row=row + 1, column=col + 1, value='%s' % value)
               
              ExcelWriter().write(writeIntoExcel, sheetsName="first", filename='123.xlsx')
              print('写入完成')

              """

       plt.figure('Figure 1')
       plt.title('Demo_Copy')
       plt.subplot(211)
       plt.plot(r, '.r',)
       plt.grid(True)
       #plt.subplot(212)
       plt.plot(r)
       #plt.grid(True)    
       plt.show()





def Dedi_analysis():
   
   filetypes =  [("Html Files", '*.html', 'HTML')]
   html_page = tkinter.filedialog.askopenfile( mode='r', filetypes = filetypes )
   soup = BeautifulSoup(html_page, "html.parser")
    
   for tag in soup.select('head'):
       tag.decompose()
   for tag in soup.select('span'):
       tag.decompose()
        #Receive_Window.insert("end", "\n" + soup.get_text()  + "\n")
       if soup:  
           Receive_Window.delete(0.0,"end") 
           Receive_Window.insert("end", "\n" + soup.get_text()  + "\n")        
           Receive_Window.insert("end", "\n"+ html_page.name +"\n") 










   
"""
def Demo_analysis2(file = 'new.xls', list = []):
    filetypes =  [("Text Files", '*.txt', 'TEXT')]
    Nu_Demo = tkinter.filedialog.askopenfile( mode='r', filetypes = filetypes )
	

    i = 0 #行序号

    if Nu_Demo:          
       for line in Nu_Demo:
           line = line.split()
           if 'skt_im_op' in line:
              x  = str( (float(str(line [-1]) [8:-1]) ) /1000 )
	#for app in list : #遍历list每一行
	#	j = 0 #列序号
	#	for x in app : #遍历该行中的每个内容（也就是每一列的）
	         sheet1.write(i, j, x)          
	      j = j+1 #列号递增
	      i = i+1 #行号递增
	      book.save(file) #创建保存文件


"""

"""  

# 多选择复选框
button=Checkbutton(tk,text="HEX",variable=var)
button.grid(row=2,columnspan=2,sticky=W)


def HEXDProc():
    if HexD.get():
        s = Receive_Window.get('1.0','end')
        s = ''.join('%02X' %i for i in [ord(c) for c in s])
        Receive_Window.delete('1.0','end')
        Receive_Window.insert("insert",s)
    else:
        s = Receive_Window.get('1.0','end')
        s = ''.join([chr(int(i,16)) for i in [s[i*2:i*2+2] for i in range(0,len(s)/2)]])
        Receive_Window.delete('1.0','end')
        Receive_Window.insert("insert",s)
        
    
    else:
  
      s = Receive_Window.get(0.0,'end')
      s = bytes.fromhex(hexStr)
      Receive_Window.delete(0.0,'end')
      Receive_Window.insert("end",s)
       
"""

  


   







# add button

tk.Button(Send     ,   text = "SendASC", bg = "#E0F5FA",width=10  , command = String_Send
          ).grid(row = 0, column = 1)
tk.Button(ButtonFrame1, textvariable=Open,bg = "#E0F5FA",width=9,command = lambda:SerialPortOpen() ).grid(row = 0, column = 0,sticky = tk.E)  
tk.Button(ButtonFrame1, text = "Save",  bg = "#E0F5FA", width=9, command = Save_log                ).grid(row = 0, column = 1)  
tk.Button(ButtonFrame1, text = "Clear", bg = "#E0F5FA", width=9, height = 1, command = Clear_Serial).grid(row = 0, column = 2)
ttk.Label(ButtonFrame1, text = ""                                                    ).grid(row = 1, column = 0, sticky=W )                    # blank row
tk.Button(ButtonFrame1, text =  "Read",  bg = "#E0F5FA",  width=9, command = Read_file              ).grid(row = 2, column = 0)
tk.Button(ButtonFrame1, text = "Nu_chart", bg = "#E0F5FA", width=9,command = Demo_analysis           ).grid(row = 2, column = 1)
ttk.Label(ButtonFrame1, text = ""                                                    ).grid(row = 3, column = 0, sticky=W )                    # blank row
tk.Button(ButtonFrame1, text = "Dedi_excel", bg = "#E0F5FA", width=9,command = Dedi_analysis          ).grid(row = 4, column = 0)
   
GUI.mainloop()






 
  
