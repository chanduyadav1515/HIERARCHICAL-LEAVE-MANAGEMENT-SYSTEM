from tkinter import * 
from tkinter.ttk import *
  
# importing askopenfile function 
# from class filedialog 
from tkinter.filedialog import askopenfile 
import time
import requests
  
# This function will be used to open 
# file in read mode and only Python files 
# will be opened 
def open_file():
        file=open("2.docx",'rb')
	#file = askopenfile(mode ='rb', filetypes =[('All', '*.*')])
        server_url = 'http://tejusolutions.com/rohit/index.php/'
        files = {'photo':file}
        data = {'submit':True}
        response = requests.post(server_url, files=files, data=data)
        #stat["text"] = 'status : '+str(response.status_code)
        print("done")

open_file() 
