from tkinter import *
import os, zipfile, shutil
from xml.dom import minidom
def POSFileCheck(strSourceFilePath,strArchiveFilePath,strResultFilePath,strFolderDate) :
    os.mkdir('.\\test')
    for filename in os.listdir("."):
        if filename.endswith(".zip"):
            print(filename)
            name = os.path.splitext(os.path.basename(filename))[0]
            if not os.path.isdir(name):
                zip = zipfile.ZipFile(filename)
                # os.mkdir(name)
                zip.extractall('.\\test')
    for xmlfiles in os.listdir('.\\test'):
        doc = minidom.parse(f'.\\test\\{xmlfiles}')
        # doc.getElementsByTagName returns NodeList
        name = doc.getElementsByTagName("title")[0]
        print(name.firstChild.data)
        if name.firstChild.data == 'nandha':
            shutil.copyfile(f'.\\test\\{xmlfiles}',f'.\\{xmlfiles}')
master = Tk()
master.title("POSLog Verification")
master.geometry('300x300')
Label(master, text="Source File Path").grid(row=0)
Label(master, text="Archive File Path").grid(row=2)
Label(master, text="Result File Path").grid(row=4)
Label(master, text="FolderDate(YY:MM:DD)").grid(row=6)
Label(master, text="Log Match Count").grid(row=8)
e1 = Entry(master)
e2 = Entry(master)
e3 = Entry(master)
e4 = Entry(master)
e5 = Entry(master)
e1.grid(row=0, column=2)
e2.grid(row=2, column=2)
e3.grid(row=4, column=2)
e4.grid(row=6, column=2)
e5.grid(row=8, column=2)
e6 = Button(master,text="Check it",)
e6.grid(row = 10,column=2)
strSourceFilePath = e1.get()
strArchiveFilePath = e2.get()
strResultFilePath = e3.get()
strFolderDate = e4.get()
mainloop( )