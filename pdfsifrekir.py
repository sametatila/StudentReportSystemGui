import os,time,threading,pikepdf
from pikepdf import _cpphelpers
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
kapatbutton = PhotoImage(file="./data/pngs/kapatbutton.png")
def pdffoldercommand():
    global pdffolder
    pdffolder = filedialog.askdirectory()
    
def closebar():
    
    window1.destroy()

def startpsk():
    global window1
    window1 = Toplevel()
    window1.config(background="#545454")
    percent = StringVar()
    text1 = StringVar()
    inputfile2 = StringVar()
    p1 = ttk.Progressbar(window1, length=350, cursor='spider',mode="determinate",orient=HORIZONTAL)
    p1.pack(padx=20,pady=10)
    
    percentLabel = Label(window1, bg="#545454", fg="#ffffff", textvariable=percent).pack()
    taskLabel = Label(window1, bg="#545454", fg="#ffffff", textvariable=text1).pack()
    closebutton = Button(window1, image=kapatbutton, border=0, bg="#545454", command=closebar).pack(pady=10)
    itemnum = -1
    x = [os.path.join(r,file) for r,d,f in os.walk(pdffolder) for file in f if file.endswith(".PDF")]
    x1 = [os.path.join(file) for r,d,f in os.walk(pdffolder) for file in f if file.endswith(".PDF")]
    y = len(x)
    a = int(itemnum)
    z = int(y)-1
    speed = 1
    while a < z :
        itemnum = itemnum+1
        inputfile = x[itemnum]
        inputfile1 = x1[itemnum]
        time.sleep(0.0001)
        p1['value']+=(speed/(z+1))*100
        a+=speed
        inputfile2.set(str(inputfile1))
        percent.set(str(int(((a+1)/(z+1))*100))+"%")
        text1.set(str(a+1)+"/"+str(z+1)+" belge tamamlandı.")
        window1.update_idletasks()
        pdf = pikepdf.open(inputfile, allow_overwriting_input=True)
        pdf.save(inputfile)
        print(inputfile)
    
def pdfsifrekirmabutton():
    t1 = threading.Thread(target=startpsk)
    t1.start()
    
