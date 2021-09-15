from tkinter import *
from tkinter import ttk

import step1srvec as s1classic
import step2srvec as s2classic
import step1castlesrvec as s1castle
import pspksrvec as pspk

def step1menu():
    hideframes()
    primarys1.pack(fill="both",expand=1,padx=5,pady=5)
    
def step2menu():
    hideframes()
    primarys2.pack(fill="both",expand=1,padx=5,pady=5)
    
def pspkmenu():
    hideframes()
    primaryspk.pack(fill="both",expand=1,padx=5,pady=5)
    
def pskmenu():
    hideframes()
    pdfsifrekirmaframe.pack(fill="both",expand=1,padx=5,pady=5)
        
def tkbmenu():
    hideframes()
    turkcekarnebolmeframe.pack(fill="both",expand=1,padx=5,pady=5)

def hideframes():
    primarys1.pack_forget()
    primarys2.pack_forget()
    primaryspk.pack_forget()
    pdfsifrekirmaframe.pack_forget()
    turkcekarnebolmeframe.pack_forget()
    
def bos():
    pass



root = Tk()
root.title('Goaltesting Otomasyon')
root.minsize(800,400)
root.maxsize(800,400)

#Ana ekran karşılama
label01 = Label(root, text="Hoşgeldiniz", bg="#545454", fg="#ffffff",font=("Calibri", 25)).place(relx=0.4,rely=0.4)

#Menubar
menubar = Menu(root, bg="#545454")
root.config(menu=menubar)

osonucmenu = Menu(menubar,tearoff=0)
menubar.add_cascade(label="Online Sonuç", menu=osonucmenu)
osonucmenu.add_command(label="Primary Step 1", command=step1menu)
osonucmenu.add_command(label="Primary Step 2", command=step2menu)
osonucmenu.add_command(label="Primary Speaking", command=pspkmenu)
osonucmenu.add_command(label="Junior Standard", command=bos)
osonucmenu.add_command(label="Junior Speaking", command=bos)
osonucmenu.add_command(label="ITP", command=bos)

ysonucmenu = Menu(menubar,tearoff=0)
menubar.add_cascade(label="Yüzyüze Sonuç", menu=ysonucmenu)
ysonucmenu.add_command(label="Primary Step 1", command=bos)
ysonucmenu.add_command(label="Primary Step 2", command=bos)
ysonucmenu.add_command(label="Primary Speaking", command=bos)
ysonucmenu.add_command(label="Junior Standard", command=bos)
ysonucmenu.add_command(label="Junior Speaking", command=bos)
ysonucmenu.add_command(label="ITP", command=bos)

araclarmenu = Menu(menubar,tearoff=0)
menubar.add_cascade(label="Araçlar", menu=araclarmenu)
araclarmenu.add_command(label="PDF Şifre Kırma", command=pskmenu)
araclarmenu.add_command(label="Türkçe Karne Bölme", command=tkbmenu)


#Buton resimleri
excelbutton = PhotoImage(file="./data/pngs/excelbutton.png")
scorebutton = PhotoImage(file="./data/pngs/scorebutton.png")
ciktibutton = PhotoImage(file="./data/pngs/ciktibutton.png")
islembutton = PhotoImage(file="./data/pngs/islembutton.png")
pdfbutton = PhotoImage(file="./data/pngs/pdfbutton.png")

#Primary Step 1
primarys1 = ttk.Notebook(root)
frame11 = Frame(primarys1, bg="#545454")
frame12 = Frame(primarys1, bg="#545454")
frame11.pack(fill="both", expand=1)
frame12.pack(fill="both", expand=1)
primarys1.add(frame11, text="   Step 1 Classic   ")
primarys1.add(frame12, text="   Step 1 Castle   ")
#Step 1 Classic
s1label11 = Label(frame11, text="1.Sonuç Excel'i Seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.16)
s1label12 = Label(frame11, text="2.İndirilmiş olan Score Report'ların ana klasörüne seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.31)
s1label13 = Label(frame11, text="3.Dosyaları kaydetmek için klasör seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.46)
s1button11 = Button(frame11,image=excelbutton, borderwidth=0,bg="#545454", command=s1classic.excelfilecommand).place(relx=0.25,rely=0.15)
s1button12 = Button(frame11,image=scorebutton, borderwidth=0,bg="#545454", command=s1classic.scorefoldercammand).place(relx=0.25,rely=0.30)
s1button13 = Button(frame11,image=ciktibutton, borderwidth=0,bg="#545454", command=s1classic.outputfoldercommand).place(relx=0.25,rely=0.45)
s1button14 = Button(frame11,image=islembutton, borderwidth=0,bg="#545454", command=s1classic.step1_classicbutton).place(relx=0.42,rely=0.60)
#Step 1 Castle
s1label21 = Label(frame12, text="1.Sonuç Excel'i Seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.16)
s1label22 = Label(frame12, text="2.İndirilmiş olan Score Report'ların ana klasörüne seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.31)
s1label23 = Label(frame12, text="3.Dosyaları kaydetmek için klasör seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.46)
s1button21 = Button(frame12,image=excelbutton, borderwidth=0,bg="#545454", command=s1castle.excelfilecommand).place(relx=0.25,rely=0.15)
s1button22 = Button(frame12,image=scorebutton, borderwidth=0,bg="#545454", command=s1castle.scorefoldercammand).place(relx=0.25,rely=0.30)
s1button23 = Button(frame12,image=ciktibutton, borderwidth=0,bg="#545454", command=s1castle.outputfoldercommand).place(relx=0.25,rely=0.45)
s1button24 = Button(frame12,image=islembutton, borderwidth=0,bg="#545454", command=s1castle.step1_castlebutton).place(relx=0.42,rely=0.60)

#Primary Step 2
primarys2 = ttk.Notebook(root)
frame21 = Frame(primarys2, bg="#545454")
frame22 = Frame(primarys2, bg="#545454")
frame21.pack(fill="both", expand=1)
frame22.pack(fill="both", expand=1)
primarys2.add(frame21, text="   Step 2 Classic   ")
primarys2.add(frame22, text="   Step 2 Castle   ")
#Step 2 Classic
s2label11 = Label(frame21, text="1.Sonuç Excel'i Seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.16)
s2label12 = Label(frame21, text="2.İndirilmiş olan Score Report'ların ana klasörüne seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.31)
s2label13 = Label(frame21, text="3.Dosyaları kaydetmek için klasör seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.46)
s2button11 = Button(frame21,image=excelbutton, borderwidth=0,bg="#545454", command=s2classic.excelfilecommand).place(relx=0.25,rely=0.15)
s2button12 = Button(frame21,image=scorebutton, borderwidth=0,bg="#545454", command=s2classic.scorefoldercammand).place(relx=0.25,rely=0.30)
s2button13 = Button(frame21,image=ciktibutton, borderwidth=0,bg="#545454", command=s2classic.outputfoldercommand).place(relx=0.25,rely=0.45)
s2button14 = Button(frame21,image=islembutton, borderwidth=0,bg="#545454", command=s2classic.step2_classicbutton).place(relx=0.42,rely=0.60)
"""
#Step 2 Castle
s2label31 = Label(frame22, text="1.Sonuç Excel'i Seç!").pack(pady=5)
s2label32 = Label(frame22, text="2.İndirilmiş olan Score Report'ların ana klasörüne seç!").pack(pady=5)
s2label33 = Label(frame22, text="3.Dosyaları kaydetmek için klasör seç!").pack(pady=5)
s2button31 = Button(frame22, text="Excel Seç!", width=20, command=s2classic.excelfilecommand).pack(pady=5)
s2button32 = Button(frame22, text="Score Report Seç!", width=20, command=s2classic.scorefoldercammand).pack(pady=5)
s2button33 = Button(frame22, text="Kayıt Klasörü Seç!", width=20, command=s2classic.outputfoldercommand).pack(pady=5)
s2button34 = Button(frame22, text="İşlemi Başlat!", width=20, command=s2classic.step2_classicbutton).pack(pady=5)
"""

#Primary Speaking
primaryspk = ttk.Notebook(root)
frame31 = Frame(primaryspk, bg="#545454")
frame31.pack(fill="both", expand=1)
primaryspk.add(frame31, text="Primary Speaking")
#Speaking Score and Certificate
spklabel11 = Label(frame31, text="1.Sonuç Excel'i Seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.16)
spklabel12 = Label(frame31, text="2.İndirilmiş olan Score Report'ların ana klasörüne seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.31)
spklabel13 = Label(frame31, text="3.Dosyaları kaydetmek için klasör seç!",bg="#545454",fg="#ffffff").place(relx=0.5,rely=0.46)
spkbutton11 = Button(frame31,image=excelbutton, borderwidth=0,bg="#545454", command=pspk.excelfilecommand).place(relx=0.25,rely=0.15)
spkbutton12 = Button(frame31,image=scorebutton, borderwidth=0,bg="#545454", command=pspk.scorefoldercammand).place(relx=0.25,rely=0.30)
spkbutton13 = Button(frame31,image=ciktibutton, borderwidth=0,bg="#545454", command=pspk.outputfoldercommand).place(relx=0.25,rely=0.45)
spkbutton14 = Button(frame31,image=islembutton, borderwidth=0,bg="#545454", command=pspk.pspksrcbutton).place(relx=0.42,rely=0.60)


#Araçlar
#PDF Şifre Kırma
import pdfsifrekir as psk
pdfsifrekirmaframe = Frame(root, bg="#545454")
a1label01 = Label(pdfsifrekirmaframe, text="PDF Şifre Kırma", bg="#545454", fg="#ffffff",font=("Calibri", 16)).place(relx=0.42,rely=0.02)
a1label11 = Label(pdfsifrekirmaframe, text="1.PDF Klasörü Seç",bg="#545454",fg="#ffffff").place(relx=0.55,rely=0.16)
a1button11 = Button(pdfsifrekirmaframe,image=pdfbutton, borderwidth=0,bg="#545454", command=psk.pdffoldercommand).place(relx=0.30,rely=0.15)
a1button12 = Button(pdfsifrekirmaframe,image=islembutton, borderwidth=0,bg="#545454", command=psk.pdfsifrekirmabutton).place(relx=0.42,rely=0.30)

#Türkçe Karne Bölme
turkcekarnebolmeframe = Frame(root, bg="#545454")
a2label01 = Label(turkcekarnebolmeframe, text="Türkçe Karne Bölme", bg="#545454", fg="#ffffff",font=("Calibri", 16)).place(relx=0.4,rely=0.02)
a2label11 = Label(turkcekarnebolmeframe, text="1.Türkçe Karne PDF Seç",bg="#545454",fg="#ffffff").place(relx=0.55,rely=0.16)
a2label12 = Label(turkcekarnebolmeframe, text="1.Çıktı Klasörü Seç",bg="#545454",fg="#ffffff").place(relx=0.55,rely=0.16)
a2button11 = Button(turkcekarnebolmeframe,image=pdfbutton, borderwidth=0,bg="#545454", command=psk.pdffoldercommand).place(relx=0.30,rely=0.15)
a2button12 = Button(turkcekarnebolmeframe,image=ciktibutton, borderwidth=0,bg="#545454", command=psk.pdffoldercommand).place(relx=0.30,rely=0.15)
a2button13 = Button(turkcekarnebolmeframe,image=islembutton, borderwidth=0,bg="#545454", command=psk.pdfsifrekirmabutton).place(relx=0.42,rely=0.30)


#Styling
style = ttk.Style()
style.theme_create('Cloud', parent="classic", settings={
    ".": {
        "configure": {
            "background": '#aeb0ce', # All colors except for active tab-button
        }
    },

    "TNotebook": {
        "configure": {
            "background":"#545454", # color behind the notebook
            "tabmargins": [3, 3, 0, 0], # [left margin, upper margin, right margin, margin beetwen tab and frames]
        }
    },
    "TNotebook.Tab": {
        "configure": {
            "background": '#797979', # Color of non selected tab-button
            "padding": [10, 4], # [space beetwen text and horizontal tab-button border, space between text and vertical tab_button border]
        },
        "map": {
            "background": [("selected", '#545454')], # Color of active tab
            "expand": [("selected", [1, 1, 1, 0])], # [expanse of text]
            "foreground": [("selected", "#ffffff"),("!disabled", "#000000")] 
        }
    }
})
style.theme_use('Cloud')
#Styling End

root.config(background="#545454")
root.mainloop()
