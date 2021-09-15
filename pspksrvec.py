import os,time,threading
from tkinter import filedialog
from tkinter import *
from tkinter import ttk

def excelfilecommand():
    global bizimexcel
    bizimexcel = filedialog.askopenfilename()

def scorefoldercammand():
    global scorefolder
    scorefolder = filedialog.askdirectory()

def outputfoldercommand():
    global outputfolder
    outputfolder = filedialog.askdirectory()

def pdfmergersinif():
    global filename,f123
    from PyPDF2 import PdfFileMerger, PdfFileReader
    for root,dirs,files in os.walk(outputfolder):
        merger = PdfFileMerger()
        for filename in files:
            if filename.endswith("_ScoreReport.PDF"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_SR.pdf'))

    for root,dirs,files in os.walk(outputfolder):
        merger = PdfFileMerger()
        for filename in files:
            if filename.endswith("_Certificate.PDF"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_C.pdf'))

    for root,dirs,files in os.walk(outputfolder):
        merger = PdfFileMerger()
        for filename in files:
            if filename.endswith("_SR.pdf"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_SR_OKUL.PDF'))

    for root,dirs,files in os.walk(outputfolder):
        merger = PdfFileMerger()
        for filename in files:
            if filename.endswith("_C.pdf"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_C_OKUL.PDF'))
    for root, _, files in os.walk(outputfolder):
        for f in files:
            fullpath = os.path.join(root, f)
            try:
                if os.path.getsize(fullpath) == 306:   #set file size in kb
                    os.remove(fullpath)
            except WindowsError:
                print("Error" + fullpath)
    f123.set(str("Birleştime Tamamlandı!"))
    
def buttons():
    global birlestirbutton,kapatbutton
    birlestirbutton = PhotoImage(file="./data/pngs/birlestirbutton.png")
    kapatbutton = PhotoImage(file="./data/pngs/kapatbutton.png")
    
def pdfmergerbutton1():
    tr3 = threading.Thread(target=pdfmergersinif)
    tr3.start()


def pspksrc():
    global p1,window1,toplamsatir,satir,filename,f123
    buttons()
    import pdfplumber, re
    from reportlab.pdfgen import canvas 
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    import openpyxl
    #Excel bağlama
    book = openpyxl.load_workbook(bizimexcel)
    sheet = book['Sheet2']
    sheetlenght = len(sheet['A'])
    #Excel 2.satır dahil alacak
    satir = 1
    toplamsatir = int(sheetlenght)-1
    #Progressbar
    def closebar():
        window1.destroy()
    window1 = Toplevel()
    window1.config(bg="#545454")
    percent = StringVar()
    text1 = StringVar()
    f123 = StringVar()
    p1 = ttk.Progressbar(window1, length=350, cursor='spider',mode="determinate",orient=HORIZONTAL)
    p1.pack(padx=20,pady=10)
    percentLabel = Label(window1, textvariable=percent,bg="#545454",fg="#ffffff").pack()
    taskLabel = Label(window1, textvariable=text1,bg="#545454",fg="#ffffff").pack()
    closebutton = Button(window1,image=kapatbutton, borderwidth=0,bg="#545454", command=closebar).pack(pady=10)
    mergerclassbutton = Button(window1,image=birlestirbutton, borderwidth=0,bg="#545454", command=pdfmergerbutton1).pack(pady=10)
    f123Label = Label(window1, textvariable=f123,bg="#545454",fg="#ffffff").pack(pady=5)
    while satir<toplamsatir+1:
        satir += 1
        #Progressbar
        p1["value"] = satir+1
        p1["maximum"] = toplamsatir+1
        window1.update()
        time.sleep(0.00001)
        percent.set(str(int((((satir-2)*2)/((toplamsatir-1)*2))*100))+"%")
        text1.set(str((satir-2)*2)+"/"+str((toplamsatir-1)*2)+" belge tamamlandı.")
        
        td1 = str(sheet['H'+str(satir)].value)
        td2 = td1[5]+td1[6]
        td3 = str(sheet['G'+str(satir)].value)
        td4 = td3[5]+td3[6]
        if td2 > "00":
            if td2 == "01":
                td2 = "Jan"
            elif td2 == "02":
                td2 = "Feb"
            elif td2 == "03":
                td2 = "Mar"
            elif td2 == "04":
                td2 = "Apr"
            elif td2 == "05":
                td2 = "May" 
            elif td2 == "06":
                td2 = "Jun"
            elif td2 == "07":
                td2 = "Jul"
            elif td2 == "08":
                td2 = "Aug"
            elif td2 == "09":
                td2 = "Sep"
            elif td2 == "10":
                td2 = "Oct"
            elif td2 == "11":
                td2 = "Nov"
            elif td2 == "12":
                td2 = "Dec"
        if td4 > "00":
            if td4 == "01":
                td4 = "Jan"
            elif td4 == "02":
                td4 = "Feb"
            elif td4 == "03":
                td4 = "Mar"
            elif td4 == "04":
                td4 = "Apr"
            elif td4 == "05":
                td4 = "May" 
            elif td4 == "06":
                td4 = "Jun"
            elif td4 == "07":
                td4 = "Jul"
            elif td4 == "08":
                td4 = "Aug"
            elif td4 == "09":
                td4 = "Sep"
            elif td4 == "10":
                td4 = "Oct"
            elif td4 == "11":
                td4 = "Nov"
            elif td4 == "12":
                td4 = "Dec"

        #Tanımlar excelden
        school = sheet['A'+str(satir)].value
        studentname = str(sheet['D'+str(satir)].value)+str(", ")+str(sheet['E'+str(satir)].value)
        studentnamec = str(sheet['E'+str(satir)].value)+str(" ")+str(sheet['D'+str(satir)].value)
        studentnumber = sheet['F'+str(satir)].value
        studentclass = sheet['C'+str(satir)].value
        testdate = td1[8]+td1[9]+" "+td2+" "+td1[0]+td1[1]+td1[2]+td1[3]
        dateofbirth = td3[8]+td3[9]+" "+td4+" "+td3[0]+td3[1]+td3[2]+td3[3]
        totalscore = str(sheet['I'+str(satir)].value)
        cefr = str(sheet['J'+str(satir)].value)
        level = str(sheet['K'+str(satir)].value)
        
        #ETS Sonuç Belge Numarası ve Cinsiyet Sorgulama
        x = [os.path.join(r,file) for r,d,f in os.walk(scorefolder) for file in f if file.endswith(str(studentnumber)+".PDF")]

        with pdfplumber.open(x[0]) as pdf1:
            page = pdf1.pages[0]
            text = page.extract_text()
            #belge no buradan gelir
            name = re.compile(r'[0-9]{8}$')
            #sadece step 1 classic score report tan alınabilir. cinsiyet buradan gelir
            gender = re.compile(r'Student Number:.*')
        for line in text.split('\n'):
            if name.match(line):
                lname = line.split()
                print(lname)
        for line in text.split('\n'):
            if gender.match(line):
                lgender = line.split()

        #Yızdız resimleri score report
        pspk_nsimage = './data/pspk_ns.jpg'
        pspk_1ribbonsimage = './data/pspk_1ribbons.jpg'
        pspk_2ribbonsimage = './data/pspk_2ribbons.jpg'
        pspk_3ribbonsimage = './data/pspk_3ribbons.jpg'
        pspk_4ribbonsimage = './data/pspk_4ribbons.jpg'
        pspk_5ribbonsimage = './data/pspk_5ribbons.jpg'
        
        #Klasör oluşturma
        def createFolder(directory):
            try:
                if not os.path.exists(directory):
                    os.makedirs(directory)
            except OSError:
                print ('Error: Creating directory. ' +  directory)
        asd = "/"+str(school)+"/"+str(studentclass)+"/"
        createFolder(outputfolder+asd)

        #Liste işareti
        nokta = u'\u2022'

        def pspk_scorereport():
            #Speaking yıldıza göre bilgiler

            def pspk_ns(pdf):
                #NS Image
                pdf.drawImage(pspk_nsimage, 281.9,522.65, width=61,height=82,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 490.2, 444.75, 490.2)
                pdf.line(167.75, 464.7, 444.75, 464.7)
                pdf.line(167.75, 411.6, 444.75, 411.6)
                pdf.line(167.75, 386.11, 444.75, 386.11)
                pdf.line(167.75, 345.66, 444.75, 345.66)
                #Noktalar Yok
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(258.41,508.29,str(level)+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,440.98,"Your responses for this test were not properly recorded")
                pdf.drawString(169.2,428.33,"Therefore, a score cannot be provided.")
                #Part2 Boş
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,328.84,"CEFR")
                pdf.drawString(200.8,328.84,"Level:")
                pdf.drawString(231.5,328.84,str(cefr))
                pdf.drawString(169.2,316.19,"The student received  "+str(totalscore)+"  out of  27  points.")
                
            def pspk_0(pdf):
                #NS Image
                pdf.drawImage(pspk_nsimage, 281.9,522.65, width=61,height=82,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 490.2, 444.75, 490.2)
                pdf.line(167.75, 464.7, 444.75, 464.7)
                pdf.line(167.75, 411.6, 444.75, 411.6)
                pdf.line(167.75, 386.11, 444.75, 386.11)
                pdf.line(167.75, 345.66, 444.75, 345.66)
                #Noktalar Yok
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,508.29,str(totalscore)+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,440.98,"The student did not respond to the test tasks or did not")
                pdf.drawString(169.2,428.33,"respond in English.")
                #Part2 Boş
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,328.84,"CEFR")
                pdf.drawString(200.8,328.84,"Level:")
                pdf.drawString(231.5,328.84,str(cefr))
                pdf.drawString(169.2,316.19,"The student received  "+str(totalscore)+"  out of  27  points.")

            def pspk_1ribbons(pdf):
                #1ribbons Image
                pdf.drawImage(pspk_1ribbonsimage, 301.42,575.13, width=21.96,height=29.52,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 542.7, 444.75, 542.7)
                pdf.line(167.75, 505.7, 444.75, 505.7)
                pdf.line(167.75, 413.18, 444.75, 413.18)
                pdf.line(167.75, 387.68, 444.75, 387.68)
                pdf.line(167.75, 279.26, 444.75, 279.26)
                pdf.line(185.75, 307, 444.75, 307)
                #Noktalar
                pdf.setFont('dot', 21)
                pdf.drawString(185.9,477,nokta)
                pdf.drawString(185.9,438.4,nokta)
                pdf.drawString(185.9,359,nokta)
                pdf.drawString(185.9,345.4,nokta)
                pdf.drawString(185.9,319,nokta)
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,560.79,str(level[0])+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,526.82,"Students attempt to speak in English using words and")
                pdf.drawString(169.2,515.32,"simple phrases. They may be able to:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,481.24,"Say some common words in familiar categories")
                pdf.drawString(205.2,468.59,"such as home, school, family, colors, animals,")
                pdf.drawString(205.2,455.95,"and actions")
                pdf.drawString(205.2,442.55,"Say simple phrases")
                #Part2
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,397.3,"To improve their speaking ability, students should:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,363.22,"Learn and practice saying common words")
                pdf.drawString(205.2,349.83,"Name what they see in pictures (example:")
                pdf.setFont('abcit', 11)
                pdf.drawString(413.07,349.83,"I see")
                pdf.drawString(205.2,337.18,"a house.")
                pdf.setFont('abc', 11)
                pdf.drawString(247.4,337.18,")")
                pdf.drawString(205.2,323.79,"Practice speaking in sentences about objects")
                pdf.drawString(205.2,311.14,"and activities they like")
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,262.44,"CEFR")
                pdf.drawString(200.8,262.44,"Level:")
                pdf.drawString(231.5,262.44,str(cefr))
                pdf.drawString(169.2,249.8,"The student received  "+str(totalscore)+"  out of  27  points.")

            def pspk_2ribbons(pdf):
                #2ribbons Image
                pdf.drawImage(pspk_2ribbonsimage, 286.13,575.13, width=52.56,height=29.52,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 542.7, 444.75, 542.7)
                pdf.line(167.75, 494.2, 444.75, 494.2)
                pdf.line(167.75, 350.34, 444.75, 350.34)
                pdf.line(167.75, 324.84, 444.75, 324.84)
                pdf.line(167.75, 203.78, 444.75, 203.78)
                pdf.line(185.75, 231.58, 444.75, 231.58)
                #Noktalar
                pdf.setFont('dot', 21)
                pdf.drawString(185.9,465.4,nokta)
                pdf.drawString(185.9,426.8,nokta)
                pdf.drawString(185.9,388,nokta)
                pdf.drawString(185.9,295.8,nokta)
                pdf.drawString(185.9,270,nokta)
                pdf.drawString(185.9,244,nokta)
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,560.79,str(level[0])+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,526.82,"Students begin to speak in English by using words and")
                pdf.drawString(169.2,515.32,"simple statements. They begin to say what they like and")
                pdf.drawString(169.2,503.82,"give some descriptions. They can:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,469.74,"Say some common words in familiar categories")
                pdf.drawString(205.2,457.09,"such as home, school, family, colors, animals,")
                pdf.drawString(205.2,444.45,"and actions")
                pdf.drawString(205.2,431.06,"Communicate meaning in short, simple")
                pdf.drawString(205.2,418.41,"statements (")
                pdf.setFont('abcit', 11)
                pdf.drawString(265.72,418.41,"examples: The tiger is big. The zoo")
                pdf.drawString(205.2,405.76,"has two birds.")
                pdf.setFont('abc', 11)
                pdf.drawString(273.06,405.76,")")
                pdf.drawString(205.2,392.37,"Pronounce words and phrases clearly but slowly")
                pdf.drawString(205.2,379.72,"some of the time")     
                #Part2
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,334.46,"To improve their speaking ability, students should:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,300.38,"Learn more words that describe familiar places,")
                pdf.drawString(205.2,287.73,"objects, and people")
                pdf.drawString(205.2,274.34,"Practice asking and answering questions about")
                pdf.drawString(205.2,261.7,"everyday topics")
                pdf.drawString(205.2,248.3,"Practice describing what happens in stories they")
                pdf.drawString(205.2,235.66,"read and programs they watch")
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,186.96,"CEFR")
                pdf.drawString(200.8,186.96,"Level:")
                pdf.drawString(231.5,186.96,str(cefr))
                pdf.drawString(169.2,174.31,"The student received  "+str(totalscore)+"  out of  27  points.")

            def pspk_3ribbons(pdf):
                #3ribbons Image
                pdf.drawImage(pspk_3ribbonsimage, 270.3,575.13, width=84.24,height=29.52,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 542.7, 444.75, 542.7)
                pdf.line(167.75, 494.2, 444.75, 494.2)
                pdf.line(167.75, 348.86, 444.75, 348.86)
                pdf.line(167.75, 323.36, 444.75, 323.36)
                pdf.line(167.75, 202.3, 444.75, 202.3)
                pdf.line(185.75, 230.09, 444.75, 230.09)
                #Noktalar
                pdf.setFont('dot', 21)
                pdf.drawString(185.9,465.4,nokta)
                pdf.drawString(185.9,439,nokta)
                pdf.drawString(185.9,413.4,nokta)
                pdf.drawString(185.9,400,nokta)
                pdf.drawString(185.9,386.5,nokta)
                pdf.drawString(185.9,295,nokta)
                pdf.drawString(185.9,268.4,nokta)
                pdf.drawString(185.9,242.6,nokta)
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,560.79,str(level[0])+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,526.82,"Students speak in English to say what they like and give")
                pdf.drawString(169.2,515.32,"some descriptions. They begin to ask questions and tell")
                pdf.drawString(169.2,503.82,"stories. They can:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,469.74,"Use words and phrases to communicate")
                pdf.drawString(205.2,457.09,"meaning")
                pdf.drawString(205.2,443.71,"Use a limited number of grammatical structures")
                pdf.drawString(205.2,431.06,"to describe objects and actions")
                pdf.drawString(205.2,417.67,"Begin to form questions and requests")
                pdf.drawString(205.2,404.28,"Begin to communicate a sequence of events")
                pdf.drawString(205.2,390.89,"Pronounce words and statements clearly most of")
                pdf.drawString(205.2,378.24,"the time")        
                #Part2
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,332.98,"To improve their speaking ability, students should:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,298.9,"Learn more words that describe familiar places,")
                pdf.drawString(205.2,286.25,"objects, and people")
                pdf.drawString(205.2,272.86,"Practice asking and answering questions about")
                pdf.drawString(205.2,260.21,"everyday topics")
                pdf.drawString(205.2,246.82,"Practice describing in sentences what happens")
                pdf.drawString(205.2,234.17,"in stories they read and programs they watch")
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,185.48,"CEFR")
                pdf.drawString(200.8,185.48,"Level:")
                pdf.drawString(231.5,185.48,str(cefr))
                pdf.drawString(169.2,172.83,"The student received  "+str(totalscore)+"  out of  27  points.")

            def pspk_4ribbons(pdf):
                #4ribbons Image
                pdf.drawImage(pspk_4ribbonsimage, 254.27,575.13, width=116.28,height=29.52,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 542.7, 444.75, 542.7)
                pdf.line(167.75, 494.2, 444.75, 494.2)
                pdf.line(167.75, 374.16, 444.75, 374.16)
                pdf.line(167.75, 348.66, 444.75, 348.66)
                pdf.line(167.75, 214.94, 444.75, 214.94)
                pdf.line(185.75, 242.74, 444.75, 242.74)
                #Noktalar
                pdf.setFont('dot', 21)
                pdf.drawString(185.9,465,nokta)
                pdf.drawString(185.9,452,nokta)
                pdf.drawString(185.9,438,nokta)
                pdf.drawString(185.9,425,nokta)
                pdf.drawString(185.9,412,nokta)
                pdf.drawString(185.9,320,nokta)
                pdf.drawString(185.9,294,nokta)
                pdf.drawString(185.9,268,nokta)
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,560.79,str(level[0])+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,526.82,"Students speak in English to express and explain what")
                pdf.drawString(169.2,515.32,"they like and give directions. They begin to expand their")
                pdf.drawString(169.2,503.82,"descriptions of things and events. They can:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,469.74,"Use appropriate word choices")
                pdf.drawString(205.2,456.35,"Use complete statements to communicate ideas")
                pdf.drawString(205.2,442.96,"Use appropriate grammatical structures")
                pdf.drawString(205.2,429.57,"Begin to form questions and requests")
                pdf.drawString(205.2,416.18,"Speak clearly with few errors in pronunciation or")
                pdf.drawString(205.2,403.54,"intonation")
                #Part2
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,358.28,"To improve their speaking ability, students should:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,324.2,"Learn less common words that describe familiar")
                pdf.drawString(205.2,311.55,"places, objects, and people")
                pdf.drawString(205.2,298.16,"Practice asking and answering questions about")
                pdf.drawString(205.2,285.51,"everyday topics")
                pdf.drawString(205.2,272.12,"Practice giving details about places, people, and")
                pdf.drawString(205.2,259.47,"events in the stories they read and programs")
                pdf.drawString(205.2,246.82,"they watch")
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,198.13,"CEFR")
                pdf.drawString(200.8,198.13,"Level:")
                pdf.drawString(231.5,198.13,str(cefr))
                pdf.drawString(169.2,185.48,"The student received  "+str(totalscore)+"  out of  27  points.")
                
            def pspk_5ribbons(pdf):
                #4ribbons Image
                pdf.drawImage(pspk_5ribbonsimage, 238.98,575.13, width=146.88,height=29.52,mask=None)
                #Çizgiler
                pdf.setLineWidth(1.5)
                pdf.line(167.75, 542.7, 444.75, 542.7)
                pdf.line(167.75, 482.7, 444.75, 482.7)
                pdf.line(167.75, 299.41, 444.75, 299.41)
                pdf.line(167.75, 273.91, 444.75, 273.91)
                pdf.line(167.75, 140.2, 444.75, 140.2)
                pdf.line(185.75, 168, 444.75, 168)
                #Noktalar
                pdf.setFont('dot', 21)
                pdf.drawString(185.9,454.4,nokta)
                pdf.drawString(185.9,428,nokta)
                pdf.drawString(185.9,402,nokta)
                pdf.drawString(185.9,363,nokta)
                pdf.drawString(185.9,337,nokta)
                pdf.drawString(185.9,245,nokta)
                pdf.drawString(185.9,219,nokta)
                pdf.drawString(185.9,193,nokta)
                #Part0
                pdf.setFont('abcbold', 11)
                pdf.drawString(250.44,611.58,"The")
                pdf.drawString(274.4,611.58,"Student's")
                pdf.drawString(328.51,611.58,"Level")
                pdf.drawString(361.26,611.58,"is:")
                pdf.setFont('abcbold', 12)
                pdf.drawString(259.07,560.79,str(level[0])+" Out of 5 Ribbons")
                #Part1
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,526.82,"Students speak in English to expand descriptions,")
                pdf.drawString(169.2,515.32,"communicate multistep directions, and tell stories")
                pdf.drawString(169.2,503.82,"effectively. They successfully ask questions and make")
                pdf.drawString(169.2,492.32,"simple requests. They can:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,458.25,"Use a wide range of vocabulary and")
                pdf.drawString(205.2,445.6,"grammatical structures effectively")
                pdf.drawString(205.2,432.21,"Include relevant details to expand descriptions,")
                pdf.drawString(205.2,419.56,"give directions, and tell stories")
                pdf.drawString(205.2,406.17,"Include structures such as connecting words and")
                pdf.drawString(205.2,393.52,"phrases that make directions and stories easy to")
                pdf.drawString(205.2,380.87,"follow")
                pdf.drawString(205.2,367.48,"Form questions and requests appropriately and")
                pdf.drawString(205.2,354.83,"use intonation to communicate meaning")
                pdf.drawString(205.2,341.44,"Speak fluidly with few errors in pronunciation or")
                pdf.drawString(205.2,328.79,"intonation")
                #Part2
                pdf.setFont('abcbold', 10)
                pdf.drawString(169.2,283.53,"To improve their speaking ability, students should:")
                pdf.setFont('abc', 11)
                pdf.drawString(205.2,249.46,"Read and listen to age-appropriate academic")
                pdf.drawString(205.2,236.81,"content")
                pdf.drawString(205.2,223.42,"Speak and write about age-appropriate")
                pdf.drawString(205.2,210.77,"academic content")
                pdf.drawString(205.2,197.38,"Consider taking the")
                pdf.setFont('abcit', 11)
                pdf.drawString(303.3,197.38,"TOEFL Junior")
                pdf.setFont('abc', 11)
                pdf.drawString(371.94,197.38,"® Speaking")
                pdf.drawString(205.2,184.73,"test for more information about their speaking")
                pdf.drawString(205.2,172.08,"ability")
                #Part3
                pdf.setFont('abc', 11)
                pdf.drawString(169.2,123.38,"CEFR")
                pdf.drawString(200.8,123.38,"Level:")
                pdf.drawString(231.5,123.38,str(cefr))
                pdf.drawString(169.2,110.73,"The student received  "+str(totalscore)+"  out of  27  points.")

            #Belge özellikleri
            fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_ScoreReport.PDF")
            documentTitle = 'Document title!'
            pdf = canvas.Canvas(fileName,pagesize=(595,842))
            pdf.setTitle(documentTitle)
            #Font
            pdfmetrics.registerFont(TTFont('abc', './data/arial.ttf'))
            pdfmetrics.registerFont(TTFont('abcbold', './data/arialbold.ttf'))
            pdfmetrics.registerFont(TTFont('abcit', './data/arialitalic.ttf'))
            pdfmetrics.registerFont(TTFont('dot', './data/comic.ttf'))

            #Ana baskı alanı
            def baskialani(pdf):
                #Öğrenci Bilgisi
                pdf.setFont('abc', 12)
                pdf.drawString(56.05,723.93,"Student Name:")
                pdf.drawString(56.05,703.24,"Student Number:")
                pdf.drawString(56.05,682.54,"Test Date:")
                pdf.drawString(382.56,723.73,"Date of Birth:")
                pdf.drawString(382.56,703.03,"Gender:")
                pdf.setFont('abc', 10)
                pdf.drawString(137.83,723.93,str(studentname))
                pdf.drawString(149.3,703.24,str(studentnumber))
                pdf.drawString(113.94,682.54,str(testdate))
                pdf.drawString(456.21,723.73,str(dateofbirth))
                pdf.drawString(429.53,703.03,str(lgender[4]))
                #Alt bilgi Okul vs.
                pdf.setFont('abc', 8)
                pdf.drawCentredString(298.97,51.1,str(school)+", Turkey")
                pdf.drawCentredString(306.1,40.52,"OYD - Okul Yayin, Turkey")
                pdf.drawCentredString(297.14,29.94,str(lname[0]))

            #Anabaskı kodu
            baskialani(pdf)

            #Skora göre sonuç seçimi
            if totalscore == "NS":
                pspk_ns(pdf)
            elif totalscore == str("0"):
                pspk_0(pdf)
            elif totalscore == str("1"):
                pspk_1ribbons(pdf)
            elif totalscore == str("2"):
                pspk_1ribbons(pdf)
            elif totalscore == str("3"):
                pspk_1ribbons(pdf)
            elif totalscore == str("4"):
                pspk_1ribbons(pdf)
            elif totalscore == str("5"):
                pspk_1ribbons(pdf)
            elif totalscore == str("6"):
                pspk_1ribbons(pdf)
            elif totalscore == str("7"):
                pspk_2ribbons(pdf)
            elif totalscore == str("8"):
                pspk_2ribbons(pdf)
            elif totalscore == str("9"):
                pspk_2ribbons(pdf)
            elif totalscore == str("10"):
                pspk_2ribbons(pdf)
            elif totalscore == str("11"):
                pspk_2ribbons(pdf)
            elif totalscore == str("12"):
                pspk_2ribbons(pdf)
            elif totalscore == str("13"):
                pspk_3ribbons(pdf)
            elif totalscore == str("14"):
                pspk_3ribbons(pdf)
            elif totalscore == str("15"):
                pspk_3ribbons(pdf)
            elif totalscore == str("16"):
                pspk_3ribbons(pdf)
            elif totalscore == str("17"):
                pspk_3ribbons(pdf)
            elif totalscore == str("18"):
                pspk_4ribbons(pdf)
            elif totalscore == str("19"):
                pspk_4ribbons(pdf)
            elif totalscore == str("20"):
                pspk_4ribbons(pdf)
            elif totalscore == str("21"):
                pspk_4ribbons(pdf)
            elif totalscore == str("22"):
                pspk_4ribbons(pdf)
            elif totalscore == str("23"):
                pspk_5ribbons(pdf)
            elif totalscore == str("24"):
                pspk_5ribbons(pdf)
            elif totalscore == str("25"):
                pspk_5ribbons(pdf)
            elif totalscore == str("26"):
                pspk_5ribbons(pdf)
            elif totalscore == str("27"):
                pspk_5ribbons(pdf)   
                
            pdf.save()

        def pspk_certificate():
            if str(totalscore)>str("0"):
                if str(totalscore)!=str("NS"):
                    #Speaking sonuçları
                    def pspkc_1ribbons(pdf):
                        pdf.drawImage(pspk_1ribbonsimage, 393.9,191.27, width=21.96,height=29.52,mask=None)

                    def pspkc_2ribbons(pdf):
                        pdf.drawImage(pspk_2ribbonsimage, 393.9,191.27, width=52.56,height=29.52,mask=None)

                    def pspkc_3ribbons(pdf):
                        pdf.drawImage(pspk_3ribbonsimage, 393.9,191.27, width=84.24,height=29.52,mask=None)

                    def pspkc_4ribbons(pdf):
                        pdf.drawImage(pspk_4ribbonsimage, 393.9,191.27, width=116.28,height=29.52,mask=None)
                        
                    def pspkc_5ribbons(pdf):
                        pdf.drawImage(pspk_5ribbonsimage, 393.9,191.27, width=146.88,height=29.52,mask=None)

                    #Belge özellikleri
                    fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_Certificate.PDF")
                    documentTitle = 'Document title!'
                    pdf = canvas.Canvas(fileName,pagesize=(842,595.5))
                    pdf.setTitle(documentTitle)
                    #Font
                    pdfmetrics.registerFont(TTFont('abc', './data/arial.ttf'))
                    pdfmetrics.registerFont(TTFont('abcbold', './data/arialbold.ttf'))
                    pdfmetrics.registerFont(TTFont('abcit', './data/arialitalic.ttf'))
                    pdfmetrics.registerFont(TTFont('dot', './data/comic.ttf'))

                    #Ana baskı alanı
                    def baskialani(pdf):
                        #Sertifika içeriği
                        pdf.setFont('abc', 32)
                        pdf.drawString(233.42,378.98,"Certificate of Achievement")
                        pdf.setFont('abc', 18)
                        pdf.drawString(338.92,332.57,"This is to certify that")
                        pdf.setFont('abcbold', 20)
                        pdf.drawCentredString(420.38,288.41,str(studentnamec))
                        pdf.setFont('abc', 16)
                        pdf.drawString(201.65,244.39,"has earned the following level on the TOEFL")
                        pdf.drawString(529.6,244.39,"Primary™ Test")
                        pdf.setFont('abc', 10.50)
                        pdf.drawString(517.39,249.21,"®")
                        pdf.setFont('abc', 16)
                        pdf.drawString(310.29,203.19,"Speaking:")
                        #Alt bilgi Okul vs.
                        pdf.setFont('abc', 7.5)
                        pdf.drawString(88.1,104.12,"Test Date: "+str(testdate))
                        pdf.drawString(88.1,95.49,str(school)+", Turkey")
                        pdf.drawString(88.1,86.87,"OYD - Okul Yayin, Turkey")
                        pdf.drawString(88.1,78.24,str(lname[0]))
                        
                    #Anabaskı kodu
                    baskialani(pdf)

                    #Skora göre sonuç seçimi
                    if totalscore == str("NS"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("0"):
                        pspkc_1ribbons(pdf)
                    if totalscore == str("1"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("2"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("3"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("4"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("5"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("6"):
                        pspkc_1ribbons(pdf)
                    elif totalscore == str("7"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("8"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("9"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("10"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("11"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("12"):
                        pspkc_2ribbons(pdf)
                    elif totalscore == str("13"):
                        pspkc_3ribbons(pdf)
                    elif totalscore == str("14"):
                        pspkc_3ribbons(pdf)
                    elif totalscore == str("15"):
                        pspkc_3ribbons(pdf)
                    elif totalscore == str("16"):
                        pspkc_3ribbons(pdf)
                    elif totalscore == str("17"):
                        pspkc_3ribbons(pdf)
                    elif totalscore == str("18"):
                        pspkc_4ribbons(pdf)
                    elif totalscore == str("19"):
                        pspkc_4ribbons(pdf)
                    elif totalscore == str("20"):
                        pspkc_4ribbons(pdf)
                    elif totalscore == str("21"):
                        pspkc_4ribbons(pdf)
                    elif totalscore == str("22"):
                        pspkc_4ribbons(pdf)
                    elif totalscore == str("23"):
                        pspkc_5ribbons(pdf)
                    elif totalscore == str("24"):
                        pspkc_5ribbons(pdf)
                    elif totalscore == str("25"):
                        pspkc_5ribbons(pdf)
                    elif totalscore == str("26"):
                        pspkc_5ribbons(pdf)
                    elif totalscore == str("27"):
                        pspkc_5ribbons(pdf)   

                    pdf.save()

        pspk_scorereport()
        pspk_certificate()

def pspksrcbutton():
    buttons()
    tr1 = threading.Thread(target=pspksrc)
    tr1.start()