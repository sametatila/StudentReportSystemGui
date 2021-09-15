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
            if filename.endswith("_TurkceKarne.PDF"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_TK.pdf'))

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

    for root,dirs,files in os.walk(outputfolder):
        merger = PdfFileMerger()
        for filename in files:
            if filename.endswith("_TK.pdf"):
                filepath = os.path.join(root, filename)
                merger.append(PdfFileReader(open(filepath, 'rb')))
                f123.set(str(filename))
        merger.write(os.path.join(outputfolder,os.path.normpath(root)+'_TK_OKUL.PDF'))
        
    for root, _, files in os.walk(outputfolder):
        for f in files:
            fullpath = os.path.join(root, f)
            try:
                if os.path.getsize(fullpath) < 300 * 350:   #set file size in kb
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

def step1_classic():
    global p1,window1,toplamsatir,satir,filename,f123
    buttons()
    import pdfplumber, re,math
    from reportlab.pdfgen import canvas 
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    import openpyxl
    #Excel bağlama
    book = openpyxl.load_workbook(bizimexcel)
    sheet = book['Sheet2']
    sheetlenght = len(sheet['A'])
    #Excel 3.satır dahil alacak
    satir = 2
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
    while satir<toplamsatir:
        satir += 1
        #Progressbar
        p1["value"] = satir+2
        p1["maximum"] = toplamsatir
        window1.update()
        time.sleep(0.00001)
        percent.set(str(int((((satir-2)*3)/((toplamsatir-2)*3))*100))+"%")
        text1.set(str((satir-2)*3)+"/"+str((toplamsatir-2)*3)+" belge tamamlandı.")

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
        sinavturu = 'Step 1'
        school = sheet['A'+str(satir)].value
        studentname = str(sheet['E'+str(satir)].value)+str(" ")+str(sheet['D'+str(satir)].value)
        studentnumber = sheet['F'+str(satir)].value
        studentclass = sheet['C'+str(satir)].value
        testdate = td1[8]+td1[9]+" "+td2+" "+td1[0]+td1[1]+td1[2]+td1[3]
        dateofbirth = td3[8]+td3[9]+" "+td4+" "+td3[0]+td3[1]+td3[2]+td3[3]
        rcefr = str(sheet['J'+str(satir)].value)
        rscore = str(sheet['I'+str(satir)].value)
        lcefr = str(sheet['N'+str(satir)].value)
        lscore = str(sheet['M'+str(satir)].value)
        lexile = str(sheet['K'+str(satir)].value)
        
        rstar = str(sheet['L'+str(satir)].value)
        lstar = str(sheet['O'+str(satir)].value)
        
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
        step1_nsimage = './data/step1_ns.jpg'
        step1_1starimage = './data/step1_1star.jpg'
        step1_2starimage = './data/step1_2star.jpg'
        step1_3starimage = './data/step1_3star.jpg'
        step1_4starimage = './data/step1_4star.jpg'

        #Yızdız resimleri seritifika
        step1c_nsimage = './data/step1c_ns.jpg'
        step1c_1starimage = './data/step1c_1star.jpg'
        step1c_2starimage = './data/step1c_2star.jpg'
        step1c_3starimage = './data/step1c_3star.jpg'
        step1c_4starimage = './data/step1c_4star.jpg'
        primarytk = './data/primarytk.jpg'

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

        def step1_scorereport():
            #Reading yıldıza göre bilgiler

            def step1_nsreading(pdf):
                #NS Image
                pdf.drawImage(step1_nsimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.29, 567.17, 279.36, 567.17)
                pdf.line(61.29, 497.02, 279.36, 497.02)
                pdf.line(61.29, 446.43, 279.36, 446.43)
                #Noktalar Yok
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,528.71,"The test taker did not respond to any questions in this")
                pdf.drawString(60.88,518.36,"section. Therefore, the scores for this section cannot")
                pdf.drawString(60.88,508.01,"be provided.")
                #Part2 Boş
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,427.52,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,417.18,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,406.82,"The student received "+str(rscore)+" on a scale of 100 to 109")

            def step1_1starreading(pdf):
                #1Star Image
                pdf.drawImage(step1_1starimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.29, 567.17, 279.36, 567.17)
                pdf.line(61.29, 496.42, 279.36, 496.42)
                pdf.line(61.29, 372.17, 279.36, 372.17)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(77.8,514,nokta)
                pdf.drawString(77.8,452.45,nokta)
                pdf.drawString(77.8,410.44,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,549.42,"Students begin to recognize some basic words. They")
                pdf.drawString(60.88,539.05,"may be able to:")
                pdf.drawString(96.9,517.76,"Identify basic vocabulary with visual support")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,477.51,"To improve their reading ability, students should:")
                pdf.drawString(96.9,456.21,"Learn and practice reading common words in")
                pdf.drawString(96.9,445.86,"familiar categories such as home, school,")
                pdf.drawString(96.9,435.51,"family, colors, body parts, animals, and")
                pdf.drawString(96.9,425.16,"actions")
                pdf.drawString(96.9,414.2,"Read short, simple sentences about familiar")
                pdf.drawString(96.9,403.86,"people, objects, and actions (example:")
                pdf.setFont('abcit', 9)
                pdf.drawString(253,403.86,"The")
                pdf.drawString(96.9,393.51,"boy is eating an apple.")
                pdf.setFont('abc', 9)
                pdf.drawString(186.96,393.51,")")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,353.27,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,342.91,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,332.57,"The student received "+str(rscore)+" on a scale of 100 to 109")

            def step1_2starreading(pdf):
                #2Star Image
                pdf.drawImage(step1_2starimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.29, 567.17, 279.36, 567.17)
                pdf.line(61.29, 433.11, 279.36, 433.11)
                pdf.line(61.29, 339.91, 279.36, 339.91)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(77.8,514,nokta)
                pdf.drawString(77.8,482.44,nokta)
                pdf.drawString(77.8,461.04,nokta)
                pdf.drawString(77.8,389.14,nokta)
                pdf.drawString(77.8,367.81,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,549.42,"Students begin to understand words and some short")
                pdf.drawString(60.88,539.05,"descriptions. They can:")
                pdf.drawString(96.9,517.76,"Understand common words in familiar")
                pdf.drawString(96.9,507.41,"categories such as home, school, family,")
                pdf.drawString(96.9,497.06,"colors, body parts, animals, and actions")
                pdf.drawString(96.9,486.1,"Recognize key words for understanding")
                pdf.drawString(96.9,475.75,"simple sentences")
                pdf.drawString(96.9,465.41,"Understand everyday actions in the present")
                pdf.drawString(96.9,454.45,"(")
                pdf.setFont('abcit', 9)
                pdf.drawString(99.9,454.45,"examples: The children play. He is eating.")
                pdf.setFont('abc', 9)
                pdf.drawString(266.98,454.45,")")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,414.21,"To improve their reading ability, students should:")
                pdf.drawString(96.9,392.9,"Learn vocabulary and common expressions")
                pdf.drawString(96.9,382.55,"used in social and familiar settings")
                pdf.drawString(96.9,371.6,"Practice reading simple sentences and short")
                pdf.drawString(96.9,361.25,"texts about familiar topics")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,321,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,310.65,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,300.3,"The student received "+str(rscore)+" on a scale of 100 to 109")

            def step1_3starreading(pdf):
                #3Star Image
                pdf.drawImage(step1_3starimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.29, 567.17, 279.36, 567.17)
                pdf.line(61.29, 370.4, 279.36, 370.4)
                pdf.line(61.29, 245.56, 279.36, 245.56)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(77.8,514,nokta)
                pdf.drawString(77.8,482.44,nokta)
                pdf.drawString(77.8,450.69,nokta)
                pdf.drawString(77.8,408.79,nokta)
                pdf.drawString(77.8,326.44,nokta)
                pdf.drawString(77.8,305.14,nokta)
                pdf.drawString(77.8,283.83,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,549.42,"Students understand short descriptions and find")
                pdf.drawString(60.88,539.05,"information in signs, forms, and schedules. They can:")
                pdf.drawString(96.9,517.76,"Understand common words and social")
                pdf.drawString(96.9,507.41,"expressions (")
                pdf.setFont('abcit', 9)
                pdf.drawString(150.42,507.41,"examples: play a game, go to a")
                pdf.drawString(96.9,497.06,"museum, wave goodbye")
                pdf.setFont('abc', 9)
                pdf.drawString(194.45,497.06,")")
                pdf.drawString(96.9,486.1,"Comprehend simple descriptions of current")
                pdf.drawString(96.9,475.75,"and past events (")
                pdf.setFont('abcit', 9)
                pdf.drawString(165.94,475.75,"examples: The mouse is on")
                pdf.drawString(96.9,465.41,"top of the table. He is washing his hands.")
                pdf.setFont('abc', 9)
                pdf.drawString(261,465.41,")")
                pdf.drawString(96.9,454.45,"Recognize relationships among words and")
                pdf.drawString(96.9,444.1,"phrases within familiar categories (")
                pdf.setFont('abcit', 9)
                pdf.drawString(235.45,444.1,"examples:")
                pdf.drawString(96.9,433.75,"food-fruit-strawberries; rain-sky-clouds; one")
                pdf.drawString(96.9,423.4,"more time-again")
                pdf.setFont('abc', 9)
                pdf.drawString(161.92,423.4,")")
                pdf.drawString(96.9,412.45,"Make connections across simple sentences")
                pdf.drawString(96.9,402.1,"(")
                pdf.setFont('abcit', 9)
                pdf.drawString(99.9,402.1,"example: Clouds are in the sky. Rain comes")
                pdf.drawString(96.9,391.75,"from them. Sometimes they cover the sun.")
                pdf.setFont('abc', 9)
                pdf.drawString(265.96,391.75,")")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.88,351.5,"To improve their reading ability, students should:")
                pdf.drawString(96.9,330.2,"Read longer paragraphs and stories about")
                pdf.drawString(96.9,319.85,"familiar people, objects, and information")
                pdf.drawString(96.9,308.9,"Learn more words that describe objects,")
                pdf.drawString(96.9,298.55,"places, people, actions, and ideas")
                pdf.drawString(96.9,287.59,"Speak or write in their own words about")
                pdf.drawString(96.9,277.24,"paragraphs, stories, and information they")
                pdf.drawString(96.9,266.9,"read")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,226.64,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,216.29,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,205.94,"The student received "+str(rscore)+" on a scale of 100 to 109")

            def step1_4starreading(pdf):
                #4Star Image
                pdf.drawImage(step1_4starimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.29, 567.38, 279.35, 567.38)
                pdf.line(61.29, 358.62, 279.35, 358.62)
                pdf.line(61.29, 221.26, 279.35, 221.26)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(77.8,514,nokta)
                pdf.drawString(77.8,472.21,nokta)
                pdf.drawString(77.8,430.2,nokta)
                pdf.drawString(77.8,388.2,nokta)
                pdf.drawString(77.8,314.65,nokta)
                pdf.drawString(77.8,303.7,nokta)
                pdf.drawString(77.8,282.39,nokta)
                pdf.drawString(77.8,261.09,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.65,549.62,"Students understand short descriptions, information in")
                pdf.drawString(60.65,539.27,"signs, and short messages. They can:")
                pdf.drawString(96.65,517.97,"Understand common words and some less")
                pdf.drawString(96.65,507.62,"common words about objects, places, people,")
                pdf.drawString(96.65,497.27,"actions, and ideas (")
                pdf.setFont('abcit', 9)
                pdf.drawString(174.7,497.27,"examples: ring,")
                pdf.drawString(96.65,486.92,"adventures, whisper, double")
                pdf.setFont('abc', 9)
                pdf.drawString(209.16,486.92,")")
                pdf.drawString(96.65,475.97,"Comprehend the meaning of complex")
                pdf.drawString(96.65,465.62,"sentences (")
                pdf.setFont('abcit', 9)
                pdf.drawString(143.15,465.6,"examples: This is a friendly thing")
                pdf.drawString(96.65,455.27,"to do when you say goodbye. People do this")
                pdf.drawString(96.65,444.92,"when they talk quietly.")
                pdf.setFont('abc', 9)
                pdf.drawString(185.13,444.92,")")
                pdf.drawString(96.65,433.96,"Connect information in longer sentences and")
                pdf.drawString(96.65,423.61,"across different sentences to infer")
                pdf.drawString(96.65,413.27,"information, identify main ideas, and")
                pdf.drawString(96.65,402.92,"understand the meaning of unfamiliar words.")
                pdf.drawString(96.65,391.96,"Locate key information in texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.65,339.72,"To improve their reading ability, students should:")
                pdf.drawString(96.65,318.41,"Study new, unfamiliar words")
                pdf.drawString(96.65,307.46,"Practice reading stories and informational")
                pdf.drawString(96.65,297.11,"texts about a variety of topics")
                pdf.drawString(96.65,286.15,"Practice reading longer and more complex")
                pdf.drawString(96.65,275.8,"texts")
                pdf.drawString(96.65,264.85,"Speak or write in their own words about")
                pdf.drawString(96.65,254.5,"stories and information they read")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.65,202.25,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.65,191.91,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.65,181.56,"The student received "+str(rscore)+" on a scale of 100 to 109")

            #Listening yıldıza göre bilgiler

            def step1_nslistening(pdf):
                #NS Image
                pdf.drawImage(step1_nsimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.64, 567.26, 543.2, 567.26)
                pdf.line(325.64, 497.02, 543.2, 497.02)
                pdf.line(325.64, 447.58, 543.2, 447.58)
                #Noktalar Yok
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,528.71,"The test taker did not respond to any questions in this")
                pdf.drawString(325.3,518.36,"section. Therefore, the scores for this section cannot")
                pdf.drawString(325.3,508.01,"be provided.")
                #Part2 Boş
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,428.68,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,407.98,"The student received "+str(lscore)+" on a scale of 100 to 109")

            def step1_1starlistening(pdf):
                #1Star Image
                pdf.drawImage(step1_1starimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.64, 567.26, 543.2, 567.26)
                pdf.line(325.64, 486.06, 543.2, 486.06)
                pdf.line(325.64, 341.05, 543.2, 341.05)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(342.2,503.65,nokta)
                pdf.drawString(342.2,443.25,nokta)
                pdf.drawString(342.2,411.8,nokta)
                pdf.drawString(342.2,400.64,nokta)
                pdf.drawString(342.2,368.99,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.4,"Students begin to recognize some familiar words in")
                pdf.drawString(325.3,539.06,"speech, such as words for objects, places, and people.")
                pdf.drawString(325.3,528.71,"They may be able to:")
                pdf.drawString(361.3,507.41,"Understand familiar words with visual support")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,468.32,"To improve their listening ability, students should:")
                pdf.drawString(361.3,447.01,"Learn everyday words for objects and people")
                pdf.drawString(361.3,436.66,"in familiar categories such as home, school,")
                pdf.drawString(361.3,426.31,"family, colors, body parts, and animals")
                pdf.drawString(361.3,415.36,"Use pictures to help learn new words")
                pdf.drawString(361.3,404.4,"Listen to short, simple sentences about")
                pdf.drawString(361.3,394.05,"everyday actions, objects, and people.")
                pdf.drawString(361.3,383.7,"(example:")
                pdf.setFont('abcit', 9)
                pdf.drawString(403.31,383.7,"She is swimming.")
                pdf.setFont('abc', 9)
                pdf.drawString(473.33,383.7,")")
                pdf.drawString(361.3,372.75,"Practice using common, everyday")
                pdf.drawString(361.3,362.4,"expressions, such as greetings")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,322.16,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,301.46,"The student received "+str(lscore)+" on a scale of 100 to 109")

            def step1_2starlistening(pdf):
                #2Star Image
                pdf.drawImage(step1_2starimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.64, 567.26, 543.2, 567.26)
                pdf.line(325.64, 454.41, 543.2, 454.41)
                pdf.line(325.64, 330.1, 543.2, 330.1)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(342.2,514,nokta)
                pdf.drawString(342.2,482.34,nokta)
                pdf.drawString(342.2,411.9,nokta)
                pdf.drawString(342.2,390.29,nokta)
                pdf.drawString(342.2,379.34,nokta)
                pdf.drawString(342.2,358.03,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.4,"Students begin to recognize some familiar words in")
                pdf.drawString(325.3,539.06,"speech. They can:")
                pdf.drawString(361.3,517.76,"Understand words for objects and people in")
                pdf.drawString(361.3,507.41,"familiar categories such as school, home,")
                pdf.drawString(361.3,496.45,"family, colors, body parts, and animals")  
                pdf.drawString(361.3,486.1,"Recognize action words in simple sentences")
                pdf.drawString(361.3,475.75,"(")
                pdf.setFont('abcit', 9)
                pdf.drawString(364.3,475.75,"examples: The children play. He is eating.")
                pdf.setFont('abc', 9)
                pdf.drawString(531.38,475.75,")")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,436.66,"To improve their listening ability, students should:")
                pdf.drawString(361.3,415.36,"Practice saying and listening to familiar words")
                pdf.drawString(361.3,405.01,"used in simple sentences")
                pdf.drawString(361.3,394.05,"Practice having short, simple conversations")
                pdf.drawString(361.3,383.1,"Practice listening to messages spoken by")
                pdf.drawString(361.3,372.75,"teachers, friends, and family")
                pdf.drawString(361.3,361.79,"Begin listening to and identifying basic")
                pdf.drawString(361.3,351.44,"information in short, simple stories")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,311.2,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,290.5,"The student received "+str(lscore)+" on a scale of 100 to 109")

            def step1_3starlistening(pdf):
                #3Star Image
                pdf.drawImage(step1_3starimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.64, 567.26, 543.2, 567.26)
                pdf.line(325.64, 401.46, 543.2, 401.46)
                pdf.line(325.64, 266.8, 543.2, 266.8)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(342.2,514,nokta)
                pdf.drawString(342.2,492.69,nokta)
                pdf.drawString(342.2,461.04,nokta)
                pdf.drawString(342.2,439.74,nokta)
                pdf.drawString(342.2,358.64,nokta)
                pdf.drawString(342.2,337.33,nokta)
                pdf.drawString(342.2,316,nokta)
                pdf.drawString(342.2,294.73,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.4,"Students understand short, simple descriptions,")
                pdf.drawString(325.3,539.06,"conversations, and messages. They can:")
                pdf.drawString(361.3,517.76,"Understand common expressions used in")
                pdf.drawString(361.3,507.41,"everyday conversations")
                pdf.drawString(361.3,496.45,"Understand a simple, single instruction")  
                pdf.drawString(361.3,486.1,"spoken in familiar words, with key words")
                pdf.drawString(361.3,475.75,"repeated")
                pdf.drawString(361.3,464.8,"Understand the purpose of messages in")
                pdf.drawString(361.3,454.45,"which key information is repeated")
                pdf.drawString(361.3,443.5,"Understand the main ideas of simple stories")
                pdf.drawString(361.3,433.15,"in which key information is explicitly stated")
                pdf.drawString(361.3,422.8,"and repeated")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,383.7,"To improve their listening ability, students should:")
                pdf.drawString(361.3,362.4,"Study more words that describe familiar")
                pdf.drawString(361.3,352.05,"topics, settings, and actions")
                pdf.drawString(361.3,341.09,"Practice using less common words and")
                pdf.drawString(361.3,330.74,"expressions in conversations")
                pdf.drawString(361.3,319.74,"Listen to age-appropriate academic talks and")
                pdf.drawString(361.3,309.44,"longer stories")
                pdf.drawString(361.3,298.49,"Speak or write in their own words about")
                pdf.drawString(361.3,288.14,"stories and information they listen to")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,247.89,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,227.29,"The student received "+str(lscore)+" on a scale of 100 to 109")

            def step1_4starlistening(pdf):
                #4Star Image
                pdf.drawImage(step1_4starimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.64, 567.38, 543.56, 567.38)
                pdf.line(325.64, 368.37, 543.56, 368.37)
                pdf.line(325.64, 243.01, 543.56, 243.01)
                #Noktalar
                pdf.setFont('dot', 17)
                pdf.drawString(341.95,514,nokta)
                pdf.drawString(341.95,482.56,nokta)
                pdf.drawString(341.95,461.25,nokta)
                pdf.drawString(341.95,439.95,nokta)
                pdf.drawString(341.95,408.29,nokta)
                pdf.drawString(341.95,325.55,nokta)
                pdf.drawString(341.95,304.24,nokta)
                pdf.drawString(341.95,282.94,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.05,549.62,"Students understand simple descriptions, instructions,")
                pdf.drawString(325.05,539.27,"conversations, and messages. They can:")
                pdf.drawString(361.05,517.97,"Understand less common words that describe")
                pdf.drawString(361.05,507.62,"familiar topics, settings, and actions")
                pdf.drawString(361.05,497.27,"(")  
                pdf.setFont('abcit', 9)
                pdf.drawString(364.05,497.27,"examples: pocket, pour, lamp, branch")
                pdf.setFont('abc', 9)
                pdf.drawString(514.12,497.6,")")
                pdf.drawString(361.05,486.32,"Understand indirect responses to questions in")
                pdf.drawString(361.05,475.97,"conversations")
                pdf.drawString(361.05,465.01,"Understand messages in which information is")
                pdf.drawString(361.05,454.66,"not explicitly stated")
                pdf.drawString(361.05,443.71,"Connect information to infer the main idea or")
                pdf.drawString(361.05,433.36,"topic of messages, stories, and informational")
                pdf.drawString(361.05,423.01,"texts")
                pdf.drawString(361.05,412.05,"Synthesize information from multiple locations")
                pdf.drawString(361.05,401.7,"in a longer spoken text")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.05,350.61,"To improve their listening ability, students should:")
                pdf.drawString(361.05,329.31,"Learn new, unfamiliar words they hear in")
                pdf.drawString(361.05,318.96,"longer stories and academic talks")
                pdf.drawString(361.05,308,"Practice using less common words and")
                pdf.drawString(361.05,297.65,"expressions in conversations")
                pdf.drawString(361.05,286.7,"Speak or write in their own words about")
                pdf.drawString(361.05,276.35,"stories and information they listen to")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(324.7,224,"CEFR Level: "+str(lcefr))
                pdf.drawString(324.7,204,"The student received "+str(lscore)+" on a scale of 100 to 109")

            #Belge özellikleri
            fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_ScoreReport.PDF")
            documentTitle = 'Document title!'
            pdf = canvas.Canvas(fileName,pagesize=(595.27,841.89))
            pdf.setTitle(documentTitle)
            #Font
            pdfmetrics.registerFont(TTFont('abc', './data/arial.ttf'))
            pdfmetrics.registerFont(TTFont('abcbold', './data/arialbold.ttf'))
            pdfmetrics.registerFont(TTFont('abcit', './data/arialitalic.ttf'))
            pdfmetrics.registerFont(TTFont('dot', './data/comic.ttf'))

            #Ana baskı alanı
            def baskialani(pdf):
                #Sınav Türü
                pdf.setFont('abcbold', 22)
                pdf.drawString(480.24,767.3 ,str(sinavturu))
                #Öğrenci Bilgisi
                pdf.setFont('abc', 9.5)
                pdf.drawString(60.45,739.05,"Student Name:  "+str(studentname))
                pdf.drawString(61.15,722.93,"Student Number:  "+str(studentnumber))
                pdf.drawString(61.15,706.54,"Test Date:  "+str(testdate))
                pdf.drawString(422.9,739.05,"Date of Birth:  "+str(dateofbirth))
                pdf.drawString(422.9,723.11,"Gender:  "+str(lgender[4]))
                #Alt bilgi Okul vs.
                pdf.setFont('abc', 6.5)
                #BURAYA DİKKAT ORTALAMAK LAZIM
                pdf.drawCentredString(302.61,38.75,str(school)+", Turkey")
                pdf.drawCentredString(302.61,31.28,"OYD - Okul Yayin, Turkey")
                pdf.drawCentredString(302.61,23.8,str(lname[0]))

            #Anabaskı kodu
            baskialani(pdf)

            #Reading skora göre sonuç seçimi
            if rscore == "NS":
                step1_nsreading(pdf)
            else:
                if rscore == str("100"):
                    step1_1starreading(pdf)
                else:
                    if rscore ==  str("101"):
                        step1_2starreading(pdf)
                    else:
                        if rscore == str("102"):
                            step1_2starreading(pdf)
                        else:
                            if rscore == str("103"):
                                step1_2starreading(pdf)
                            else:
                                if rscore == str("104"):
                                    step1_3starreading(pdf)
                                else:
                                    if rscore ==  str("105"):
                                        step1_3starreading(pdf)
                                    else:
                                        if rscore == str("106"):
                                            step1_3starreading(pdf)
                                        else:
                                            if rscore == str("107"):
                                                step1_4starreading(pdf)
                                            else:
                                                if rscore == str("108"):
                                                    step1_4starreading(pdf)
                                                else:
                                                    if rscore == str("109"):
                                                        step1_4starreading(pdf)

            #Listening skora göre sonuç seçimi
            if lscore == "NS":
                step1_nslistening(pdf)
            else:
                if lscore == str("100"):
                    step1_1starlistening(pdf)
                else:
                    if lscore ==  str("101"):
                        step1_2starlistening(pdf)
                    else:
                        if lscore == str("102"):
                            step1_2starlistening(pdf)
                        else:
                            if lscore == str("103"):
                                step1_2starlistening(pdf)
                            else:
                                if lscore == str("104"):
                                    step1_3starlistening(pdf)
                                else:
                                    if lscore ==  str("105"):
                                        step1_3starlistening(pdf)
                                    else:
                                        if lscore == str("106"):
                                            step1_3starlistening(pdf)
                                        else:
                                            if lscore == str("107"):
                                                step1_4starlistening(pdf)
                                            else:
                                                if lscore == str("108"):
                                                    step1_4starlistening(pdf)
                                                else:
                                                    if lscore == str("109"):
                                                        step1_4starlistening(pdf)

            pdf.save()

        def step1_certificate():
            #Reading sonuçları
            def step1c_nsreading(pdf):
                pdf.drawImage(step1c_nsimage, 418.6295,225.6, width=19.3,height=18.2,mask=None)

            def step1c_1starreading(pdf):
                pdf.drawImage(step1c_1starimage, 418.3,225.95, width=80,height=15.9,mask=None)

            def step1c_2starreading(pdf):
                pdf.drawImage(step1c_2starimage, 418.3,225.95, width=80,height=15.9,mask=None)

            def step1c_3starreading(pdf):
                pdf.drawImage(step1c_3starimage, 418.3,225.95, width=80,height=15.9,mask=None)

            def step1c_4starreading(pdf):
                pdf.drawImage(step1c_4starimage, 418.3,225.95, width=80,height=15.9,mask=None)

            #Listening sonuçları
            def step1c_nslistening(pdf):
                pdf.drawImage(step1c_nsimage, 418.6295,184.8, width=19.3,height=18.2,mask=None)

            def step1c_1starlistening(pdf):
                pdf.drawImage(step1c_1starimage, 418.9,185, width=80,height=15.9,mask=None)

            def step1c_2starlistening(pdf):
                pdf.drawImage(step1c_2starimage, 418.9,185, width=80,height=15.9,mask=None)

            def step1c_3starlistening(pdf):
                pdf.drawImage(step1c_3starimage, 418.9,185, width=80,height=15.9,mask=None)

            def step1c_4starlistening(pdf):
                pdf.drawImage(step1c_4starimage, 418.9,185, width=80,height=15.9,mask=None)

            #Belge özellikleri
            fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_Certificate.PDF")
            documentTitle = 'Document title!'
            pdf = canvas.Canvas(fileName,pagesize=(841.89,595.27))
            pdf.setTitle(documentTitle)
            #Font
            pdfmetrics.registerFont(TTFont('abc', './data/arial.ttf'))
            pdfmetrics.registerFont(TTFont('abcbold', './data/arialbold.ttf'))
            pdfmetrics.registerFont(TTFont('abcit', './data/arialitalic.ttf'))
            pdfmetrics.registerFont(TTFont('dot', './data/comic.ttf'))

            #Ana baskı alanı
            def baskialani(pdf):
                #Sınav Türü
                pdf.setFont('abcbold', 24)
                pdf.drawCentredString(402.87,430.64 ,str(sinavturu))
                #Sertifika içeriği
                pdf.setFont('abcbold', 32)
                pdf.drawString(203.69,387.23,"Certificate of Achievement")
                pdf.setFont('abc', 18)
                pdf.drawString(325.5,363.57,"This is to certify that")
                pdf.setFont('abcbold', 20)
                pdf.drawCentredString(404.12,314.84,str(studentname))
                pdf.setFont('abcit', 18)
                pdf.drawString(153.74,271.1,"Has earned the following levels on the TOEFL")
                pdf.drawString(532.5,271.1,"Primary")
                pdf.drawString(617.38,271.1,"Test")
                pdf.drawString(594.99,271.5,"™")
                pdf.setFont('abc', 11.50)
                pdf.drawString(520.97,277.32,"®")
                pdf.setFont('abc', 18)
                pdf.drawString(319.85,229.37,"Reading:")
                pdf.drawString(319.85,190.67,"Listening:")
                #Alt bilgi Okul vs.
                pdf.setFont('abc', 8)
                pdf.drawString(79.7,115.1,"Test Date: "+str(testdate))
                pdf.drawString(79.7,105.9,str(school)+", Turkey")
                pdf.drawString(79.7,97.37,"OYD - Okul Yayin, Turkey")
                pdf.setFont('abc', 7)
                pdf.drawString(79.7,89.07,str(lname[0]))
                
            #Anabaskı kodu
            baskialani(pdf)

            #Reading skora göre sonuç seçimi
            if rscore == "NS":
                step1c_nsreading(pdf)
            else:
                if rscore == str("100"):
                    step1c_1starreading(pdf)
                else:
                    if rscore ==  str("101"):
                        step1c_2starreading(pdf)
                    else:
                        if rscore == str("102"):
                            step1c_2starreading(pdf)
                        else:
                            if rscore == str("103"):
                                step1c_2starreading(pdf)
                            else:
                                if rscore == str("104"):
                                    step1c_3starreading(pdf)
                                else:
                                    if rscore ==  str("105"):
                                        step1c_3starreading(pdf)
                                    else:
                                        if rscore == str("106"):
                                            step1c_3starreading(pdf)
                                        else:
                                            if rscore == str("107"):
                                                step1c_4starreading(pdf)
                                            else:
                                                if rscore == str("108"):
                                                    step1c_4starreading(pdf)
                                                else:
                                                    if rscore == str("109"):
                                                        step1c_4starreading(pdf)

            #Listening skora göre sonuç seçimi
            if lscore == "NS":
                step1c_nslistening(pdf)
            else:
                if lscore == str("100"):
                    step1c_1starlistening(pdf)
                else:
                    if lscore ==  str("101"):
                        step1c_2starlistening(pdf)
                    else:
                        if lscore == str("102"):
                            step1c_2starlistening(pdf)
                        else:
                            if lscore == str("103"):
                                step1c_2starlistening(pdf)
                            else:
                                if lscore == str("104"):
                                    step1c_3starlistening(pdf)
                                else:
                                    if lscore ==  str("105"):
                                        step1c_3starlistening(pdf)
                                    else:
                                        if lscore == str("106"):
                                            step1c_3starlistening(pdf)
                                        else:
                                            if lscore == str("107"):
                                                step1c_4starlistening(pdf)
                                            else:
                                                if lscore == str("108"):
                                                    step1c_4starlistening(pdf)
                                                else:
                                                    if lscore == str("109"):
                                                        step1c_4starlistening(pdf)

            pdf.save()
            
        def step1_tk():
            if rscore != "NS" and lscore !="NS":
                totalscore = str(math.ceil((int(rscore)+int(lscore))/2))
                
                #Reading yıldıza göre bilgiler
                def step1tk_1starreading(pdf):
                    pdf.drawString(36.85,702.01,"Students begin to recognize some basic words. They may be able to;")
                    pdf.drawString(36.85,690.67,"- Identify basic vocabulary with visual support")
                    pdf.drawString(36.85,667.99,"Öğrenci bazı temel kelimeleri tanımaya başlayabilir.Okuma becerileri:")
                    pdf.drawString(36.85,656.66,"- Görsel destekle temel kelimeleri tanıma")
                    pdf.drawString(36.85,639.65,"Next Steps")
                    pdf.drawString(36.85,625.47,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,614.14,"- Learn and practice reading common words in familiar categories such as home, school, family, colors, body parts, animals, and actions")
                    pdf.drawString(36.85,602.8,"- Read short, simple sentences about familiar people, objects, and action")
                    pdf.drawString(36.85,591.46,"Okuma becerilerini geliştirmek için öğrenci:")
                    #Eksik Var!!!!!!!!!
                    pdf.drawString(36.85,557.44,"- Tanıdık insanlar, nesneler ve eylemler hakkında kısa, basit cümleler okuyabilir.(örnek: Çocuk elma yiyor.)")
                    
                def step1tk_2starreading(pdf):
                    pdf.drawString(36.85,702.01,"Students begin to understand words and some short descriptions. They can;")
                    pdf.drawString(36.85,690.67,"- Understand common words in familiar categories such as home, school, family, colors, body parts, animals, and actions")
                    pdf.drawString(36.85,679.33,"- Recognize key words for understanding simple sentences")
                    pdf.drawString(36.85,667.99,"- Understand everyday actions in the present (examples: The children play. He is eating.)")
                    pdf.drawString(36.85,645.32,"Öğrenci kelimeleri ve bazı kısa açıklamaları anlamaya başlar.Okuma becerileri:")
                    pdf.drawString(36.85,633.98,"- Ev, okul, aile, renkler, vücudun bölümleri, hayvanlar ve eylemler gibi bilindik kategorilerdeki yaygın kullanılan kelimeleri anlama")
                    pdf.drawString(36.85,622.64,"- Basit cümleleri anlamak için gerekli anahtar kelimeleri tanıma")
                    pdf.drawString(36.85,611.3,"- Günlük eylemleri anlama (örnekler: Çocuklar oynar. O yemek yiyor.)")
                    pdf.drawString(36.85,594.29,"Next Steps")
                    pdf.drawString(36.85,580.12,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,568.78,"- Learn vocabulary and common expressions used in social and familiar settings")
                    pdf.drawString(36.85,557.44,"- Practice reading simple sentences and short texts about familiar topics")
                    pdf.drawString(36.85,534.77,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,523.43,"- Sosyal ve tanıdık ortamlarda yaygın olarak kullanılan kelime ve ifadeleri öğrenebilir.")
                    pdf.drawString(36.85,512.09,"- Basit cümleler ve tanıdık konularda kısa okuma alıştırmaları yapabilir.")
                    
                def step1tk_3starreading(pdf):
                    pdf.drawString(36.85,702.01,"Students understand short descriptions and find information in signs, forms, and schedules. They can;")
                    pdf.drawString(36.85,690.67,"- Understand common words and social expressions (examples: play a game, go to a museum, wave goodbye)")
                    pdf.drawString(36.85,679.33,"- Comprehend simple descriptions of current and past events (examples: The mouse is on top of the table. He is washing his hands.)")
                    pdf.drawString(36.85,667.99,"- Recognize relationships among words and phrases within familiar categories (examples: food,fruit,strawberries; rain,sky,clouds;)")
                    pdf.drawString(36.85,656.66,"- Make connections across simple sentences (example: Clouds are in the sky. Rain comes from them. Sometimes they cover the sun.)")
                    pdf.drawString(36.85,633.98,"Öğrenci, kısa açıklamaları, tabela, levha, bilgilendirme amaçlı belge ve çizelgelerdeki bilgileri anlayabilir.Okuma becerileri:")
                    pdf.drawString(36.85,622.64,"- Sık kullanılan kelimeleri ve sosyal yaşama dair ifadeleri anlama, (örnekler: oyun oynamak, müzeye gitmek, veda etmek)")
                    pdf.drawString(36.85,611.30,"- Güncel ve geçmişte yasanmış olayların basit tanımlamalarını kavrama, (örnekler: -Fare masanın üstünde. -Ellerini yıkıyor.)")
                    pdf.drawString(36.85,599.96,"- Aynı kategoriye ait kelimeler ve kelime öbekleri arasındaki ilişkileri anlama, (örnekler: -gıda,meyve,çilek; yağmur,gök,bulutlar;)")
                    pdf.drawString(36.85,588.62,"- Basit cümleler arasında bağlantı kurma (örnek: Bulutlar gökyüzünde. Yağmur onlardan geliyor. Bazen güneşi kapatıyorlar.)")
                    pdf.drawString(36.85,571.62,"Next Steps")
                    pdf.drawString(36.85,557.44,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,546.10,"- Read longer paragraphs and stories about familiar people, objects, and information")
                    pdf.drawString(36.85,534.77,"- Learn more words that describe objects, places, people, actions, and ideas")
                    pdf.drawString(36.85,523.43,"- Speak or write in their own words about paragraphs, stories, and information they read")
                    pdf.drawString(36.85,500.75,"Okuma yeteneğini geliştirmek için öğrenciler;")
                    pdf.drawString(36.85,489.41,"- Tanıdık insanlar, nesneler ve bilgiler hakkında daha uzun paragraflar ve hikayeler okuyabilir.")
                    pdf.drawString(36.85,478.07,"- Nesneleri, yerleri, insanları, eylemleri ve fikirleri ifade etmek için daha fazla kelime öğrenebilir.")
                    pdf.drawString(36.85,466.73,"- Okuduğu paragraflar, hikayeler ve bilgiler hakkında kendi cümleleriyle anlatma veya yazma alıştırmaları yapabilir.")
                    
                def step1tk_4starreading(pdf):
                    pdf.setFont('abc', 8)
                    pdf.drawString(36.85,702.01,"Students understand short descriptions, information in signs, and short messages. They can;")
                    pdf.drawString(36.85,690.67,"- Understand common words and some less common words about objects, places, people, actions, and ideas, (examples: ring, adventures etc)")
                    pdf.drawString(36.85,679.33,"- Comprehend the meaning of complex sentences (e.g. This is a friendly thing to do when you say goodbye.People do this when they talk quietly.)")
                    pdf.drawString(36.85,667.99,"- Connect information in longer sentences and across different sentences to infer information, identify main ideas,")
                    pdf.drawString(36.85,656.66,"- Understand the meaning of unfamiliar words and Locate key information in texts")
                    pdf.drawString(36.85,633.98,"Öğrenci, kısa açıklamaları, tabela ve levhalardaki bilgileri ve kısa mesajları anlar. Okuma becerileri:")
                    pdf.drawString(36.85,622.64,"- Nesneler, yerler, insanlar, eylemler ve fikirler hakkında, daha az yaygın kullanılan kelimeleri anlama(örn: yüzük, macera,)")
                    pdf.drawString(36.8,611.30,"- Karmaşık cümlelerin anlamını kavrama.(örn: Güle güle derken bunu yapmak dostça bir tavır.İnsanlar bunu sessizce konuituklarında yapar.")
                    pdf.drawString(36.85,599.96,"- Uzun cümlelerde cümle içerisindeki bilgilerle veya paragraflarda cümleler arasındaki bilgilerle bağlantı kurarak, bilgi çıkarımında bulunma")
                    pdf.drawString(36.85,588.62,"- Ana fikirleri bulma ve bilmediği kelimelerin anlamlarını çıkarma ve metinlerdeki önemli bilgileri tespit etme")
                    pdf.drawString(36.85,571.62,"Next Steps")
                    pdf.drawString(36.85,557.44,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,546.1,"- Study new, unfamiliar words")
                    pdf.drawString(36.85,534.77,"- Practice reading stories and informational texts about a variety of topics")
                    pdf.drawString(36.85,523.43,"- Practice reading longer and more complex texts")
                    pdf.drawString(36.85,512.09,"- Speak or write in their own words about stories and information they read")
                    pdf.drawString(36.85,489.41,"Okuma becerilerini geliştirmek için, öğrenci:")
                    pdf.drawString(36.85,478.07,"- Yeni, yabancı kelimeler üzerinde çalışabilir.")
                    pdf.drawString(36.85,466.73,"- Hikaye ve çesitli konular üzerine bilgilendirici metinler okuyabilir.")
                    pdf.drawString(36.85,455.40,"- Daha uzun ve daha karmasşık metinlerle okuma alıştırması yapabilir.")
                    pdf.drawString(36.85,444.06,"- Okuduğu hikayeleri ve edindiği bilgileri kendi cümleleriyle anlatma veya yazma alıştırmaları yapabilir.")

                #Listening yıldıza göre bilgiler
                b2 = 11.34
                b1 = 384.53
                b3 = 17.01
                b4 = 14.16
                def step1tk_1starlistening(pdf):
                    pdf.drawString(36.85,b1,"Students begin to recognize some familiar words in speech, such as words for objects, places, and people. They may be able to;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand familiar words with visual support")
                    pdf.drawString(36.85,b1-b2*3,"Öğrenci, nesneler, yerler ve insanları tanımlamak için kullanılan bazı temel kelimeleri tanımaya başlayabilir.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*4,"- Görsel destek yardımıyla temel kelimeleri anlama")
                    pdf.drawString(36.85,b1-b3-b2*4,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*4,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*5,"- Learn everyday words for objects and people in familiar categories such as home, school, family, colors, body parts, and animals")
                    pdf.drawString(36.85,b1-b3-b4-b2*6,"- Use pictures to help learn new words")
                    pdf.drawString(36.85,b1-b3-b4-b2*7,"- Listen to short, simple sentences about everyday actions, objects, and people. (example: She is swimming.)")
                    pdf.drawString(36.85,b1-b3-b4-b2*8,"- Practice using common, everyday expressions, such as greetings")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Ev, okul, aile, renkler, vücudun bölümleri ve hayvanlar gibi temel kategorilerdeki nesneler ve insanları tanımlamak için yaygın olarak kullanılan")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"kelimeleri ögrenebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Yeni kelimeler öğrenmesine yardımcı olması için resimleri kullanabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Günlük hayattaki eylemler, nesneler ve kisiler hakkında kısa, basit cümleler dinleyebilir.(örnek: Yüzüyor.)")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Selamlaşma ifadeleri gibi günlük hayatta yaygın olarak kullanılan ifadeleri kullanarak pratik yapabilir.")
                    
                def step1tk_2starlistening(pdf):
                    pdf.drawString(36.85,b1,"Students begin to recognize some familiar words in speech. They can;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand words for objects and people in familiar categories such as school, home, family, colors, body parts, and animals")
                    pdf.drawString(36.85,b1-b2*2,"- Recognize action words in simple sentences (examples: The children play. He is eating.)")
                    pdf.drawString(36.85,b1-b2*4,"Öğrenci konuşmalardaki bazı tanıdık kelimeleri anlamaya baslar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*5,"- Okul, ev, aile, renkler, vücudun bölümleri ve hayvanlar gibi tanıdık kategorilerdeki nesne ve insanlarla ilgili kelimeleri anlama")
                    pdf.drawString(36.85,b1-b2*6,"- Temel eylemleri basit cümle içinde tanıma (örnekler: -Çocuklar oynar. -Yemek yiyor.)")
                    pdf.drawString(36.85,b1-b3-b2*6,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*6,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*7,"- Practice saying and listening to familiar words used in simple sentences")
                    pdf.drawString(36.85,b1-b3-b4-b2*8,"- Practice having short, simple conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*9,"- Practice listening to messages spoken by teachers, friends, and family")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"- Begin listening to and identifying basic information in short, simple stories")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Basit cümlelerde kullanılan tanıdık kelimeleri söyleme ve dinleme alıştırması yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Kısa ve basit konuşmalarla pratik yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Öğretmenler, arkadaşlar ve aile bireyleri tarafından söylenen mesajları dinleyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"- Kısa ve basit hikayelerdeki temel bilgileri dinlemeye ve tanımlamaya çalısabilir.")
                    
                def step1tk_3starlistening(pdf):
                    pdf.drawString(36.85,b1,"Students understand short, simple descriptions, conversations, and messages.")
                    pdf.drawString(36.85,b1-b2*1,"- Understand common expressions used in everyday conversations")
                    pdf.drawString(36.85,b1-b2*2,"- Understand a simple, single instruction spoken in familiar words, with key words repeated")
                    pdf.drawString(36.85,b1-b2*3,"- Understand the purpose of messages in which key information is repeated")
                    pdf.drawString(36.85,b1-b2*4,"- Understand the main ideas of simple stories in which key information is explicitly stated and repeated")
                    pdf.drawString(36.85,b1-b2*6,"Öğrenci, kısa ve basit açıklamaları, konuşmaları ve mesajları anlar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*7,"- Günlük konuşmalarda yaygın olarak kullanılan ifadeleri anlama")
                    pdf.drawString(36.85,b1-b2*8,"- Anahtar kelimeler tekrarlanarak konuşulan, basit, tanıdık kelimelerden olusan tek bir talimatı anlama")
                    pdf.drawString(36.85,b1-b2*9,"- Önemli olan bilginin tekrarlandığı mesajların amacını anlama")
                    pdf.drawString(36.85,b1-b2*10,"- Önemli olan bilginin açıkça ifade edildigi ve tekrarlandığı, basit hikayelerin ana fikrini anlama")
                    pdf.drawString(36.85,b1-b3-b2*10,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Study more words that describe familiar topics, settings, and actions")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Listen to age-appropriate academic talks and longer stories")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"Dinleme becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,b1-b3-b4-b2*17,"- Tanıdık konuları, durumları ve eylemleri tanımlayan daha fazla kelime öğrenebilir")
                    pdf.drawString(36.85,b1-b3-b4-b2*18,"- Konuşmalarda, yeni öğrendiği kelime ve ifadeleri kullanmaya çalışabilir")
                    pdf.drawString(36.85,b1-b3-b4-b2*19,"- Yaşına uygun akademik konuşmaları ve daha uzun hikayeleri dinleyebilir")
                    pdf.drawString(36.85,b1-b3-b4-b2*20,"- Dinlediği hikayeler ve edindiği bilgileri kendi cümleleriyle sözel veya yazılı ifade etme alıştırmaları yapabilir.")
                    
                def step1tk_4starlistening(pdf):
                    pdf.drawString(36.85,384.53,"Students understand simple descriptions, instructions, conversations, and messages. They can;")
                    pdf.drawString(36.85,373.19,"- Understand less common words that describe familiar topics, settings, and actions (examples: pocket, pour, lamp, branch)")
                    pdf.drawString(36.85,361.85,"- Understand indirect responses to questions in conversations")
                    pdf.drawString(36.85,350.51,"- Understand messages in which information is not explicitly stated")
                    pdf.drawString(36.85,339.18,"- Connect information to infer the main idea or topic of messages, stories, and informational texts")
                    pdf.drawString(36.85,327.84,"- Synthesize information from multiple locations in a longer spoken text")
                    pdf.drawString(36.85,305.16,"Öğrenciler basit açıklamalar, talimatlar, konuşmaları ve mesajları anlar.Dinleme becerileri:")
                    pdf.drawString(36.85,293.82,"- Yaygın kullanılmayan ve bilinen konular, kurgular ve eylemleri anlatan sözcükleri anlayabilir.(örnekler: cep, dökmek, lamba, şube)")
                    pdf.drawString(36.85,282.48,"- Konuşmalarda sorulara verilen dolaylı yanıtları anlama")
                    pdf.drawString(36.85,271.14,"- Açıkça belirtilmemiş bilgiler içeren mesajları anlama")
                    pdf.drawString(36.85,259.80,"- Mesajlar, hikayeler ve bilgilendirme amaçlı metinlerde ana fikri veya konuyu anlamak için bilgiler arasında bağlantı kurma")
                    pdf.drawString(36.85,248.47,"- Uzun konuşma metinlerinin içinde, farklı yerlerde geçen bilgileri sentezleme")
                    pdf.drawString(36.85,231.46,"Next Steps")
                    pdf.drawString(36.85,217.29,"To improve their listening ability, students should;")
                    pdf.drawString(36.85,205.95,"- Learn new, unfamiliar words they hear in longer stories and academic talks")
                    pdf.drawString(36.85,194.61,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,183.27,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,160.59,"Dinleme becerilerini geliştirmek için, öğrenci:")
                    pdf.drawString(36.85,149.25,"- Daha uzun metin ve akademik konusmalarda duydukları yeni ve yabancı kelimeleri öğrenebilir.")
                    pdf.drawString(36.85,137.92,"- Karşılıklı konuşmalarda yeni öğrendiği kelime ve ifadeleri kullanarak alıştırma yapabilir")
                    pdf.drawString(36.85,126.58,"- Dinlediği hikayeler ve edindiği bilgileri kendi cümleleriyle sözel veya yazılı ifade etme alıştırmaları yapabilir.")
                    
                #Belge özellikleri
                if str(totalscore)!=str("NS"):
                    fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_TurkceKarne.PDF")
                    documentTitle = 'Document title!'
                    pdf = canvas.Canvas(fileName,pagesize=(595.28,841.89))
                    pdf.setTitle(documentTitle)
                    #Font
                    pdfmetrics.registerFont(TTFont('abc', './data/tahoma.ttf'))
                    
                    #Ana baskı alanı
                    def baskialani(pdf):
                        #Arkaplan
                        pdf.drawImage(primarytk, 0,0, width=595,height=842,mask=None)
                        #Çizgiler
                        pdf.setLineWidth(0.665)
                        pdf.line(32.75, 754.02, 562.98, 754.02)
                        pdf.line(32.75, 31.19, 562.98, 31.19)
                        pdf.line(32.75, 119.06, 562.98, 119.06)
                        pdf.line(32.75, 436.54, 562.98, 436.54)
                        pdf.line(32.6, 754.35, 32.6, 30.87)
                        pdf.line(562.98, 754.35, 562.98, 30.87)
                        #Öğrenci Bilgisi
                        pdf.setFont('abc', 8)
                        pdf.drawString(123.31,804.06,"Student Name:")
                        pdf.drawString(123.31,789.88,"Test Date:")
                        pdf.drawString(123.31,775.71,"Class:")
                        pdf.drawString(272.13,804.06,"Student Number:")
                        pdf.drawString(272.13,789.88,"School Name:")
                        pdf.drawString(272.13,775.71,"Test Name:")        
                        pdf.drawString(180,804.06,str(studentname))
                        pdf.drawString(180,789.88,str(testdate))
                        pdf.drawString(180,775.71,str(studentclass))
                        pdf.drawString(334.49,804.06,str(studentnumber))
                        pdf.drawString(334.49,789.88,str(school))
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 1")
                        pdf.drawString(257.95,761.54,"Total Score : "+str(totalscore))
                        #Bazı içerikler
                        pdf.setFont('abc', 10)
                        pdf.drawString(45.35,741.09,"READING - OKUMA")
                        pdf.drawString(45.35,423.61,"LISTENING - DİNLEME")
                        pdf.drawString(45.35,106.13,"TOEFL History - TOEFL Geçmişi")
                        pdf.setFont('abc', 8)
                        pdf.drawString(160,741.69,"Score: "+str(rscore)+",")
                        pdf.drawString(210,741.69,"CEFR: "+str(rcefr)+",")
                        pdf.drawString(260,741.69,"Lexile: "+str(lexile))
                        pdf.drawString(527.25,741.69,"Stars: "+str(rstar[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(527.25,424.21,"Stars: "+str(lstar[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(rscore)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(lscore)+" puan almıştır.)")
                        

                    #Anabaskı kodu
                    baskialani(pdf)

                    #Reading skora göre sonuç seçimi
                    if rscore == str("100"):
                        step1tk_1starreading(pdf)
                    else:
                        if rscore == str("101"):
                            step1tk_2starreading(pdf)
                        elif rscore == str("102"):
                            step1tk_2starreading(pdf)
                        elif rscore == str("103"):
                            step1tk_2starreading(pdf)
                        elif rscore == str("104"):
                            step1tk_3starreading(pdf)
                        elif rscore == str("105"):
                            step1tk_3starreading(pdf)
                        elif rscore == str("106"):
                            step1tk_3starreading(pdf)
                        elif rscore == str("107"):
                            step1tk_4starreading(pdf)
                        elif rscore == str("108"):
                            step1tk_4starreading(pdf)
                        elif rscore == str("109"):
                            step1tk_4starreading(pdf)

                    #Listening skora göre sonuç seçimi
                    if lscore == str("100"):
                        step1tk_1starlistening(pdf)
                    else:
                        if lscore == str("101"):
                            step1tk_2starlistening(pdf)
                        elif lscore == str("102"):
                            step1tk_2starlistening(pdf)
                        elif lscore == str("103"):
                            step1tk_2starlistening(pdf)
                        elif lscore == str("104"):
                            step1tk_3starlistening(pdf)
                        elif lscore == str("105"):
                            step1tk_3starlistening(pdf)
                        elif lscore == str("106"):
                            step1tk_3starlistening(pdf)
                        elif lscore == str("107"):
                            step1tk_4starlistening(pdf)
                        elif lscore == str("108"):
                            step1tk_4starlistening(pdf)
                        elif lscore == str("109"):
                            step1tk_4starlistening(pdf)
                            
                    pdf.save()
                    
            #Türkçe karne listeningi ns, readingi ns olmayan sonuç
            if lscore == "NS" and rscore != "NS":
                lscore2 = "100"
                totalscore = str(math.ceil((int(rscore)+int(lscore2))/2))

                #Listening yıldıza göre bilgiler
                b2 = 11.34
                b1 = 384.53
                b3 = 17.01
                b4 = 14.16
                def step1tk_1starlistening(pdf):
                    pdf.drawString(36.85,b1,"Students begin to recognize some familiar words in speech, such as words for objects, places, and people. They may be able to;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand familiar words with visual support")
                    pdf.drawString(36.85,b1-b2*3,"Öğrenci, nesneler, yerler ve insanları tanımlamak için kullanılan bazı temel kelimeleri tanımaya başlayabilir.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*4,"- Görsel destek yardımıyla temel kelimeleri anlama")
                    pdf.drawString(36.85,b1-b3-b2*4,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*4,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*5,"- Learn everyday words for objects and people in familiar categories such as home, school, family, colors, body parts, and animals")
                    pdf.drawString(36.85,b1-b3-b4-b2*6,"- Use pictures to help learn new words")
                    pdf.drawString(36.85,b1-b3-b4-b2*7,"- Listen to short, simple sentences about everyday actions, objects, and people. (example: She is swimming.)")
                    pdf.drawString(36.85,b1-b3-b4-b2*8,"- Practice using common, everyday expressions, such as greetings")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Ev, okul, aile, renkler, vücudun bölümleri ve hayvanlar gibi temel kategorilerdeki nesneler ve insanları tanımlamak için yaygın olarak kullanılan")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"kelimeleri ögrenebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Yeni kelimeler öğrenmesine yardımcı olması için resimleri kullanabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Günlük hayattaki eylemler, nesneler ve kisiler hakkında kısa, basit cümleler dinleyebilir.(örnek: Yüzüyor.)")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Selamlaşma ifadeleri gibi günlük hayatta yaygın olarak kullanılan ifadeleri kullanarak pratik yapabilir.")
                   
                #Belge özellikleri
                if str(totalscore)!=str("NS"):
                    fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_TurkceKarne.PDF")
                    documentTitle = 'Document title!'
                    pdf = canvas.Canvas(fileName,pagesize=(595.28,841.89))
                    pdf.setTitle(documentTitle)
                    #Font
                    pdfmetrics.registerFont(TTFont('abc', './data/tahoma.ttf'))
                    
                    #Ana baskı alanı
                    def baskialani(pdf):
                        #Arkaplan
                        pdf.drawImage(primarytk, 0,0, width=595,height=842,mask=None)
                        #Çizgiler
                        pdf.setLineWidth(0.665)
                        pdf.line(32.75, 754.02, 562.98, 754.02)
                        pdf.line(32.75, 31.19, 562.98, 31.19)
                        pdf.line(32.75, 119.06, 562.98, 119.06)
                        pdf.line(32.75, 436.54, 562.98, 436.54)
                        pdf.line(32.6, 754.35, 32.6, 30.87)
                        pdf.line(562.98, 754.35, 562.98, 30.87)
                        #Öğrenci Bilgisi
                        pdf.setFont('abc', 8)
                        pdf.drawString(123.31,804.06,"Student Name:")
                        pdf.drawString(123.31,789.88,"Test Date:")
                        pdf.drawString(123.31,775.71,"Class:")
                        pdf.drawString(272.13,804.06,"Student Number:")
                        pdf.drawString(272.13,789.88,"School Name:")
                        pdf.drawString(272.13,775.71,"Test Name:")        
                        pdf.drawString(180,804.06,str(studentname))
                        pdf.drawString(180,789.88,str(testdate))
                        pdf.drawString(180,775.71,str(studentclass))
                        pdf.drawString(334.49,804.06,str(studentnumber))
                        pdf.drawString(334.49,789.88,str(school))
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 1")
                        pdf.drawString(257.95,761.54,"Total Score : "+str(totalscore))
                        #Bazı içerikler
                        pdf.setFont('abc', 10)
                        pdf.drawString(45.35,741.09,"READING - OKUMA")
                        pdf.drawString(45.35,423.61,"LISTENING - DİNLEME")
                        pdf.drawString(45.35,106.13,"TOEFL History - TOEFL Geçmişi")
                        pdf.setFont('abc', 8)
                        pdf.drawString(160,741.69,"Score: "+str(rscore)+",")
                        pdf.drawString(210,741.69,"CEFR: "+str(rcefr)+",")
                        pdf.drawString(260,741.69,"Lexile: "+str(lexile))
                        pdf.drawString(527.25,741.69,"Stars: "+str(rstar[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore2)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(527.25,424.21,"Stars: "+str(lstar[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(rscore)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore2)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(lscore2)+" puan almıştır.)")
                        

                    #Anabaskı kodu
                    baskialani(pdf)
                    step1tk_1starlistening(pdf)


                    pdf.save()
            
            #Türkçe karne readingi ns, listeningi ns olmayan sonuç
            if rscore == "NS" and lscore != "NS":
                rscore2 = "100"
                totalscore = str(math.ceil((int(rscore2)+int(lscore))/2))

                #Reading yıldıza göre bilgiler
                def step1tk_1starreading(pdf):
                    pdf.drawString(36.85,702.01,"Students begin to recognize some basic words. They may be able to;")
                    pdf.drawString(36.85,690.67,"- Identify basic vocabulary with visual support")
                    pdf.drawString(36.85,667.99,"Öğrenci bazı temel kelimeleri tanımaya başlayabilir.Okuma becerileri:")
                    pdf.drawString(36.85,656.66,"- Görsel destekle temel kelimeleri tanıma")
                    pdf.drawString(36.85,639.65,"Next Steps")
                    pdf.drawString(36.85,625.47,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,614.14,"- Learn and practice reading common words in familiar categories such as home, school, family, colors, body parts, animals, and actions")
                    pdf.drawString(36.85,602.8,"- Read short, simple sentences about familiar people, objects, and action")
                    pdf.drawString(36.85,591.46,"Okuma becerilerini geliştirmek için öğrenci:")
                    #Eksik Var!!!!!!!!!
                    pdf.drawString(36.85,557.44,"- Tanıdık insanlar, nesneler ve eylemler hakkında kısa, basit cümleler okuyabilir.(örnek: Çocuk elma yiyor.)")
                   
                #Belge özellikleri
                if str(totalscore)!=str("NS"):
                    fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_TurkceKarne.PDF")
                    documentTitle = 'Document title!'
                    pdf = canvas.Canvas(fileName,pagesize=(595.28,841.89))
                    pdf.setTitle(documentTitle)
                    #Font
                    pdfmetrics.registerFont(TTFont('abc', './data/tahoma.ttf'))
                    
                    #Ana baskı alanı
                    def baskialani(pdf):
                        #Arkaplan
                        pdf.drawImage(primarytk, 0,0, width=595,height=842,mask=None)
                        #Çizgiler
                        pdf.setLineWidth(0.665)
                        pdf.line(32.75, 754.02, 562.98, 754.02)
                        pdf.line(32.75, 31.19, 562.98, 31.19)
                        pdf.line(32.75, 119.06, 562.98, 119.06)
                        pdf.line(32.75, 436.54, 562.98, 436.54)
                        pdf.line(32.6, 754.35, 32.6, 30.87)
                        pdf.line(562.98, 754.35, 562.98, 30.87)
                        #Öğrenci Bilgisi
                        pdf.setFont('abc', 8)
                        pdf.drawString(123.31,804.06,"Student Name:")
                        pdf.drawString(123.31,789.88,"Test Date:")
                        pdf.drawString(123.31,775.71,"Class:")
                        pdf.drawString(272.13,804.06,"Student Number:")
                        pdf.drawString(272.13,789.88,"School Name:")
                        pdf.drawString(272.13,775.71,"Test Name:")        
                        pdf.drawString(180,804.06,str(studentname))
                        pdf.drawString(180,789.88,str(testdate))
                        pdf.drawString(180,775.71,str(studentclass))
                        pdf.drawString(334.49,804.06,str(studentnumber))
                        pdf.drawString(334.49,789.88,str(school))
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 1")
                        pdf.drawString(257.95,761.54,"Total Score : "+str(totalscore))
                        #Bazı içerikler
                        pdf.setFont('abc', 10)
                        pdf.drawString(45.35,741.09,"READING - OKUMA")
                        pdf.drawString(45.35,423.61,"LISTENING - DİNLEME")
                        pdf.drawString(45.35,106.13,"TOEFL History - TOEFL Geçmişi")
                        pdf.setFont('abc', 8)
                        pdf.drawString(160,741.69,"Score: "+str(rscore2)+",")
                        pdf.drawString(210,741.69,"CEFR: "+str(rcefr)+",")
                        pdf.drawString(260,741.69,"Lexile: "+str(lexile))
                        pdf.drawString(527.25,741.69,"Stars: "+str(rstar[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(527.25,424.21,"Stars: "+str(lstar[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore2)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(rscore2)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore)+" on a scale of 100 to 109 (Öğrenci 100 ile 109 arasındaki ölçekte "+str(lscore)+" puan almıştır.)")
                        
                    #Anabaskı kodu
                    baskialani(pdf)
                    step1tk_1starreading(pdf)


                    pdf.save()

        step1_scorereport()
        step1_certificate()
        step1_tk()


def step1_classicbutton():
    buttons()
    tr1 = threading.Thread(target=step1_classic)
    tr1.start()
