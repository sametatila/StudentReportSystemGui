import os,time,threading,math
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

def step2_classic():
    global p1,window1,toplamsatir,satir,filename,f123
    import pdfplumber, re
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
    while satir<toplamsatir+1:
        satir += 1
        #Progressbar
        p1["value"] = satir+1
        p1["maximum"] = toplamsatir
        window1.update()
        time.sleep(0.00001)
        percent.set(str(int((((satir-2)*3)/((toplamsatir-1)*3))*100))+"%")
        text1.set(str((satir-2)*3)+"/"+str((toplamsatir-1)*3)+" belge tamamlandı.")

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
        sinavturu = 'Step 2'
        school = sheet['A'+str(satir)].value
        studentname = str(sheet['E'+str(satir)].value)+str(" ")+str(sheet['D'+str(satir)].value)
        studentnumber = sheet['F'+str(satir)].value
        studentclass = sheet['C'+str(satir)].value
        testdate = td1[8]+td1[9]+" "+td2+" "+td1[0]+td1[1]+td1[2]+td1[3]
        dateofbirth = td3[8]+td3[9]+" "+td4+" "+td3[0]+td3[1]+td3[2]+td3[3]
        rcefr = sheet['J'+str(satir)].value
        rscore = str(sheet['I'+str(satir)].value)
        lcefr = sheet['N'+str(satir)].value
        lscore = str(sheet['M'+str(satir)].value)
        lexile = sheet['K'+str(satir)].value
        
        rbadges = str(sheet['L'+str(satir)].value)
        lbadges = str(sheet['O'+str(satir)].value)

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
        step2_nsimage = './data/step2_ns.jpg'
        step2_1badgeimage = './data/step2_1badge.jpg'
        step2_2badgeimage = './data/step2_2badge.jpg'
        step2_3badgeimage = './data/step2_3badge.jpg'
        step2_4badgeimage = './data/step2_4badge.jpg'
        step2_5badgeimage = './data/step2_5badge.jpg'

        #Yızdız resimleri seritifika
        step2c_nsimage = './data/step2c_ns.jpg'
        step2c_1badgeimage = './data/step2c_1badge.jpg'
        step2c_2badgeimage = './data/step2c_2badge.jpg'
        step2c_3badgeimage = './data/step2c_3badge.jpg'
        step2c_4badgeimage = './data/step2c_4badge.jpg'
        step2c_5badgeimage = './data/step2c_5badge.jpg'
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

        def step2_scorereport():
            #Reading yıldıza göre bilgiler

            def step2_nsreading(pdf):
                #NS Image
                pdf.drawImage(step2_nsimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.17, 279.44, 567.17)
                pdf.line(61.22, 500.48, 279.44, 500.48)
                pdf.line(61.22, 451.03, 279.44, 451.03)
                #Noktalar Yok
                #Part1
                pdf.setFont('abc', 8)
                pdf.drawString(60.9,529.65,"The test taker did not respond to any questions in this")
                pdf.drawString(60.9,520.45,"section. Therefore, the scores for this section cannot be")
                pdf.drawString(60.9,511.25,"provided.")
                #Part2 Boş
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,432.12,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,421.77,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,411.42,"The student received "+str(rscore)+" on a scale of 100 to 115")

            def step2_1badgereading(pdf):
                #1badge Image
                pdf.drawImage(step2_1badgeimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.16, 279.36, 567.16)
                pdf.line(61.22, 498.78, 279.36, 498.78)
                pdf.line(61.22, 336.79, 279.36, 336.79)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(77.965,516.95,nokta)
                pdf.drawString(77.965,457.75,nokta)
                pdf.drawString(77.965,438.45,nokta)
                pdf.drawString(77.965,410.45,nokta)
                pdf.drawString(77.965,391.18,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,549.41,"Students begin to recognize some basic words. They")
                pdf.drawString(60.9,539.06,"may be able to:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,518.76,"Identify basic vocabulary with visual support")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,479.88,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,459.58,"Learn words and common expressions used in")
                pdf.drawString(96.9,450.38,"familiar social settings")
                pdf.drawString(96.9,440.64,"Learn words that show relationships among people,")
                pdf.drawString(96.9,431.44,"objects, and places (examples:")
                pdf.setFont('abcit', 8)
                pdf.drawString(208.96,431.44,"at, on, around,")
                pdf.drawString(96.9,422.24,"between, on top of")
                pdf.setFont('abc', 8)
                pdf.drawString(162.73,422.24,")")
                pdf.drawString(96.9,412.51,"Practice reading simple sentences and short texts")
                pdf.drawString(96.9,403.31,"about familiar topics")
                pdf.drawString(96.9,393.57,"Consider taking the TOEFL Primary Step 1 test for")
                pdf.drawString(96.9,384.37,"more information about their reading ability")
                pdf.drawString(60.9,375.17,"Note: Lexile information provided for students at this score")
                pdf.drawString(60.9,365.97,"level is less precise than at other score levels. Students")
                pdf.drawString(60.9,356.77,"should consider taking the TOEFL Primary Step 1 test for")
                pdf.drawString(60.9,347.57,"more precise information about their Lexile level.")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,317.89,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,307.54,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,297.19,"The student received "+str(rscore)+" on a scale of 100 to 115")

            def step2_2badgereading(pdf):
                #2badge Image
                pdf.drawImage(step2_2badgeimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.16, 279.36, 567.16)
                pdf.line(61.22, 386.78, 279.36, 386.78)
                pdf.line(61.22, 280.52, 279.36, 280.52)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(77.965,515.95,nokta)
                pdf.drawString(77.965,486.75,nokta)
                pdf.drawString(77.965,459.45,nokta)
                pdf.drawString(77.965,421.45,nokta)
                pdf.drawString(77.965,344.18,nokta)
                pdf.drawString(77.965,325.24,nokta)
                pdf.drawString(77.965,306.24,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,549.41,"Students understand short descriptions and find")
                pdf.drawString(60.9,539.06,"information in signs, messages, and stories. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,518.76,"Understand common words and social expressions")
                pdf.drawString(96.9,509.56,"(")
                pdf.setFont('abcit', 8)
                pdf.drawString(99.56,509.56,"examples: play a game, go to a museum, wave")
                pdf.drawString(96.9,500.36,"goodbye")
                pdf.setFont('abc', 8)
                pdf.drawString(127.6,500.36,")")
                pdf.drawString(96.9,490.63,"Comprehend simple descriptions of current and")
                pdf.drawString(96.9,481.43,"past events (")
                pdf.setFont('abcit', 8)
                pdf.drawString(142.7,481.43,"examples: The mouse is on top of the")
                pdf.drawString(96.9,472.37,"table. He is washing his hands.")
                pdf.setFont('abc', 8)
                pdf.drawString(207.19,472.37,")")
                pdf.drawString(96.9,462.49,"Recognize relationships among words and phrases")
                pdf.drawString(96.9,453.29,"within familiar categories (")
                pdf.setFont('abcit', 8)
                pdf.drawString(189.38,453.29,"examples: food–fruit–")
                pdf.drawString(96.9,444.09,"strawberries; rain–sky–clouds; one more time–")
                pdf.drawString(96.9,434.89,"again")
                pdf.setFont('abc', 8)
                pdf.drawString(116.48,434.89,")")
                pdf.drawString(96.9,425.15,"Make connections across simple sentences")
                pdf.drawString(96.9,415.96,"(")
                pdf.setFont('abcit', 8)
                pdf.drawString(99.56,415.96,"example: Clouds are in the sky. Rain comes from")
                pdf.drawString(96.9,406.76,"them. Sometimes they cover the sun.")
                pdf.setFont('abc', 8)
                pdf.drawString(228.96,406.76,")")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,367.87,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,347.57,"Read longer paragraphs and stories about familiar")
                pdf.drawString(96.9,338.37,"people, objects, and information")
                pdf.drawString(96.9,328.64,"Learn more words that describe objects, places,")
                pdf.drawString(96.9,319.44,"people, actions, and ideas")
                pdf.drawString(96.9,309.7,"Speak or write in their own words about")
                pdf.drawString(96.9,300.5,"paragraphs, stories, and information they read")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,261.62,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,251.27,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,240.92,"The student received "+str(rscore)+" on a scale of 100 to 115")

            def step2_3badgereading(pdf):
                #3badge Image
                pdf.drawImage(step2_3badgeimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.16, 279.36, 567.16)
                pdf.line(61.22, 376.43, 279.36, 376.43)
                pdf.line(61.22, 269.63, 279.36, 269.63)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(77.965,504.95,nokta)
                pdf.drawString(77.965,467.75,nokta)
                pdf.drawString(77.965,430.45,nokta)
                pdf.drawString(77.965,392.45,nokta)
                pdf.drawString(77.965,334.18,nokta)
                pdf.drawString(77.965,324.24,nokta)
                pdf.drawString(77.965,305.24,nokta)
                pdf.drawString(77.965,295.24,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,549.41,"Students understand simple stories and are beginning")
                pdf.drawString(60.9,539.06,"to understand age-appropriate academic texts. They")
                pdf.drawString(60.9,528.71,"can:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,508.41,"Understand common words and some less")
                pdf.drawString(96.9,499.21,"common words about objects, places, people,")
                pdf.drawString(96.9,490.02,"actions, and ideas (")
                pdf.setFont('abcit', 8)
                pdf.drawString(166.28,490.02,"examples: ring, adventures,")
                pdf.drawString(96.9,480.82,"whisper, double")
                pdf.setFont('abc', 8)
                pdf.drawString(152.93,480.82,")")
                pdf.drawString(96.9,471.08,"Comprehend the meaning of complex sentences")
                pdf.drawString(96.9,461.88,"(")
                pdf.setFont('abcit', 8)
                pdf.drawString(99.56,461.88,"examples: This is a friendly thing to do when you")
                pdf.drawString(96.9,452.68,"say goodbye. People do this when they talk")
                pdf.drawString(96.9,443.48,"quietly.")
                pdf.setFont('abc', 8)
                pdf.drawString(122.25,443.48,")")
                pdf.drawString(96.9,433.74,"Connect information in longer sentences and")
                pdf.drawString(96.9,424.54,"across different sentences to infer information,")
                pdf.drawString(96.9,415.35,"identify main ideas, and understand the meaning of")
                pdf.drawString(96.9,406.15,"unfamiliar words.")
                pdf.drawString(96.9,396.41,"Locate key information in texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,357.52,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,337.23,"Study new, unfamiliar words")
                pdf.drawString(96.9,327.49,"Practice reading stories and informational texts")
                pdf.drawString(96.9,318.29,"about a variety of topics")
                pdf.drawString(96.9,308.55,"Practice reading longer and more complex texts")
                pdf.drawString(96.9,298.81,"Speak or write in their own words about stories and")
                pdf.drawString(96.9,289.61,"information they read")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,250.73,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,240.38,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,230.03,"The student received "+str(rscore)+" on a scale of 100 to 115")

            def step2_4badgereading(pdf):
                #4badge Image
                pdf.drawImage(step2_4badgeimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.16, 279.35, 567.16)
                pdf.line(61.22, 423.47, 279.35, 423.47)
                pdf.line(61.22, 336.25, 279.35, 336.25)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(77.965,515.95,nokta)
                pdf.drawString(77.965,486.75,nokta)
                pdf.drawString(77.965,468.45,nokta)
                pdf.drawString(77.965,439.45,nokta)
                pdf.drawString(77.965,381.18,nokta)
                pdf.drawString(77.965,361.24,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,549.41,"Students understand simple stories and age")
                pdf.drawString(60.9,539.06,"appropriate academic texts. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,518.76,"Understand a variety of common words and many")
                pdf.drawString(96.9,509.56,"less common words about objects, places, people,")
                pdf.drawString(96.9,500.36,"actions, and ideas")
                pdf.drawString(96.9,490.63,"Comprehend the meanings of complex sentences")
                pdf.drawString(96.9,481.43,"and paragraphs")
                pdf.drawString(96.9,471.69,"Connect information in longer sentences and")
                pdf.drawString(96.9,462.49,"across several sentences to infer information, main")
                pdf.drawString(96.9,453.29,"ideas, and the meaning of unfamiliar words")
                pdf.drawString(96.9,443.55,"Identify specific details in texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,404.67,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,384.37,"Read longer and more complex stories and")
                pdf.drawString(96.9,375.17,"informational texts about a variety of topics")
                pdf.drawString(96.9,365.43,"Speak or write in their own words about stories and")
                pdf.drawString(96.9,356.23,"information they read")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,317.35,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,307,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,296.65,"The student received "+str(rscore)+" on a scale of 100 to 115")
                
            def step2_5badgereading(pdf):
                #4badge Image
                pdf.drawImage(step2_5badgeimage, 89.3,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(61.22, 567.16, 279.35, 567.16)
                pdf.line(61.22, 423.57, 279.35, 423.57)
                pdf.line(61.22, 317.32, 279.35, 317.32)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(77.965,515.95,nokta)
                pdf.drawString(77.965,487.75,nokta)
                pdf.drawString(77.965,468.45,nokta)
                pdf.drawString(77.965,439.45,nokta)
                pdf.drawString(77.965,380.18,nokta)
                pdf.drawString(77.965,361.24,nokta)
                pdf.drawString(77.965,343.31,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,549.41,"Students perform exceptionally well on this test. They")
                pdf.drawString(60.9,539.06,"can:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,518.76,"Understand a wide variety of common and less")
                pdf.drawString(96.9,509.56,"common words to describe objects, places, people,")
                pdf.drawString(96.9,500.36,"actions, and ideas")
                pdf.drawString(96.9,490.63,"Comprehend the meaning of complex sentences,")
                pdf.drawString(96.9,481.43,"paragraphs, and longer texts")
                pdf.drawString(96.9,471.69,"Connect information across several sentences and")
                pdf.drawString(96.9,462.49,"paragraphs to infer information, identify main ideas,")
                pdf.drawString(96.9,453.29,"and understand the meaning of unfamiliar words")
                pdf.drawString(96.9,443.55,"Identify specific details in longer texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(60.9,404.67,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(96.9,384.37,"Read longer and more complex stories and")
                pdf.drawString(96.9,375.17,"academic texts about a variety of topics")
                pdf.drawString(96.9,364.43,"Speak or write in their own words about stories and")
                pdf.drawString(96.9,356.23,"information they read")
                pdf.drawString(96.9,346.50,"Consider taking the TOEFL Junior® test for more")
                pdf.drawString(96.9,337.30,"accurate information about their reading ability")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(60.9,298.41,"CEFR Level: "+str(rcefr))
                pdf.drawString(60.9,288.06,"Lexile Measure: "+str(lexile))
                pdf.drawString(60.9,277.71,"The student received "+str(rscore)+" on a scale of 100 to 115")

            #Listening yıldıza göre bilgiler

            def step2_nslistening(pdf):
                #NS Image
                pdf.drawImage(step2_nsimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.2, 567.16)
                pdf.line(325.57, 500.47, 543.2, 500.47)
                pdf.line(325.57, 452.18, 543.2, 452.18)
                #Noktalar Yok
                #Part1
                pdf.setFont('abc', 8)
                pdf.drawString(325.3,529.65,"The test taker did not respond to any questions in this")
                pdf.drawString(325.3,520.45,"section. Therefore, the scores for this section cannot be")
                pdf.drawString(325.3,511.25,"provided.")
                #Part2 Boş
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,433.27,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,412.57,"The student received "+str(lscore)+" on a scale of 100 to 115")

            def step2_1badgelistening(pdf):
                #1badge Image
                pdf.drawImage(step2_1badgeimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.2, 567.16)
                pdf.line(325.57, 488.43, 543.2, 488.43)
                pdf.line(325.57, 345.45, 543.2, 345.45)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(342.29,505.95,nokta)
                pdf.drawString(342.29,447.45,nokta)
                pdf.drawString(342.29,419.78,nokta)
                pdf.drawString(342.29,410.55,nokta)
                pdf.drawString(342.29,391.84,nokta)
                pdf.drawString(342.29,371.84,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.41,"Students begin to recognize a few familiar words in")
                pdf.drawString(325.3,539.06,"speech, such as words for objects, places, and people.")
                pdf.drawString(325.3,528.71,"They may be able to:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,508.41,"Understand familiar words with visual support")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,470.68,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,450.38,"Learn everyday words for objects and people in")
                pdf.drawString(361.3,441.18,"familiar categories such as home, school, family,")
                pdf.drawString(361.3,431.98,"colors, body parts, and animals")
                pdf.drawString(361.3,422.24,"Practice having short, simple conversations")
                pdf.drawString(361.3,412.51,"Practice listening to teacher instructions and short")
                pdf.drawString(361.3,403.31,"messages")
                pdf.drawString(361.3,393.57,"Begin listening to and identifying information in")
                pdf.drawString(361.3,384.37,"short, simple stories")
                pdf.drawString(361.3,374.63,"Consider taking the TOEFL Primary Step 1 test for")
                pdf.drawString(361.3,365.43,"more information about their listening ability")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,326.55,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,305.85,"The student received "+str(lscore)+" on a scale of 100 to 115")

            def step2_2badgelistening(pdf):
                #2badge Image
                pdf.drawImage(step2_2badgeimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.2, 567.16)
                pdf.line(325.57, 413.22, 543.2, 413.22)
                pdf.line(325.57, 289.18, 543.2, 289.18)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(342.29,504.95,nokta)
                pdf.drawString(342.29,485.55,nokta)
                pdf.drawString(342.29,467.45,nokta)
                pdf.drawString(342.29,448.45,nokta)
                pdf.drawString(342.29,371.78,nokta)
                pdf.drawString(342.29,352.55,nokta)
                pdf.drawString(342.29,333.84,nokta)
                pdf.drawString(342.29,314.84,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.41,"Students understand basic conversations and")
                pdf.drawString(325.3,539.06,"messages and begin to understand stories and")
                pdf.drawString(325.3,528.71,"informational texts. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,508.41,"Understand common expressions used in everyday")
                pdf.drawString(361.3,499.21,"conversations")
                pdf.drawString(361.3,489.48,"Understand a simple, single instruction spoken in")  
                pdf.drawString(361.3,480.28,"familiar words, with key words repeated")
                pdf.drawString(361.3,470.54,"Understand the purpose of messages in which key")
                pdf.drawString(361.3,461.34,"information is repeated")
                pdf.drawString(361.3,451.6,"Understand the main ideas of simple stories in")
                pdf.drawString(361.3,442.4,"which key information is explicitly stated and")
                pdf.drawString(361.3,433.2,"repeated")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,395.47,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,375.17,"Study more words that describe familiar topics,")
                pdf.drawString(361.3,365.97,"settings, and actions")
                pdf.drawString(361.3,356.23,"Practice using less common words and")
                pdf.drawString(361.3,347.03,"expressions in conversations")
                pdf.drawString(361.3,337.3,"Listen to age-appropriate academic talks and")
                pdf.drawString(361.3,328.1,"longer stories")
                pdf.drawString(361.3,318.36,"Speak or write in their own words about stories and")
                pdf.drawString(361.3,309.16,"information they listen to")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,270.28,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,249.58,"The student received "+str(lscore)+" on a scale of 100 to 115")

            def step2_3badgelistening(pdf):
                #3badge Image
                pdf.drawImage(step2_3badgeimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.2, 567.16)
                pdf.line(325.57, 394.29, 543.2, 394.29)
                pdf.line(325.57, 289.18, 543.2, 289.18)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(342.29,504.95,nokta)
                pdf.drawString(342.29,476.55,nokta)
                pdf.drawString(342.29,458.45,nokta)
                pdf.drawString(342.29,438.45,nokta)
                pdf.drawString(342.29,420.78,nokta)
                pdf.drawString(342.29,352.55,nokta)
                pdf.drawString(342.29,314.84,nokta)
                pdf.drawString(342.29,333.84,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.41,"Students understand conversations and simple stories.")
                pdf.drawString(325.3,539.06,"They begin to understand age-appropriate academic")
                pdf.drawString(325.3,528.71,"talks. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,508.41,"Understand less common words that describe")
                pdf.drawString(361.3,499.21,"familiar topics, settings, and actions (")
                pdf.setFont('abcit', 8)
                pdf.drawString(492.47,499.21,"examples:")
                pdf.drawString(361.3,490.02,"pocket, pour, lamp, branch")
                pdf.setFont('abc', 8)
                pdf.drawString(456.02,490.02,")")
                pdf.drawString(361.3,480.28,"Understand indirect responses to questions in")
                pdf.drawString(361.3,471.08,"conversations")
                pdf.drawString(361.3,461.34,"Understand messages in which information is not")
                pdf.drawString(361.3,452.14,"explicitly stated")
                pdf.drawString(361.3,442.4,"Connect information to infer the main idea or topic")
                pdf.drawString(361.3,433.2,"of messages, stories, and informational texts")
                pdf.drawString(361.3,423.47,"Synthesize information from multiple locations in a")
                pdf.drawString(361.3,414.27,"longer spoken text")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,376.53,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,356.23,"Learn new, unfamiliar words they hear in longer")
                pdf.drawString(361.3,347.03,"stories and academic talks")
                pdf.drawString(361.3,337.3,"Practice using less common words and")
                pdf.drawString(361.3,328.1,"expressions in conversations")
                pdf.drawString(361.3,318.36,"Speak or write in their own words about stories and")
                pdf.drawString(361.3,309.16,"information they listen to")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,270.28,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,249.58,"The student received "+str(lscore)+" on a scale of 100 to 115")

            def step2_4badgelistening(pdf):
                #4badge Image
                pdf.drawImage(step2_4badgeimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.56, 567.16)
                pdf.line(325.57, 441.97, 543.56, 441.97)
                pdf.line(325.57, 336.86, 543.56, 336.86)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(342.29,515.95,nokta)
                pdf.drawString(342.29,496.55,nokta)
                pdf.drawString(342.29,477.45,nokta)
                pdf.drawString(342.29,458.45,nokta)
                pdf.drawString(342.29,400.78,nokta)
                pdf.drawString(342.29,381.55,nokta)
                pdf.drawString(342.29,362.84,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.41,"Students understand conversations, simple stories,")
                pdf.drawString(325.3,539.06,"and age-appropriate academic talks. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,518.76,"Understand less frequently used words that")
                pdf.drawString(361.3,509.56,"describe familiar topics, settings, and actions")
                pdf.drawString(361.3,499.82,"Understand messages and stories that include")  
                pdf.drawString(361.3,490.63,"unfamiliar words and some idiomatic expressions")
                pdf.drawString(361.3,480.89,"Consistently connect information throughout stories")
                pdf.drawString(361.3,471.69,"and academic talks to infer meaning")
                pdf.drawString(361.3,461.95,"Identify specific information in longer texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,424.22,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,403.92,"Listen to longer and more complex stories and")
                pdf.drawString(361.3,394.72,"academic texts about a variety of topics")
                pdf.drawString(361.3,384.98,"Practice using less common words and")
                pdf.drawString(361.3,375.78,"expressions in conversations")
                pdf.drawString(361.3,366.04,"Speak or write in their own words about stories and")
                pdf.drawString(361.3,356.85,"information they listen to")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,317.96,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,297.26,"The student received "+str(lscore)+" on a scale of 100 to 115")

            def step2_5badgelistening(pdf):
                #4badge Image
                pdf.drawImage(step2_5badgeimage, 353.65,576.25, width=162,height=63,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(325.57, 567.16, 543.56, 567.16)
                pdf.line(325.57, 441.97, 543.56, 441.97)
                pdf.line(325.57, 317.93, 543.56, 317.93)
                #Noktalar
                pdf.setFont('dot', 14.5)
                pdf.drawString(342.29,515.95,nokta)
                pdf.drawString(342.29,496.55,nokta)
                pdf.drawString(342.29,477.45,nokta)
                pdf.drawString(342.29,458.45,nokta)
                pdf.drawString(342.29,399.78,nokta)
                pdf.drawString(342.29,381.55,nokta)
                pdf.drawString(342.29,362.84,nokta)
                pdf.drawString(342.29,342.91,nokta)
                #Part1
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,549.41,"Students perform exceptionally well on this test. They")
                pdf.drawString(325.3,539.06,"can:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,518.76,"Understand less frequently used words that")
                pdf.drawString(361.3,509.56,"describe familiar topics, settings, and actions")
                pdf.drawString(361.3,499.82,"Understand messages and stories that include")  
                pdf.drawString(361.3,490.63,"unfamiliar words and some idiomatic expressions")
                pdf.drawString(361.3,480.89,"Consistently connect information throughout stories")
                pdf.drawString(361.3,471.69,"and academic talks to infer meaning")
                pdf.drawString(361.3,461.95,"Identify specific information in longer texts")
                #Part2
                pdf.setFont('abc', 9)
                pdf.drawString(325.3,424.22,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(361.3,403.92,"Listen to longer and more complex stories and")
                pdf.drawString(361.3,394.72,"academic texts about a variety of topics")
                pdf.drawString(361.3,384.98,"Practice using less common words and")
                pdf.drawString(361.3,375.78,"expressions in conversations")
                pdf.drawString(361.3,366.04,"Speak or write in their own words about stories and")
                pdf.drawString(361.3,356.85,"information they listen to")
                pdf.drawString(361.3,347.11,"Consider taking the TOEFL Junior® test for more")
                pdf.drawString(361.3,337.91,"accurate information about their listening ability")
                #Part3
                pdf.setFont('abcbold', 9)
                pdf.drawString(325.3,299.02,"CEFR Level: "+str(lcefr))
                pdf.drawString(325.3,278.33,"The student received "+str(lscore)+" on a scale of 100 to 115")

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
                pdf.drawString(479.99,767.04 ,str(sinavturu))
                #Öğrenci Bilgisi
                pdf.setFont('abc', 9.5)
                pdf.drawString(60.7,739.26,"Student Name:  "+str(studentname))
                pdf.drawString(61.38,723.04,"Student Number:  "+str(studentnumber))
                pdf.drawString(61.38,706.65,"Test Date:  "+str(testdate))
                pdf.drawString(423.15,739.25,"Date of Birth:  "+str(dateofbirth))
                pdf.drawString(423.15,722.86,"Gender:  "+str(lgender[4]))
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
                step2_nsreading(pdf)
            else:
                if rscore == str("100"):
                    step2_1badgereading(pdf)
                else:
                    if rscore ==  str("104"):
                        step2_2badgereading(pdf)
                    else:
                        if rscore == str("105"):
                            step2_2badgereading(pdf)
                        else:
                            if rscore == str("106"):
                                step2_2badgereading(pdf)
                            else:
                                if rscore == str("107"):
                                    step2_3badgereading(pdf)
                                else:
                                    if rscore ==  str("108"):
                                        step2_3badgereading(pdf)
                                    else:
                                        if rscore == str("109"):
                                            step2_3badgereading(pdf)
                                        else:
                                            if rscore == str("110"):
                                                step2_4badgereading(pdf)
                                            else:
                                                if rscore == str("111"):
                                                    step2_4badgereading(pdf)
                                                else:
                                                    if rscore == str("112"):
                                                        step2_4badgereading(pdf)
                                                    else:
                                                        if rscore == str("113"):
                                                            step2_5badgereading(pdf)
                                                        else:
                                                            if rscore == str("114"):
                                                                step2_5badgereading(pdf)
                                                            else:
                                                                if rscore == str("115"):
                                                                    step2_5badgereading(pdf)

            #Listening skora göre sonuç seçimi
            if lscore == "NS":
                step2_nslistening(pdf)
            else:
                if lscore == str("100"):
                    step2_1badgelistening(pdf)
                else:
                    if lscore ==  str("104"):
                        step2_2badgelistening(pdf)
                    else:
                        if lscore == str("105"):
                            step2_2badgelistening(pdf)
                        else:
                            if lscore == str("106"):
                                step2_2badgelistening(pdf)
                            else:
                                if lscore == str("107"):
                                    step2_3badgelistening(pdf)
                                else:
                                    if lscore ==  str("108"):
                                        step2_3badgelistening(pdf)
                                    else:
                                        if lscore == str("109"):
                                            step2_3badgelistening(pdf)
                                        else:
                                            if lscore == str("110"):
                                                step2_4badgelistening(pdf)
                                            else:
                                                if lscore == str("111"):
                                                    step2_4badgelistening(pdf)
                                                else:
                                                    if lscore == str("112"):
                                                        step2_4badgelistening(pdf)
                                                    else:
                                                        if lscore == str("113"):
                                                            step2_5badgelistening(pdf)
                                                        else:
                                                            if lscore == str("114"):
                                                                step2_5badgelistening(pdf)
                                                            else:
                                                                if lscore == str("115"):
                                                                    step2_5badgelistening(pdf)

            pdf.save()

        def step2_certificate():
            #Reading sonuçları
            def step2c_nsreading(pdf):
                pdf.drawImage(step2c_nsimage, 418.6295,225.6, width=19.3,height=18.2,mask=None)

            def step2c_1badgereading(pdf):
                pdf.drawImage(step2c_1badgeimage, 418.3,225.95, width=92.8,height=18.45,mask=None)

            def step2c_2badgereading(pdf):
                pdf.drawImage(step2c_2badgeimage, 418.3,225.95, width=92.8,height=18.45,mask=None)

            def step2c_3badgereading(pdf):
                pdf.drawImage(step2c_3badgeimage, 418.3,225.95, width=92.8,height=18.45,mask=None)

            def step2c_4badgereading(pdf):
                pdf.drawImage(step2c_4badgeimage, 418.3,225.95, width=92.8,height=18.45,mask=None)
                
            def step2c_5badgereading(pdf):
                pdf.drawImage(step2c_5badgeimage, 418.3,225.95, width=92.8,height=18.45,mask=None)

            #Listening sonuçları
            def step2c_nslistening(pdf):
                pdf.drawImage(step2c_nsimage, 418.6295,184.8, width=19.3,height=18.2,mask=None)

            def step2c_1badgelistening(pdf):
                pdf.drawImage(step2c_1badgeimage, 418.9,185, width=92.8,height=18.45,mask=None)

            def step2c_2badgelistening(pdf):
                pdf.drawImage(step2c_2badgeimage, 418.9,185, width=92.8,height=18.45,mask=None)

            def step2c_3badgelistening(pdf):
                pdf.drawImage(step2c_3badgeimage, 418.9,185, width=92.8,height=18.45,mask=None)

            def step2c_4badgelistening(pdf):
                pdf.drawImage(step2c_4badgeimage, 418.9,185, width=92.8,height=18.45,mask=None)
                
            def step2c_5badgelistening(pdf):
                pdf.drawImage(step2c_5badgeimage, 418.9,185, width=92.8,height=18.45,mask=None)

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
                pdf.drawString(153.74,271.3,"Has earned the following levels on the TOEFL")
                pdf.drawString(532.5,271.3,"Primary")
                pdf.drawString(617.38,271.3,"Test")
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
                step2c_nsreading(pdf)
            else:
                if rscore == str("100"):
                    step2c_1badgereading(pdf)
                else:
                    if rscore ==  str("104"):
                        step2c_2badgereading(pdf)
                    else:
                        if rscore == str("105"):
                            step2c_2badgereading(pdf)
                        else:
                            if rscore == str("106"):
                                step2c_2badgereading(pdf)
                            else:
                                if rscore == str("107"):
                                    step2c_3badgereading(pdf)
                                else:
                                    if rscore ==  str("108"):
                                        step2c_3badgereading(pdf)
                                    else:
                                        if rscore == str("109"):
                                            step2c_3badgereading(pdf)
                                        else:
                                            if rscore == str("110"):
                                                step2c_4badgereading(pdf)
                                            else:
                                                if rscore == str("111"):
                                                    step2c_4badgereading(pdf)
                                                else:
                                                    if rscore == str("112"):
                                                        step2c_4badgereading(pdf)
                                                    else:
                                                        if rscore == str("113"):
                                                            step2c_5badgereading(pdf)
                                                        else:
                                                            if rscore == str("114"):
                                                                step2c_5badgereading(pdf)
                                                            else:
                                                                if rscore == str("115"):
                                                                    step2c_5badgereading(pdf)

            #Listening skora göre sonuç seçimi
            if lscore == "NS":
                step2c_nslistening(pdf)
            else:
                if lscore == str("100"):
                    step2c_1badgelistening(pdf)
                else:
                    if lscore ==  str("104"):
                        step2c_2badgelistening(pdf)
                    else:
                        if lscore == str("105"):
                            step2c_2badgelistening(pdf)
                        else:
                            if lscore == str("106"):
                                step2c_2badgelistening(pdf)
                            else:
                                if lscore == str("107"):
                                    step2c_3badgelistening(pdf)
                                else:
                                    if lscore ==  str("108"):
                                        step2c_3badgelistening(pdf)
                                    else:
                                        if lscore == str("109"):
                                            step2c_3badgelistening(pdf)
                                        else:
                                            if lscore == str("110"):
                                                step2c_4badgelistening(pdf)
                                            else:
                                                if lscore == str("111"):
                                                    step2c_4badgelistening(pdf)
                                                else:
                                                    if lscore == str("112"):
                                                        step2c_4badgelistening(pdf)
                                                    else:
                                                        if lscore == str("113"):
                                                            step2c_5badgelistening(pdf)
                                                        else:
                                                            if lscore == str("114"):
                                                                step2c_5badgelistening(pdf)
                                                            else:
                                                                if lscore == str("115"):
                                                                    step2c_5badgelistening(pdf)

            pdf.save()

        def step2_tk():
            if rscore != "NS" and lscore !="NS":
                totalscore = str(math.ceil((int(rscore)+int(lscore))/2))
                
                #Reading yıldıza göre bilgiler
                def step2tk_1badgesreading(pdf):
                    pdf.drawString(36.85,702.01,"Students begin to recognize some basic words. They may be able to;")
                    pdf.drawString(36.85,690.67,"- Identify basic vocabulary with visual support")
                    pdf.drawString(36.85,667.99,"Öğrenci, bazı temel kelimeleri tanımaya başlayabilir.Dinleme becerileri:")
                    pdf.drawString(36.85,656.66,"- Görsel destek yardımıyla temel kelimeleri anlama")
                    pdf.drawString(36.85,639.65,"Next Steps")
                    pdf.drawString(36.85,625.47,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,614.14,"- Learn words and common expressions used in familiar social settings")
                    pdf.drawString(36.85,602.80,"- Learn words that show relationships among people, objects,and places (examples: at, on)")
                    pdf.drawString(36.85,591.46,"- Practice reading simple sentences and short texts about familiar topics")
                    pdf.drawString(36.85,580.12,"- Consider taking the TOEFL Primary Step 1 test for more information about their reading ability")
                    pdf.drawString(36.85,557.44,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,546.10,"- Günlük hayatta kullanılan temel kelime ve ifadeleri öğrenebilir")
                    pdf.drawString(36.85,534.77,"- İnsanlar, nesneler ve yerler arasındaki ilişkileri tanımlayan ifadeleri öğrenebilir.(örnekler: at, on)")
                    pdf.drawString(36.85,523.43,"- Bilindik konular hakkında basit cümleler ve kısa metinlerle okuma alıştırmaları yapabilir.")
                    pdf.drawString(36.85,512.09,"- Okuma becerileri hakkında daha fazla bilgi edinmek için TOEFL Primary Step 1 sınavına girmeyi düşünebilir.")

                def step2tk_2badgesreading(pdf):
                    pdf.drawString(36.85,702.01,"Students understand short descriptions and find information in signs, messages, and stories. They can;")
                    pdf.drawString(36.85,690.67,"- Understand common words and social expressions (examples: play a game, go to a museum, wave goodbye)")
                    pdf.drawString(36.85,679.33,"- Comprehend simple descriptions of current and past events (examples: The mouse is on top of the table. He is washing his hands.)")
                    pdf.drawString(36.85,667.99,"- Recognize relationships among words and phrases within familiar categories (examples: food-fruit-strawberries; rain-sky-clouds;)")
                    pdf.drawString(36.85,656.66,"- Make connections across simple sentences (example: Clouds are in the sky. Rain comes from them. Sometimes they cover the sun.)")
                    pdf.drawString(36.85,633.98,"Öğrenci, kısa açıklamaları, tabela, levha, mesaj ve kısa hikayelerdeki bilgileri anlar.Okuma becerileri:")
                    pdf.drawString(36.85,622.64,"- Sık kullanılan kelimeleri ve sosyal yaşama dair ifadeleri anlama (örnekler: oyun oynama, müzeye gitme, el sallama)")
                    pdf.drawString(36.85,611.30,"- Güncel ve geçmişte yaşanmış olayların basit tanımlamalarını kavrama (örnekler: -Fare masanın üstünde. -Ellerini yıkıyor.)")
                    pdf.drawString(36.85,599.96,"- Aynı kategoriye ait kelimeler ve kelime öbekleri arasındaki ilişkileri anlama (örn:gıda,meyve,çilek; yağmur,gök,bulutlar,)")
                    pdf.drawString(36.85,588.62,"- Basit cümleler arasında bağlantı kurma.(örnek: Bulutlar gökyüzünde. Yağmur onlardan geliyor. Bazen güneşi kapatıyorlar.)")
                    pdf.drawString(36.85,571.62,"Next Steps")
                    pdf.drawString(36.85,557.44,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,546.10,"- Read longer paragraphs and stories about familiar people, objects, and information")
                    pdf.drawString(36.85,534.77,"- Learn more words that describe objects, places, people, actions, and ideas")
                    pdf.drawString(36.85,523.43,"- Speak or write in their own words about paragraphs, stories, and information they read")
                    pdf.drawString(36.85,500.75,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,489.41,"- Tanıdık insanlar, nesneler ve bilgiler hakkında daha uzun paragraflar ve hikayeler okuyabilir.")
                    pdf.drawString(36.85,478.07,"- Nesneleri, yerleri, insanları, eylemleri ve fikirleri tanımlamak için daha fazla kelime öğrenebilir.")
                    pdf.drawString(36.85,466.73,"- Okuduğu hikayeler, bilgiler ve paragraflar hakkında, kendi cümleleriyle,konuşarak veya yazarak kendisini ifade edebilir.")
                  
                def step2tk_3badgesreading(pdf):
                    
                    pdf.drawString(36.85,702.01,"Students understand simple stories and are beginning to understand age-appropriate academic texts. They can;")
                    pdf.drawString(36.85,690.67,"- Understand common words and some less common words about objects, places, people, actions, and ideas")
                    pdf.drawString(36.85,679.33,"- Comprehend the meaning of complex sentences (e.g This is a friendly thing to do when you say goodbye. People do this when they talk quietly.)")
                    pdf.drawString(36.85,667.99,"- Connect information in longer sentences and across different sentences to infer information, identify main ideas,")
                    pdf.drawString(36.85,656.66,"- Locate key information in texts and understand the meaning of unfamiliar words")
                    pdf.drawString(36.85,633.98,"Öğrenci basit hikayeleri anlar ve yaşına uygun akademik metinleri anlamaya başlar.Okuma becerileri:")
                    pdf.drawString(36.85,622.64,"- Nesneler, yerler, insanlar, eylemler, fikirler için yaygın olarak kullanılan kelimeleri ve yaygın olarak kullanılmayan bazı kelimeleri anlar.")
                    pdf.drawString(36.85,611.30,"- Karmaşık cümlelerin anlamını kavrar.(örn:Bu hoşçakal dediğinde yapılacak dostça bir şeydir. İnsanlar bunu sessizce konuştuklarında yapar.)")
                    pdf.drawString(36.85,599.96,"- Uzun cümleler içindeki ve cümleler arasındaki bilgiler arasında bağlantı kurarak bilgi çıkarımında bulunma,")
                    pdf.drawString(36.85,588.62,"- Ana fikri bulma ve yabancı kelimelerin anlamını çıkarma ve metinlerdeki önemli bilgileri bulma.")
                    pdf.drawString(36.85,571.62,"Next Steps")
                    pdf.drawString(36.85,557.44,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,546.10,"- Study new, unfamiliar words")
                    pdf.drawString(36.85,534.77,"- Practice reading stories and informational texts about a variety of topics")
                    pdf.drawString(36.85,523.43,"- Practice reading longer and more complex texts")
                    pdf.drawString(36.85,512.09,"- Speak or write in their own words about stories and information they read")
                    pdf.drawString(36.85,489.41,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,478.07,"- Yeni, yabancı kelimeler öğrenebilir.")
                    pdf.drawString(36.85,466.73,"- Çeşitli konularda hikaye ve bilgilendirici metin okuma alıştırmaları yapabilir.")
                    pdf.drawString(36.85,455.40,"- Daha uzun ve daha karmaşık metinlerle okuma alıştırması yapabilir.")
                    pdf.drawString(36.85,444.06,"- Okuduğu hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    
                def step2tk_4badgesreading(pdf):
                    pdf.setFont('abc', 8)
                    #BURDA HATA VAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    pdf.drawString(36.85,702.01,"Öğrenci basit hikayeleri ve yaşına uygun akademik metinleri anlar.Okuma becerileri:")
                    #YUKARDADADADDD
                    pdf.drawString(36.85,690.67,"- Understand a variety of common words and many less common words about objects, places, people, actions, and ideas")
                    pdf.drawString(36.85,679.33,"- Comprehend the meanings of complex sentences and paragraphs")
                    pdf.drawString(36.85,667.99,"- Connect information in longer sentences and across several sentences to infer information,")
                    pdf.drawString(36.85,656.66,"- Main ideas, the meaning of unfamiliar words and Identify specific details in texts")
                    pdf.drawString(36.85,633.98,"Öğrenci basit hikayeleri ve yaşına uygun akademik metinleri anlar.Okuma becerileri;")
                    pdf.drawString(36.85,622.64,"- Nesneler, yerler, insanlar, eylemler, fikirler için yaygın olarak kullanılan kelimeleri ve yaygın olarak kullanılmayan çoğu kelimeyi anlar.")
                    pdf.drawString(36.85,611.3,"- Karmaşık cümleler ve paragrafların anlamlarını kavrar.")
                    pdf.drawString(36.85,599.96,"- Uzun cümleler içindeki ve cümleler arasındaki bilgiler arasında bağlantı kurarak bilgi çıkarımında bulunma,")
                    pdf.drawString(36.85,588.62,"- Ana fikri bulma,yabancı kelimelerin anlamını çıkarma ve metinlerdeki belirli ayrıntıları tanımlama")
                    pdf.drawString(36.85,571.62,"Next Steps")
                    pdf.drawString(36.85,557.44,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,546.1,"- Read longer and more complex stories and informational texts about a variety of topics")
                    pdf.drawString(36.85,534.77,"- Speak or write in their own words about stories and information they read")
                    pdf.drawString(36.85,512.09,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,500.75,"- Çeşitli konular hakkında daha uzun ve daha karmaşık hikayeler ve bilgilendirici metinler okuyabilir.")
                    pdf.drawString(36.85,489.41,"- Okuduğu hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")

                def step2tk_5badgesreading(pdf):
                    pdf.setFont('abc', 8)
                    pdf.drawString(36.85,702.01,"Students perform exceptionally well on this test. They can;")
                    pdf.drawString(36.85,690.67,"- Understand a wide variety of common and less common words to describe objects, places, people, actions, and ideas")
                    pdf.drawString(36.85,679.33,"- Comprehend the meaning of complex sentences, paragraphs, and longer texts")
                    pdf.drawString(36.85,667.99,"- Connect information across several sentences and paragraphs to infer information, identify main ideas, and understand the meaning of unfamiliar")
                    pdf.drawString(36.85,656.66," words")
                    pdf.drawString(36.85,645.32,"- Identify specific details in longer texts")
                    pdf.drawString(36.85,622.44,"Öğrenci bu testte son derece iyi performans göstermiştir. Okuma becerileri:")
                    pdf.drawString(36.85,611.3,"- Nesneleri, yerleri, insanları, eylemleri ve fikirleri tanımlamak için çok çeşitli,yaygın ve yaygın kullanılmayan kelimeleri anlama")
                    pdf.drawString(36.85,599.96,"- Karmaşık cümlelerin, paragrafların ve daha uzun metinlerin anlamını kavrama")
                    pdf.drawString(36.85,588.62,"- Birkaç cümle ve paragraf arasında bağlantı kurarak çıkarımda bulunma, ana fikirleri tanımlama ve bilinmeyen kelimelerin anlamını anlama")
                    pdf.drawString(36.85,577.29,"- Daha uzun metinlerde belirli ayrıntıları bulma")
                    pdf.drawString(36.85,560.28,"Next Steps")
                    pdf.drawString(36.85,546.1,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,534.77,"- Read longer and more complex stories and academic texts about a variety of topics")
                    pdf.drawString(36.85,523.43,"- Speak or write in their own words about stories and information they read")
                    pdf.drawString(36.85,512.09,"- Consider taking the TOEFL Junior® test for more accurate information about their reading ability")
                    pdf.drawString(36.85,489.41,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,478.07,"- Çeşitli konular hakkında daha uzun ve daha karmaşık hikayeler ve akademik metinler okuyabilir.")
                    pdf.drawString(36.85,466.73,"- Okuduğu hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    pdf.drawString(36.85,455.4,"- Okuma becerileri hakkında daha doğru ve ayrıntılı sonuçlar almak için TOEFL Junior® sınavına girmeyi düşünebilir.")

                #Listening yıldıza göre bilgiler
                b2 = 11.34
                b1 = 384.53
                b3 = 17.01
                b4 = 14.16
                def step2tk_1badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students begin to recognize a few familiar words in speech, such as words for objects, places, and people. They may be able to;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand familiar words with visual support")
                    pdf.drawString(36.85,b1-b2*3,"Ögrenci, nesneler, yerler ve insanlar için kullanılan bazı tanıdık kelimeleri tanımaya baslar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*4,"- Görsel destekle temel kelimeleri anlayabilir.")
                    pdf.drawString(36.85,b1-b3-b2*4,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*4,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*5,"- Learn everyday words for objects and people in familiar categories such as home, school, family, colors, body parts, and animals")
                    pdf.drawString(36.85,b1-b3-b4-b2*6,"- Practice having short, simple conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*7,"- Practice listening to teacher instructions and short messages")
                    pdf.drawString(36.85,b1-b3-b4-b2*8,"- Begin listening to and identifying information in short, simple stories")
                    pdf.drawString(36.85,b1-b3-b4-b2*9,"- Consider taking the TOEFL Primary Step 1 test for more information about their listening ability")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Ev, okul, aile, renkler, vücudun bölümleri ve hayvanlar gibi bilindik kategorilerde kullanılan günlük kelimeleri ögrenebilir")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Kısa, basit konusmalarla pratik yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Ögretmen talimatlarını ve kısa mesajları dinlemeyi deneyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Kısa ve basit hikayelerdeki bilgileri dinlemeye ve tanımlamaya baslayabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"- Dinleme becerileri hakkında daha detaylı bilgi için TOEFL Primary Step 1 sınavını almayı düsünebilir.")
                    
                def step2tk_2badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students understand basic conversations and messages and begin to understand stories and informational texts. They can;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand common expressions used in everyday conversations")
                    pdf.drawString(36.85,b1-b2*2,"- Understand a simple, single instruction spoken in familiar words, with key words repeated")
                    pdf.drawString(36.85,b1-b2*3,"- Understand the purpose of messages in which key information is repeated")
                    pdf.drawString(36.85,b1-b2*4,"- Understand the main ideas of simple stories in which key information is explicitly stated and repeated")
                    pdf.drawString(36.85,b1-b2*6,"Öğrenci temel konuşmaları ve mesajları anlar ve hikayeleri ve bilgilendirici metinleri anlamaya başlar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*7,"- Günlük konuşmalarda yaygın olarak kullanılan ifadeleri anlama")
                    pdf.drawString(36.85,b1-b2*8,"- Anahtar kelimelerin tekrarlandığı, tanıdık kelimelerle konuşulan basit, tek bir talimatı anlama")
                    pdf.drawString(36.85,b1-b2*9,"- Önemli bilgilerin tekrarlandığı mesajların amacını anlama")
                    pdf.drawString(36.85,b1-b2*10,"- Önemli bilgilerin açıkça ifade edildiği ve tekrarlandığı basit hikayelerin ana fikirlerini anlama")
                    pdf.drawString(36.85,b1-b3-b2*10,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"To improve their listening ability, students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Study more words that describe familiar topics, settings, and actions")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Listen to age-appropriate academic talks and longer stories")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*17,"- Bilinen konular, durumlar ve eylemleri tanımlayan daha fazla kelime öğrenmeye çalışabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*18,"- Konuşmalarda daha az bilinen kelimeler ve ifadeler kullanarak alıştırma yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*19,"- Yaşına uygun akademik konuşmaları ve daha uzun hikayeleri dinleyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*20,"- Dinlediği hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    
                def step2tk_3badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students understand conversations and simple stories. They begin to understand age-appropriate academic talks. They can;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand less common words that describe familiar topics, settings, actions (e.g. pocket, pour, branch)")
                    pdf.drawString(36.85,b1-b2*2,"- Understand indirect responses to questions in conversations")
                    pdf.drawString(36.85,b1-b2*3,"- Understand messages in which information is not explicitly stated")
                    pdf.drawString(36.85,b1-b2*4,"- Connect information to infer the main idea or topic of messages, stories, and informational texts")
                    pdf.drawString(36.85,b1-b2*5,"- Synthesize information from multiple locations in a longer spoken text")
                    pdf.drawString(36.85,b1-b2*7,"Öğrenci konuşmaları ve basit hikayeleri anlar.Yaşına uygun akademik konuşmaları anlamaya başlar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*8,"- Tanıdık konuları, durumları ve eylemleri tanımlayan az yaygın olarak kullanılan kelimeleri anlama,(örn: cep, dökmek, dal)")
                    pdf.drawString(36.85,b1-b2*9,"- Sohbetlerdeki sorulara dolaylı olarak verilen yanıtları anlama")
                    pdf.drawString(36.85,b1-b2*10,"- Bilgilerin açıkça belirtilmediği mesajları anlama")
                    pdf.drawString(36.85,b1-b2*11,"- Hikaye, mesaj ve bilgilendirme amaçlı metinlerde bilgileri birbirine bağlayarak ana fikir hakkında çıkarım yapma")
                    pdf.drawString(36.85,b1-b2*12,"- Uzun konuşma metinlerinin içinde, farklı yerlerde geçen bilgileri sentezleme")
                    pdf.drawString(36.85,b1-b3-b2*12,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"To improve their listening ability, students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Learn new, unfamiliar words they hear in longer stories and academic talks")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,b1-b3-b4-b2*17,"Dinleme becerilerini geliştirmek için, öğrenci:")
                    pdf.drawString(36.85,b1-b3-b4-b2*18,"- Uzun öykülerde ve akademik konuşmalarda duyduğu yeni, tanıdık olmayan kelimeleri öğrenebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*19,"- Konuşmalarda daha az yaygın kelimeler ve ifadeler kullanarak alıştırma yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*20,"- Dinlediği hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    
                def step2tk_4badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students understand conversations, simple stories, and age-appropriate academic talks. They can;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand less frequently used words that describe familiar topics, settings, and actions")
                    pdf.drawString(36.85,b1-b2*2,"- Understand messages and stories that include unfamiliar words and some idiomatic expressions")
                    pdf.drawString(36.85,b1-b2*3,"- Consistently connect information throughout stories and academic talks to infer meaning")
                    pdf.drawString(36.85,b1-b2*4,"- Identify specific information in longer texts")
                    pdf.drawString(36.85,b1-b2*6,"Öğrenci konuşmaları, basit hikayeleri ve yaşına uygun akademik konuşmaları anlar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*7,"- Bilindik konuları, ortamları ve eylemleri tanımlamak için yaygın olarak kullanılmayan kelimeleri anlama.")
                    pdf.drawString(36.85,b1-b2*8,"- Bilmediği kelimeleri ve bazı deyimleri içeren mesajları ve hikayeleri anlama")
                    pdf.drawString(36.85,b1-b2*9,"- Hikayelerde ve akademik konuşmalarda geçen bilgileri birbirine bağlayarak anlam çıkarımında bulunma")
                    pdf.drawString(36.85,b1-b2*10,"- Uzun metinlerdeki istenen bilgileri anlama")
                    pdf.drawString(36.85,b1-b3-b2*10,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"To improve their listening ability, students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Listen to longer and more complex stories and academic texts about a variety of topics")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"Dinleme becerilerini geliştirmek için, öğrenci:")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"- Çeşitli konular hakkında daha uzun ve karmaşık hikayeleri ve akademik metinleri dinleyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*17,"- Konuşmalarda daha az karşılaşılan kelimeler ve ifadeler kullanarak alıştırma yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*18,"- Dinlediği hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    
                def step2tk_5badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students perform exceptionally well on this test. They can;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand less frequently used words that describe familiar topics, settings, and actions")
                    pdf.drawString(36.85,b1-b2*2,"- Understand messages and stories that include unfamiliar words and some idiomatic expressions")
                    pdf.drawString(36.85,b1-b2*3,"- Consistently connect information throughout stories and academic talks to infer meaning")
                    pdf.drawString(36.85,b1-b2*4,"- Identify specific information in longer texts")
                    pdf.drawString(36.85,b1-b2*6,"Öğrenci bu testte son derece iyi performans göstermiştir.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*7,"- Tanıdık konu, ortam ve eylemleri tanımlayan daha az yaygın kelimeleri anlama")
                    pdf.drawString(36.85,b1-b2*8,"- Bilmediği kelimeleri ve bazı deyimleri içeren mesajları ve hikayeleri anlama")
                    pdf.drawString(36.85,b1-b2*9,"- Hikayelerde ve akademik konuşmalarda geçen bilgileri birbirine bağlayarak anlam çıkarımında bulunma")
                    pdf.drawString(36.85,b1-b2*10,"- Uzun metinlerde istenen bilgileri anlama")
                    pdf.drawString(36.85,b1-b3-b2*10,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*10,"To improve their listening ability, students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"- Listen to longer and more complex stories and academic texts about a variety of topics")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Practice using less common words and expressions in conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Speak or write in their own words about stories and information they listen to")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Consider taking the TOEFL Junior® test for more accurate information about their listening ability")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"Dinleme yeteneğini geliştirmek için öğrenciler;")
                    pdf.drawString(36.85,b1-b3-b4-b2*17,"- Çesitli konular hakkında daha uzun ve daha karmaşık hikayeleri ve akademik metinleri dinleyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*18,"- Konuşmalarda yeni öğrendiği kelimeleri ve ifadeleri kullanarak alıştırma yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*19,"- Dinlediği hikayeler ve bilgiler hakkında, kendi cümleleriyle, konuşarak veya yazarak kendisini ifade edebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*20,"- Dinleme becerileri hakkında daha doğru bilgi elde etmek için TOEFL Junior® sınavına girmeyi düşünebilir.")
                
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
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 2")
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
                        pdf.drawString(510,741.69,"Badges: "+str(rbadges[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(510,424.21,"Badges: "+str(lbadges[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(rscore)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(lscore)+" puan almıştır.)")
                        

                    #Anabaskı kodu
                    baskialani(pdf)

                    #Reading skora göre sonuç seçimi
                    if rscore == str("100"):
                        step2tk_1badgesreading(pdf)
                    else:
                        if rscore == str("104"):
                            step2tk_2badgesreading(pdf)
                        elif rscore == str("105"):
                            step2tk_2badgesreading(pdf)
                        elif rscore == str("106"):
                            step2tk_2badgesreading(pdf)
                        elif rscore == str("107"):
                            step2tk_3badgesreading(pdf)
                        elif rscore == str("108"):
                            step2tk_3badgesreading(pdf)
                        elif rscore == str("109"):
                            step2tk_3badgesreading(pdf)
                        elif rscore == str("110"):
                            step2tk_4badgesreading(pdf)
                        elif rscore == str("111"):
                            step2tk_4badgesreading(pdf)
                        elif rscore == str("112"):
                            step2tk_4badgesreading(pdf)
                        elif rscore == str("113"):
                            step2tk_5badgesreading(pdf)
                        elif rscore == str("114"):
                            step2tk_5badgesreading(pdf)
                        elif rscore == str("115"):
                            step2tk_5badgesreading(pdf)


                    #Listening skora göre sonuç seçimi
                    if lscore == str("100"):
                        step2tk_1badgeslistening(pdf)
                    else:
                        if lscore == str("104"):
                            step2tk_2badgeslistening(pdf)
                        elif lscore == str("105"):
                            step2tk_2badgeslistening(pdf)
                        elif lscore == str("106"):
                            step2tk_2badgeslistening(pdf)
                        elif lscore == str("107"):
                            step2tk_3badgeslistening(pdf)
                        elif lscore == str("108"):
                            step2tk_3badgeslistening(pdf)
                        elif lscore == str("109"):
                            step2tk_3badgeslistening(pdf)
                        elif lscore == str("110"):
                            step2tk_4badgeslistening(pdf)
                        elif lscore == str("111"):
                            step2tk_4badgeslistening(pdf)
                        elif lscore == str("112"):
                            step2tk_4badgeslistening(pdf)
                        elif lscore == str("113"):
                            step2tk_5badgeslistening(pdf)
                        elif lscore == str("114"):
                            step2tk_5badgeslistening(pdf)
                        elif lscore == str("115"):
                            step2tk_5badgeslistening(pdf)

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
                def step2tk_1badgeslistening(pdf):
                    pdf.drawString(36.85,b1,"Students begin to recognize a few familiar words in speech, such as words for objects, places, and people. They may be able to;")
                    pdf.drawString(36.85,b1-b2*1,"- Understand familiar words with visual support")
                    pdf.drawString(36.85,b1-b2*3,"Ögrenci, nesneler, yerler ve insanlar için kullanılan bazı tanıdık kelimeleri tanımaya baslar.Dinleme becerileri:")
                    pdf.drawString(36.85,b1-b2*4,"- Görsel destekle temel kelimeleri anlayabilir.")
                    pdf.drawString(36.85,b1-b3-b2*4,"Next Steps")
                    pdf.drawString(36.85,b1-b3-b4-b2*4,"To improve their listening ability,students should;")
                    pdf.drawString(36.85,b1-b3-b4-b2*5,"- Learn everyday words for objects and people in familiar categories such as home, school, family, colors, body parts, and animals")
                    pdf.drawString(36.85,b1-b3-b4-b2*6,"- Practice having short, simple conversations")
                    pdf.drawString(36.85,b1-b3-b4-b2*7,"- Practice listening to teacher instructions and short messages")
                    pdf.drawString(36.85,b1-b3-b4-b2*8,"- Begin listening to and identifying information in short, simple stories")
                    pdf.drawString(36.85,b1-b3-b4-b2*9,"- Consider taking the TOEFL Primary Step 1 test for more information about their listening ability")
                    pdf.drawString(36.85,b1-b3-b4-b2*11,"Dinleme becerilerini geliştirmek için öğrenci")
                    pdf.drawString(36.85,b1-b3-b4-b2*12,"- Ev, okul, aile, renkler, vücudun bölümleri ve hayvanlar gibi bilindik kategorilerde kullanılan günlük kelimeleri ögrenebilir")
                    pdf.drawString(36.85,b1-b3-b4-b2*13,"- Kısa, basit konusmalarla pratik yapabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*14,"- Ögretmen talimatlarını ve kısa mesajları dinlemeyi deneyebilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*15,"- Kısa ve basit hikayelerdeki bilgileri dinlemeye ve tanımlamaya baslayabilir.")
                    pdf.drawString(36.85,b1-b3-b4-b2*16,"- Dinleme becerileri hakkında daha detaylı bilgi için TOEFL Primary Step 1 sınavını almayı düsünebilir.")
                    
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
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 2")
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
                        pdf.drawString(510,741.69,"Badges: "+str(rbadges[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(510,424.21,"Badges: "+str(lbadges[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(rscore)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(lscore)+" puan almıştır.)")
                        

                    #Anabaskı kodu
                    baskialani(pdf)
                    step2tk_1badgeslistening(pdf)


                    pdf.save()
            
            #Türkçe karne readingi ns, listeningi ns olmayan sonuç
            if rscore == "NS" and lscore != "NS":
                rscore2 = "100"
                totalscore = str(math.ceil((int(rscore2)+int(lscore))/2))

                #Reading yıldıza göre bilgiler
                def step2tk_1badgesreading(pdf):
                    pdf.drawString(36.85,702.01,"Students begin to recognize some basic words. They may be able to;")
                    pdf.drawString(36.85,690.67,"- Identify basic vocabulary with visual support")
                    pdf.drawString(36.85,667.99,"Öğrenci, bazı temel kelimeleri tanımaya başlayabilir.Dinleme becerileri:")
                    pdf.drawString(36.85,656.66,"- Görsel destek yardımıyla temel kelimeleri anlama")
                    pdf.drawString(36.85,639.65,"Next Steps")
                    pdf.drawString(36.85,625.47,"To improve their reading ability, students should;")
                    pdf.drawString(36.85,614.14,"- Learn words and common expressions used in familiar social settings")
                    pdf.drawString(36.85,602.80,"- Learn words that show relationships among people, objects,and places (examples: at, on)")
                    pdf.drawString(36.85,591.46,"- Practice reading simple sentences and short texts about familiar topics")
                    pdf.drawString(36.85,580.12,"- Consider taking the TOEFL Primary Step 1 test for more information about their reading ability")
                    pdf.drawString(36.85,557.44,"Okuma becerilerini geliştirmek için öğrenci:")
                    pdf.drawString(36.85,546.10,"- Günlük hayatta kullanılan temel kelime ve ifadeleri öğrenebilir")
                    pdf.drawString(36.85,534.77,"- İnsanlar, nesneler ve yerler arasındaki ilişkileri tanımlayan ifadeleri öğrenebilir.(örnekler: at, on)")
                    pdf.drawString(36.85,523.43,"- Bilindik konular hakkında basit cümleler ve kısa metinlerle okuma alıştırmaları yapabilir.")
                    pdf.drawString(36.85,512.09,"- Okuma becerileri hakkında daha fazla bilgi edinmek için TOEFL Primary Step 1 sınavına girmeyi düşünebilir.")

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
                        pdf.drawString(334.49,775.71,"TOEFL PRIMARY STEP 2")
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
                        pdf.drawString(510,741.69,"Badges: "+str(rbadges[0]))
                        pdf.drawString(160,424.21,"Score: "+str(lscore)+",")
                        pdf.drawString(210,424.21,"CEFR: "+str(lcefr))
                        pdf.drawString(510,424.21,"Badges: "+str(lbadges[0]))
                        pdf.drawString(41.85,724.69,"The Student received "+str(rscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(rscore)+" puan almıştır.)")
                        pdf.drawString(41.85,407.2,"The Student received "+str(lscore)+" on a scale of 104 to 115 (Öğrenci 104 ile 115 arasındaki ölçekte "+str(lscore)+" puan almıştır.)")
                
                    #Anabaskı kodu
                    baskialani(pdf)
                    step2tk_1badgesreading(pdf)
                    
                    pdf.save()

        
        step2_scorereport()
        step2_certificate()
        step2_tk()

def step2_classicbutton():
    buttons()
    tr1 = threading.Thread(target=step2_classic)
    tr1.start()
