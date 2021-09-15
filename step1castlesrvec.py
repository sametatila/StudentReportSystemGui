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


def step1_castle():
    global p1,window1,toplamsatir,satir,filename,f123
    buttons()
    import pdfplumber, re, math
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
        p1["value"] = satir+1
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
        rstar = str(sheet['L'+str(satir)].value)
        lstar = str(sheet['O'+str(satir)].value)

        #ETS Sonuç Belge Numarası ve Cinsiyet Sorgulama
        x = [os.path.join(r,file) for r,d,f in os.walk(scorefolder) for file in f if file.endswith(str(studentnumber)+".PDF")]

        with pdfplumber.open(x[0]) as pdf1:
            page = pdf1.pages[0]
            text = page.extract_text()
            #belge no buradan gelir
            name = re.compile(r'OYD.*')
            #sadece step 1 classic score report tan alınabilir. cinsiyet buradan gelir
            gender = re.compile(r'Gender:.*')
        for line in text.split('\n'):
            if name.match(line):
                lname = line.split()
                print(lname[5])
        for line in text.split('\n'):
            if gender.match(line):
                lgender = line.split()

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
                pdf.drawImage(step1c_nsimage, 481,477.45, width=19.5,height=18.2,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.45, 429.82, 490.95, 429.82)
                pdf.line(542.15, 429.82, 798.65, 429.82)
                #Part1
                pdf.setFont('abc', 8)
                pdf.drawString(233.8,413.1,"The test taker did not respond to any questions in this section.")
                pdf.drawString(233.8,403.9,"Therefore, the scores for this section cannot be provided.")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(803.53,493.93,"CEFR Level: "+str(rcefr)+"  |  "+"Lexile Measure: "+str(lexile))
                pdf.setFont('abc', 8)
                pdf.drawRightString(803.53,484.41,"The student received "+str(rscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(474.25,483.04,"Student's Level: "+rstar[0]+rstar[1]+" out of 4 Stars")

            def step1_1starreading(pdf):
                #1star Image
                pdf.drawImage(step1c_1starimage, 481,480.27, width=81,height=15.95,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.45, 429.82, 490.95, 429.82)
                pdf.line(542.15, 429.82, 798.65, 429.82)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,409.8,nokta)
                pdf.drawString(559.1,410.2,nokta)
                pdf.drawString(559.1,381.88,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.75,450.9,"Students begin to recognize some basic words. They")
                pdf.drawString(233.75,438.83,"may be able to:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,412.56,"Identify basic vocabulary with visual support")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,451.6,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,412.96,"Learn and practice reading common words in familiar")
                pdf.drawString(577.9,403.76,"categories such as home, school, family, colors, body parts,")
                pdf.drawString(577.9,394.56,"animals, and actions")
                pdf.drawString(577.9,384.82,"Read short, simple sentences about familiar people, objects,")
                pdf.drawString(577.9,375.62,"and actions (example:")
                pdf.setFont('abcit', 8)
                pdf.drawString(658.39,375.62,"The boy is eating an apple.")
                pdf.setFont('abc', 8)
                pdf.drawString(754.46,375.62,")")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(803.53,493.93,"CEFR Level: "+str(rcefr)+"  |  "+"Lexile Measure: "+str(lexile))
                pdf.setFont('abc', 8)
                pdf.drawRightString(803.53,484.41,"The student received "+str(rscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(474.25,483.04,"Student's Level: "+rstar[0]+rstar[1]+" out of 4 Stars")

            def step1_2starreading(pdf):
                #2star Image
                pdf.drawImage(step1c_2starimage, 480,480.67, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.45, 429.82, 490.95, 429.82)
                pdf.line(542.15, 429.82, 798.65, 429.82)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,409.8,nokta)
                pdf.drawString(251,381.3,nokta)
                pdf.drawString(251,371.5,nokta)
                pdf.drawString(559.1,410.2,nokta)
                pdf.drawString(559.1,391,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.75,450.9,"Students begin to understand words and some short")
                pdf.drawString(233.75,438.83,"descriptions. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,412.56,"Understand common words in familiar categories such as")
                pdf.drawString(269.8,403.36,"home, school, family, colors, body parts, animals, and")
                pdf.drawString(269.8,394.16,"actions")
                pdf.drawString(269.8,384.42,"Recognize key words for understanding simple sentences")
                pdf.drawString(269.8,374.68,"Understand everyday actions in the present (")
                pdf.setFont('abcit', 8)
                pdf.drawString(429,374.68,"examples: The")
                pdf.drawString(269.8,365.48,"children play. He is eating.")
                pdf.setFont('abc', 8)
                pdf.drawString(363.63,365.48,")")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,451.6,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,412.96,"Learn vocabulary and common expressions used in social")
                pdf.drawString(577.9,403.76,"and familiar settings")
                pdf.drawString(577.9,394.02,"Practice reading simple sentences and short texts about")
                pdf.drawString(577.9,384.82,"familiar topics")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(803.53,493.93,"CEFR Level: "+str(rcefr)+"  |  "+"Lexile Measure: "+str(lexile))
                pdf.setFont('abc', 8)
                pdf.drawRightString(803.53,484.41,"The student received "+str(rscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(474.25,483.04,"Student's Level: "+rstar[0]+rstar[1]+" out of 4 Stars")

            def step1_3starreading(pdf):
                #3star Image
                pdf.drawImage(step1c_3starimage, 480,480.67, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.45, 429.82, 490.95, 429.82)
                pdf.line(542.15, 429.82, 798.65, 429.82)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,409.8,nokta)
                pdf.drawString(251,390.45,nokta)
                pdf.drawString(251,362.68,nokta)
                pdf.drawString(251,334.15,nokta)
                pdf.drawString(559.1,410.2,nokta)
                pdf.drawString(559.1,391,nokta)
                pdf.drawString(559.1,372.2,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.75,450.9,"Students understand short descriptions and find")
                pdf.drawString(233.75,438.83,"information in signs, forms, and schedules. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,412.56,"Understand common words and social expressions")
                pdf.drawString(269.8,403.36,"(")
                pdf.setFont('abcit', 8)
                pdf.drawString(272.46,403.36,"examples: play a game, go to a museum, wave goodbye")
                pdf.setFont('abc', 8)
                pdf.drawString(472.58,403.36,")")
                pdf.drawString(269.8,393.62,"Comprehend simple descriptions of current and past events")
                pdf.drawString(269.8,384.42,"(")
                pdf.setFont('abcit', 8)
                pdf.drawString(272.46,384.42,"examples: The mouse is on top of the table. He is washing")
                pdf.drawString(269.8,375.22,"his hands.")
                pdf.setFont('abc', 8)
                pdf.drawString(306.27,375.22,")")
                pdf.drawString(269.8,365.48,"Recognize relationships among words and phrases within")
                pdf.drawString(269.8,356.29,"familiar categories (")
                pdf.setFont('abcit', 8)
                pdf.drawString(339.6,356.29,"examples: food-fruit-strawberries; rain-")
                pdf.drawString(269.8,347.09,"sky-clouds; one more time-again")
                pdf.setFont('abc', 8)
                pdf.drawString(385.41,347.09,")")
                pdf.drawString(269.8,337.35,"Make connections across simple sentences (")
                pdf.setFont('abcit', 8)
                pdf.drawString(428.98,337.35,"example:")
                pdf.drawString(269.8,328.15,"Clouds are in the sky. Rain comes from them. Sometimes")
                pdf.drawString(269.8,318.95,"they cover the sun.")
                pdf.setFont('abc', 8)
                pdf.drawString(337.39,318.95,")")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,451.6,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,412.96,"Read longer paragraphs and stories about familiar people,")
                pdf.drawString(577.9,403.76,"objects, and information")
                pdf.drawString(577.9,394.02,"Learn more words that describe objects, places, people,")
                pdf.drawString(577.9,384.82,"actions, and ideas")
                pdf.drawString(577.9,375.08,"Speak or write in their own words about paragraphs, stories,")
                pdf.drawString(577.9,365.88,"and information they read")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(803.53,493.93,"CEFR Level: "+str(rcefr)+"  |  "+"Lexile Measure: "+str(lexile))
                pdf.setFont('abc', 8)
                pdf.drawRightString(803.53,484.41,"The student received "+str(rscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(474.25,483.04,"Student's Level: "+rstar[0]+rstar[1]+" out of 4 Stars")

            def step1_4starreading(pdf):
                #4star Image
                pdf.drawImage(step1c_4starimage, 480,480.67, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.45, 429.82, 490.95, 429.82)
                pdf.line(542.15, 429.82, 798.65, 429.82)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,409.8,nokta)
                pdf.drawString(251,381.3,nokta)
                pdf.drawString(251,353.3,nokta)
                pdf.drawString(251,325,nokta)
                pdf.drawString(559.1,410.2,nokta)
                pdf.drawString(559.1,400,nokta)
                pdf.drawString(559.1,381.3,nokta)
                pdf.drawString(559.1,371.5,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.75,450.9,"Students understand short descriptions, information in")
                pdf.drawString(233.75,438.83,"signs, and short messages. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,412.56,"Understand common words and some less common words")
                pdf.drawString(269.8,403.36,"about objects, places, people, actions, and ideas (")
                pdf.setFont('abcit', 8)
                pdf.drawString(447.26,403.36,"examples:")
                pdf.drawString(269.8,394.16,"ring, adventures, whisper, double")
                pdf.setFont('abc', 8)
                pdf.drawString(387.65,394.16,")")
                pdf.drawString(269.8,384.42,"Comprehend the meaning of complex sentences (")
                pdf.setFont('abcit', 8)
                pdf.drawString(446.79,384.42,"examples:")
                pdf.drawString(269.8,375.22,"This is a friendly thing to do when you say goodbye. People")
                pdf.drawString(269.8,366.02,"do this when they talk quietly.")
                pdf.setFont('abc', 8)
                pdf.drawString(374.3,366.02,")")
                pdf.drawString(269.8,356.29,"Connect information in longer sentences and across")
                pdf.drawString(269.8,347.09,"different sentences to infer information, identify main ideas,")
                pdf.drawString(269.8,337.89,"and understand the meaning of unfamiliar words.")
                pdf.drawString(269.8,328.15,"Locate key information in texts")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,451.6,"To improve their reading ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,412.96,"Study new, unfamiliar words")
                pdf.drawString(577.9,403.22,"Practice reading stories and informational texts about a")
                pdf.drawString(577.9,394.02,"variety of topics")
                pdf.drawString(577.9,384.28,"Practice reading longer and more complex texts")
                pdf.drawString(577.9,374.54,"Speak or write in their own words about stories and")
                pdf.drawString(577.9,365.35,"information they read")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(803.53,493.93,"CEFR Level: "+str(rcefr)+"  |  "+"Lexile Measure: "+str(lexile))
                pdf.setFont('abc', 8)
                pdf.drawRightString(803.53,484.41,"The student received "+str(rscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(474.25,483.04,"Student's Level: "+rstar[0]+rstar[1]+" out of 4 Stars")

            #Listening yıldıza göre bilgiler

            def step1_nslistening(pdf):
                #NS Image
                pdf.drawImage(step1c_nsimage, 481,244, width=19.5,height=18.2,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.49, 187.39, 490.99, 187.39)
                pdf.line(543.05, 187.11, 799.55, 187.11)
                #Part1
                pdf.setFont('abc', 8)
                pdf.drawString(233.8,170.8,"The test taker did not respond to any questions in this section.")
                pdf.drawString(233.8,161.6,"Therefore, the scores for this section cannot be provided.")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(799.75,254.61,"CEFR Level: "+str(lcefr))
                pdf.setFont('abc', 8)
                pdf.drawRightString(799.75,245.09,"The student received "+str(lscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9)
                pdf.drawRightString(474.34,248.58,"Student's Level: "+lstar[0]+lstar[1]+" out of 4 Stars")

            def step1_1starlistening(pdf):
                #1star Image
                pdf.drawImage(step1c_1starimage, 480,246.8, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.49, 187.39, 490.99, 187.39)
                pdf.line(543.05, 187.11, 799.55, 187.11)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,167.15,nokta)
                pdf.drawString(559.1,167.35,nokta)
                pdf.drawString(559.1,139.2,nokta)
                pdf.drawString(559.1,129.6,nokta)
                pdf.drawString(559.1,110.8,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.95,211.7,"Students begin to recognize some familiar words in")
                pdf.drawString(233.95,199.63,"speech, such as words for objects, places, and people.")
                pdf.drawString(233.95,187.55,"They may be able to:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,170.26,"Understand familiar words with visual support")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,211.7,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,170.46,"Learn everyday words for objects and people in familiar")
                pdf.drawString(577.9,161.26,"categories such as home, school, family, colors, body parts,")
                pdf.drawString(577.9,152.06,"and animals")
                pdf.drawString(577.9,142.32,"Use pictures to help learn new words")
                pdf.drawString(577.9,132.58,"Listen to short, simple sentences about everyday actions,")
                pdf.drawString(577.9,123.38,"objects, and people. (example:")
                pdf.setFont('abcit', 8)
                pdf.drawString(689.08,123.38,"She is swimming.")
                pdf.setFont('abc', 8)
                pdf.drawString(751.32,123.38,")")
                pdf.drawString(577.9,113.65,"Practice using common, everyday expressions, such as")
                pdf.drawString(577.9,104.45,"greetings")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(799.75,254.61,"CEFR Level: "+str(lcefr))
                pdf.setFont('abc', 8)
                pdf.drawRightString(799.75,245.09,"The student received "+str(lscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9)
                pdf.drawRightString(474.34,248.58,"Student's Level: "+lstar[0]+lstar[1]+" out of 4 Stars")
                pdf.setStrokeColorRGB(255,255,255)
                pdf.setLineWidth(2.1)
                pdf.line(234.49, 188.962, 490.99, 188.962)
                pdf.setLineWidth(2)
                pdf.line(234.49, 185.95, 490.99, 185.95)

            def step1_2starlistening(pdf):
                #2star Image
                pdf.drawImage(step1c_2starimage, 480,246.8, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.49, 187.39, 490.99, 187.39)
                pdf.line(543.05, 187.11, 799.55, 187.11)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,167.15,nokta)
                pdf.drawString(251,139,nokta)
                pdf.drawString(559.1,167.35,nokta)
                pdf.drawString(559.1,148.5,nokta)
                pdf.drawString(559.1,138.5,nokta)
                pdf.drawString(559.1,120,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.95,211.7,"Students begin to recognize some familiar words in")
                pdf.drawString(233.95,199.63,"speech. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,170.26,"Understand words for objects and people in familiar")
                pdf.drawString(269.8,161.06,"categories such as school, home, family, colors, body parts,")
                pdf.drawString(269.8,151.86,"and animals")
                pdf.drawString(269.8,142.12,"Recognize action words in simple sentences (")
                pdf.setFont('abcit', 8)
                pdf.drawString(432.1,142.12,"examples: The")
                pdf.drawString(269.8,132.92,"children play. He is eating.")
                pdf.setFont('abc', 8)
                pdf.drawString(363.63,132.92,")")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,211.7,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,170.46,"Practice saying and listening to familiar words used in simple")
                pdf.drawString(577.9,161.26,"sentences")
                pdf.drawString(577.9,151.52,"Practice having short, simple conversations")
                pdf.drawString(577.9,141.78,"Practice listening to messages spoken by teachers, friends,")
                pdf.drawString(577.9,132.58,"and family")
                pdf.drawString(577.9,122.85,"Begin listening to and identifying basic information in short,")
                pdf.drawString(577.9,113.65,"simple stories")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(799.75,254.61,"CEFR Level: "+str(lcefr))
                pdf.setFont('abc', 8)
                pdf.drawRightString(799.75,245.09,"The student received "+str(lscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9)
                pdf.drawRightString(474.34,248.58,"Student's Level: "+lstar[0]+lstar[1]+" out of 4 Stars")

            def step1_3starlistening(pdf):
                #3star Image
                pdf.drawImage(step1c_3starimage, 480,246.8, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.49, 187.39, 490.99, 187.39)
                pdf.line(543.05, 187.11, 799.55, 187.11)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,167.15,nokta)
                pdf.drawString(251,148.2,nokta)
                pdf.drawString(251,129.5,nokta)
                pdf.drawString(251,110.4,nokta)
                pdf.drawString(559.1,167.35,nokta)
                pdf.drawString(559.1,148.5,nokta)
                pdf.drawString(559.1,129.5,nokta)
                pdf.drawString(559.1,120,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.95,211.7,"Students understand short, simple descriptions,")
                pdf.drawString(233.95,199.63,"conversations, and messages. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,170.26,"Understand common expressions used in everyday")
                pdf.drawString(269.8,161.06,"conversations")
                pdf.drawString(269.8,151.32,"Understand a simple, single instruction spoken in familiar")
                pdf.drawString(269.8,142.12,"words, with key words repeated")
                pdf.drawString(269.8,132.38,"Understand the purpose of messages in which key information")
                pdf.drawString(269.8,123.18,"is repeated")
                pdf.drawString(269.8,113.45,"Understand the main ideas of simple stories in which key")
                pdf.drawString(269.8,104.25,"information is explicitly stated and repeated")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,211.7,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,170.46,"Study more words that describe familiar topics, settings, and")
                pdf.drawString(577.9,161.26,"actions")
                pdf.drawString(577.9,151.52,"Practice using less common words and expressions in")
                pdf.drawString(577.9,142.32,"conversations")
                pdf.drawString(577.9,132.58,"Listen to age-appropriate academic talks and longer stories")
                pdf.drawString(577.9,122.85,"Speak or write in their own words about stories and")
                pdf.drawString(577.9,113.65,"information they listen to")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(799.75,254.61,"CEFR Level: "+str(lcefr))
                pdf.setFont('abc', 8)
                pdf.drawRightString(799.75,245.09,"The student received "+str(lscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9)
                pdf.drawRightString(474.34,248.58,"Student's Level: "+lstar[0]+lstar[1]+" out of 4 Stars")

            def step1_4starlistening(pdf):
                #4star Image
                pdf.drawImage(step1c_4starimage, 480,246.8, width=81,height=15.3,mask=None)
                #Çizgiler
                pdf.setLineWidth(0.665)
                pdf.line(234.49, 187.39, 490.99, 187.39)
                pdf.line(543.05, 187.11, 799.55, 187.11)
                #Noktalar
                pdf.setFont('dot', 14)
                pdf.drawString(251,167.15,nokta)
                pdf.drawString(251,148.2,nokta)
                pdf.drawString(251,138.8,nokta)
                pdf.drawString(251,119.5,nokta)
                pdf.drawString(251,100.7,nokta)
                pdf.drawString(559.1,167.35,nokta)
                pdf.drawString(559.1,148.5,nokta)
                pdf.drawString(559.1,129.3,nokta)
                #Part1
                pdf.setFont('abc', 10.5)
                pdf.drawString(233.95,211.7,"Students understand simple descriptions, instructions,")
                pdf.drawString(233.95,199.63,"conversations, and messages. They can:")
                pdf.setFont('abc', 8)
                pdf.drawString(269.8,170.26,"Understand less common words that describe familiar topics,")
                pdf.drawString(269.8,161.06,"settings, and actions (")
                pdf.setFont('abcit', 8)
                pdf.drawString(347.62,161.06,"examples: pocket, pour, lamp, branch")
                pdf.setFont('abc', 8)
                pdf.drawString(481.02,161.06,")")
                pdf.drawString(269.8,151.32,"Understand indirect responses to questions in conversations")
                pdf.drawString(269.8,141.58,"Understand messages in which information is not explicitly")
                pdf.drawString(269.8,132.38,"stated")
                pdf.drawString(269.8,122.65,"Connect information to infer the main idea or topic of")
                pdf.drawString(269.8,113.45,"messages, stories, and informational texts")
                pdf.drawString(269.8,103.71,"Synthesize information from multiple locations in a longer")
                pdf.drawString(269.8,94.51,"spoken text")
                #Part2
                pdf.setFont('abc', 10.5)
                pdf.drawString(541.9,211.7,"To improve their listening ability, students should:")
                pdf.setFont('abc', 8)
                pdf.drawString(577.9,170.46,"Learn new, unfamiliar words they hear in longer stories and")
                pdf.drawString(577.9,161.26,"academic talks")
                pdf.drawString(577.9,151.52,"Practice using less common words and expressions in")
                pdf.drawString(577.9,142.32,"conversations")
                pdf.drawString(577.9,132.58,"Speak or write in their own words about stories and")
                pdf.drawString(577.9,123.38,"information they listen to")
                #Part3
                pdf.setFont('abc', 9.5)
                pdf.drawRightString(799.75,254.61,"CEFR Level: "+str(lcefr))
                pdf.setFont('abc', 8)
                pdf.drawRightString(799.75,245.09,"The student received "+str(lscore)+" on a scale of 100 to 109")
                pdf.setFont('abc', 9)
                pdf.drawRightString(474.34,248.58,"Student's Level: "+lstar[0]+lstar[1]+" out of 4 Stars")

            #Belge özellikleri
            fileName = outputfolder+asd+str(sheet['D'+str(satir)].value)+str(sheet['E'+str(satir)].value)[0]+"_"+str(studentnumber)+str("_ScoreReport.PDF")
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
                #Sınav Türü
                pdf.setFont('abcbold', 13)
                pdf.drawString(585.75,533.84 ,"Reading")
                pdf.drawString(641.65,533.84 ,"and")
                pdf.drawString(669.06,533.84 ,"Listening")
                pdf.drawString(730.88,533.84 ,"Test")
                pdf.drawString(761.82,533.84 ,"-")
                pdf.drawString(770.7,533.84 ,"Step")
                pdf.drawString(803.28,533.84 ,"1")
                pdf.drawString(681.29,559.71 ,"Official Score Report")
                #Öğrenci Bilgisi
                pdf.setFont('abcbold', 8.5)
                pdf.drawString(25.16,516.32,"Student Name:")
                pdf.drawString(25.38,497.9,"Student Number:")
                pdf.drawString(25.38,462.5,"Test Date:")
                pdf.drawString(25.38,480.2,"Date of Birth: ")
                pdf.drawString(25.38,444.8,"Gender: ")
                pdf.setFont('abc', 10)
                pdf.drawString(88.8,516.32,str(studentname))
                pdf.drawString(98.09,497.9,str(studentnumber))
                pdf.drawString(70.19,462.5,str(testdate))
                pdf.drawString(83.48,480.2,str(dateofbirth))
                pdf.drawString(61.99,444.8,str(lgender[1]))
                #Alt bilgi Okul vs.
                pdf.setFont('abc', 6.5)
                pdf.drawCentredString(514.92,24.97,str(school)+", Turkey")
                pdf.drawCentredString(514.92,17.5,"OYD - Okul Yayin, Turkey")
                pdf.drawString(770.57,15.5,str(lname[5]))

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
                pdf.drawImage(step1c_nsimage, 548.66,179.1, width=19.6,height=18.2,mask=None)

            def step1c_1starreading(pdf):
                pdf.drawImage(step1c_1starimage, 548,182.71, width=80.5,height=15.375,mask=None)

            def step1c_2starreading(pdf):
                pdf.drawImage(step1c_2starimage, 548,182.71, width=80.5,height=15.375,mask=None)

            def step1c_3starreading(pdf):
                pdf.drawImage(step1c_3starimage, 548,182.71, width=80.5,height=15.375,mask=None)

            def step1c_4starreading(pdf):
                pdf.drawImage(step1c_4starimage, 548,182.71, width=80.5,height=15.375,mask=None)

            #Listening sonuçları
            def step1c_nslistening(pdf):
                pdf.drawImage(step1c_nsimage, 548.66,156, width=19.6,height=18.2,mask=None)

            def step1c_1starlistening(pdf):
                pdf.drawImage(step1c_1starimage, 548,159.6, width=80.5,height=15.375,mask=None)

            def step1c_2starlistening(pdf):
                pdf.drawImage(step1c_2starimage, 548,159.6, width=80.5,height=15.375,mask=None)

            def step1c_3starlistening(pdf):
                pdf.drawImage(step1c_3starimage, 548,159.6, width=80.5,height=15.375,mask=None)

            def step1c_4starlistening(pdf):
                pdf.drawImage(step1c_4starimage, 548,159.6, width=80.5,height=15.375,mask=None)

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
            pdfmetrics.registerFont(TTFont('abcbi', './data/arialbi.ttf'))

            #Ana baskı alanı
            def baskialani(pdf):
                #Sertifika içeriği
                pdf.setFont('abcbold', 14.5)
                pdf.drawCentredString(562.07,291.55,str(studentname))
                pdf.setFont('abcbold', 13)
                pdf.drawString(352.35,230.66,"Earned the following levels on the")
                pdf.drawString(666.82,230.66,"™")
                pdf.drawString(683.5,230.66,"Test  - Step 1")
                pdf.drawString(471.6,185.63,"Reading:")
                pdf.drawString(471.6,163.2,"Listening:")
                pdf.setFont('abcbi', 13)
                pdf.drawString(566.21,230.66,"TOEFL")
                pdf.drawString(618.39,230.66,"Primary")
                pdf.setFont('abcbold', 8.50)
                pdf.drawString(608.57,234.66,"®")

                
                #Alt bilgi Okul vs.
                pdf.setFont('abc', 7.5)
                pdf.drawString(346.25,100.61,"Test Date: "+str(testdate))
                pdf.drawString(346.25,89.5,str(school)+", Turkey")
                pdf.drawString(346.25,78.53,"OYD - Okul Yayin, Turkey")
                pdf.setFont('abc', 5.5)
                pdf.drawString(346.25,68.43,str(lname[5]))
                
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
    

def step1_castlebutton():
    buttons()
    tr1 = threading.Thread(target=step1_castle)
    tr1.start()
    
