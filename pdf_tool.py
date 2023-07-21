from tkinter import *
from tkinter import filedialog
from PIL import Image,ImageTk
from pdf2docx import parse
from typing import Tuple
from PyPDF2 import PdfMerger,PdfWriter,PdfReader
from docx2pdf import convert
from win32com import client
import sys,os
import pdftables_api
# import tabula
from pdf2image import convert_from_path
import img2pdf
s=0
p=0
def upload_img():
    global filepath
    filepath=filedialog.askopenfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',filetypes=(('Image file','*.jpg'),('all files','*.*'))) 
def upload_excel():
    global filepath
    filepath=filedialog.askopenfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',filetypes=(('Excel workbook','*.xlsx'),('all files','*.*')))
    
def excel_to_pdf():
    global filepath
    app=client.DispatchEx('Excel.Application')
    
    app.Interactive=False
    app.Visible=False
    # f='C:\\Users\\SAYAN\Desktop\\python projects\\ex.xlsx'
    # f=str(filepath)
    workbook=app.Workbooks.open(filepath)
    # output=os.path.splitext(filepath)[0]
    output='C:\\Users\\SAYAN\Desktop\\python projects\\excel_pdf_converted.pdf'
    # output=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
    # print(filepath)
    # f=filepath.split('.')
    # print(f)
    
    # output=f[0]+'.pdf'
    # print(output)
    
    workbook.ActiveSheet.ExportAsFixedFormat(0,output)
    workbook.Close()
    lbl=Label(root,text='CONVERTED TO PDF SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def pdf_to_excel():
    #apikey conversion
    api=getapi.get()
    conversion = pdftables_api.Client(api)
    file=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.xlsx')
    
    
    conversion.xlsx(filepath,file)
    lbl=Label(root,text='CONVERTED TO EXCEL SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def jpg_to_pdf():
    img_path =filepath
 
# storing pdf path
    pdf_path =filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
    
    # opening image
    image = Image.open(img_path)
    
    # converting into chunks using img2pdf
    pdf_bytes = img2pdf.convert(image.filename)
    
    # opening or creating pdf file
    file = open(pdf_path, "wb")
    
    # writing pdf files with chunks
    file.write(pdf_bytes)
    
    # closing image file
    image.close()
    
    # closing pdf file
    file.close()
    lbl=Label(root,text='CONVERTED TO PDF SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def pdf_to_jpg():
    image=convert_from_path(filepath,500,poppler_path='C:\\Program Files\\poppler-23.05.0\\Library\\bin')
    folderpath=filedialog.askdirectory(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',title='select folder')
    for i in range(len(image)):
        image[i].save(folderpath+'/'+'page'+str(i)+'.jpg','JPEG')
    lbl=Label(root,text='CONVERTED TO JPG SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def upload_pdfs():
    global pdfs
    a=int(entrypdf.get())
    i=0
    pdfs=[]
    while(i<a):
        filepath=filedialog.askopenfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',filetypes=(('PDF files','*.pdf'),('all files','*.*')))
        pdfs.append(filepath)
        i=i+1
def no_pdfs():
    global entrypdf
    noofpdf=Label(root,text='ENTER NUMBER OF PDFS',font=('times new roman',20,'bold'),bg='#EADDCA')
    noofpdf.place(x=200,y=300)
    entrypdf=Entry(root,font=('times new roman',15,'bold'))
    entrypdf.place(x=590,y=300)
    setbtn=Button(root,text='SET',font=('times new roman',15,'bold'),command=upload_pdfs)
    setbtn.place(x=810,y=300)

def pdfmerge():
    merger=PdfMerger()
    for pdf in pdfs:
        merger.append(pdf)
    file=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
    merger.write(file)
    merger.close()
    lbl=Label(root,text='PDFs MERGED SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def pdfsplit():
    a1=int(A1.get())
    a2=int(A2.get())
    a3=int(A3.get())
    b1=int(B1.get())
    b2=int(B2.get())
    b3=int(B3.get())
    ch1=list(range(a1-1,b1))
    ch2=list(range(a2-1,b2))
    ch3=list(range(a3-1,b3))
    with open(filepath,'rb') as f:
        reader=PdfReader(f)
        ch1wirte=PdfWriter()
        ch2write=PdfWriter()
        ch3write=PdfWriter()

        for page in range(len(reader.pages)):
            if page in ch1:
                ch1wirte.add_page(reader.pages[page])
            elif page in ch2:
                ch2write.add_page(reader.pages[page])
            elif page in ch3:
                ch3write.add_page(reader.pages[page])
        if(b1>0):
            file1=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
            with open(file1,'wb') as f2:
                ch1wirte.write(f2)
        if(b2>0):
            file2=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
            with open(file2,'wb') as f2:
                ch2write.write(f2)
        if(b3>0):
            file3=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
            with open(file3,'wb') as f2:
                ch3write.write(f2)
        
    lbl=Label(root,text='PDFs SPLITTED SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)    
    
def word_to_pdf():
    global filepath
    op=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.pdf')
    convert(filepath,op)
    lbl=Label(root,text='FILE CONVERTED TO PDF SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def convert_pdf2docs(input_file:str,output_file:str,pages: Tuple=None):
    if pages:
        pages=[int(i) for i in list(pages) if i.isnumeric()]
    result=parse(pdf_file=input_file,docx_file=output_file,pages=pages)
    lbl=Label(root,text='FILE CONVERTED TO DOCX SUCCESSFULLY!!',bg='#FF4500',fg='white',font=('times new roman',15,'bold'))
    lbl.place(x=900,y=450)
def upload_word():
    global filepath
    filepath=filedialog.askopenfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',filetypes=(('DOCX files','*.docx'),('all files','*.*')))
def upload_pdf():
    global filepath,A1,A2,A3,B1,B2,B3,getapi
    filepath=filedialog.askopenfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',filetypes=(('PDF files','*.pdf'),('all files','*.*')))
    if(s==1):
        part1=Label(root,text='Part 1 Range',bg='#EADDCA',fg='red',font=('times new roman',15,'bold'))
        part1.place(x=200,y=300)
        page1=Label(root,text='Page',bg='#EADDCA',font=('times new roman',15,'bold'))
        page1.place(x=500,y=300)
        A1=Entry(root,width=5,font=('times new roman',15,'bold'))
        A1.place(x=600,y=300)
        to1=Label(root,text='-',bg='#EADDCA',font=('times new roman',15,'bold'))
        to1.place(x=700,y=300)
        B1=Entry(root,width=5,font=('times new roman',15,'bold'))
        B1.place(x=750,y=300)
        part2=Label(root,text='Part 2 Range',bg='#EADDCA',fg='red',font=('times new roman',15,'bold'))
        part2.place(x=200,y=350)
        page2=Label(root,text='Page',bg='#EADDCA',font=('times new roman',15,'bold'))
        page2.place(x=500,y=350)
        A2=Entry(root,width=5,font=('times new roman',15,'bold'))
        A2.place(x=600,y=350)
        to2=Label(root,text='-',bg='#EADDCA',font=('times new roman',15,'bold'))
        to2.place(x=700,y=350)
        B2=Entry(root,width=5,font=('times new roman',15,'bold'))
        B2.place(x=750,y=350)
        part3=Label(root,text='Part 3 Range',bg='#EADDCA',fg='red',font=('times new roman',15,'bold'))
        part3.place(x=200,y=400)
        page3=Label(root,text='Page',bg='#EADDCA',font=('times new roman',15,'bold'))
        page3.place(x=500,y=400)
        A3=Entry(root,width=5,font=('times new roman',15,'bold'))
        A3.place(x=600,y=400)
        to3=Label(root,text='-',bg='#EADDCA',font=('times new roman',15,'bold'))
        to3.place(x=700,y=400)
        B3=Entry(root,width=5,font=('times new roman',15,'bold'))
        B3.place(x=750,y=400)
    if(p==1):
            
            apilab=Label(root,text='Enter API Key',bg='#EADDCA',font=('times new roman',15,'bold'))
            apilab.place(x=400,y=280)
            lab=Label(root,text='(Go to PDFtables.com,generate a new api key)',bg='#EADDCA',font=('times new roman',15,'bold'))
            lab.place(x=400,y=310)
            getapi=Entry(root,font=('times new roman',15,'bold'))
            getapi.place(x=550,y=280)
            
def pdf_to_word():
    global filepath

    input_file=filepath
    output_file=filedialog.asksaveasfilename(initialdir='C:\\Users\\SAYAN\Desktop\\python projects',defaultextension='.docx')
    convert_pdf2docs(input_file,output_file)
def fun1():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='MERGE PDF',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD PDF FILES',font=('times new roman',20,'bold'),bg='#90EE90',command=no_pdfs)
    upload.pack(pady=120)
    merge=Button(root,text='MERGE FILES',font=('times new roman',20,'bold'),bg='#90EE90',command=pdfmerge)
    merge.place(x=580,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('mergeicon.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('merge_icon.png')
    img=ImageTk.PhotoImage(file='merge_icon.png')
    labl.config(image=img)
    labl.image=img
    

def fun2():
    global s
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='SPLIT PDF',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_pdf)
    upload.pack(pady=120)
    split=Button(root,text='SPLIT FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=pdfsplit)
    split.place(x=630,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=450)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('icon-splitpdf.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img
    s=s+1
def fun3():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='CONVERT PDF TO WORD',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_pdf)
    upload.pack(pady=120)
    word=Button(root,text='COVERT TO WORD FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=pdf_to_word)
    word.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('pdf-to-word.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img
def fun4():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    # root.config(bg='blue')
    titl=Label(root,text='CONVERT WORD TO PDF',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD WORD FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_word)
    upload.pack(pady=120)
    pdf=Button(root,text='CONVERT TO PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=word_to_pdf)
    pdf.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('word_to_pdf.jpg')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img
def fun5():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='CONVERT EXCEL TO PDF',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD EXCEL FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_excel)
    upload.pack(pady=120)
    pdf=Button(root,text='CONVERT TO PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=excel_to_pdf)
    pdf.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('xlpdf.jpg')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img

def fun6():
    global p
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='CONVERT PDF TO EXCEL',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_pdf)
    upload.pack(pady=120)
    excel=Button(root,text='CONVERT TO EXCEL FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=pdf_to_excel)
    excel.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('pdfxl.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img
    p=p+1
    print(p)
def fun7():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='CONVERT PDF TO JPG',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_pdf)
    upload.pack(pady=120)
    jpg=Button(root,text='CONVERT TO JPG FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=pdf_to_jpg)
    jpg.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('pdfimg.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img

def fun8():
    title.destroy()
    lbl1.destroy()
    merge_pdf.destroy()
    split_pdf.destroy()
    pdf_word.destroy()
    word_pdf.destroy()
    excel_pdf.destroy()
    pdf_excel.destroy()
    pdf_jpg.destroy()
    jpg_pdf.destroy()
    label.destroy()
    titl=Label(root,text='CONVERT JPG TO PDF',font=('times new roman',25,'bold'),bg='#FFBF00').place(x=0,y=20,relwidth=1)
    upload=Button(root,text='UPLOAD IMAGE IN JPG',font=('times new roman',20,'bold'),bg='#90EE90',command=upload_img)
    upload.pack(pady=120)
    jpg=Button(root,text='CONVERT TO PDF FILE',font=('times new roman',20,'bold'),bg='#90EE90',command=jpg_to_pdf)
    jpg.place(x=550,y=200)
    labl=Label(root,height=300,width=300,bg='white')
    labl.place(x=550,y=400)
    icon=Image.new("RGB",(400,400),('white'))
    bg=Image.open('imgpdf.png')
    bg=bg.resize((300,300),Image.ANTIALIAS)
    icon.paste(bg,(50,50))
    icon.save('_icon.png')
    img=ImageTk.PhotoImage(file='_icon.png')
    labl.config(image=img)
    labl.image=img

root=Tk()
root.geometry("1400x800")
root.title("Pdf tool")
title=Label(root,text='PDF TOOL',font=('times new roman',50,'bold'),bg='#40E0D0',fg='white',bd=5,relief='solid')
title.place(x=0,y=20,relwidth=1)
root.config(bg='#EADDCA')
lbl1=Label(root,text='CHOOSE AN OPTION',font=('times new roman',25,'bold'),bg='green',fg='white')
lbl1.pack(pady=120)
merge_pdf=Button(root,text='PDF MERGE',font=('times new roman',20,'bold'),bg='yellow',command=fun1)
merge_pdf.place(x=150,y=200)
split_pdf=Button(root,text='PDF SPLIT',font=('times new roman',20,'bold'),bg='yellow',command=fun2)
split_pdf.place(x=450,y=200)
pdf_word=Button(root,text='PDF TO WORD',font=('times new roman',20,'bold'),bg='yellow',command=fun3)
pdf_word.place(x=700,y=200)
word_pdf=Button(root,text='WORD TO PDF',font=('times new roman',20,'bold'),bg='yellow',command=fun4)
word_pdf.place(x=980,y=200)
excel_pdf=Button(root,text='EXCEL TO PDF',font=('times new roman',20,'bold'),bg='yellow',command=fun5)
excel_pdf.place(x=150,y=300)
pdf_excel=Button(root,text='PDF TO EXCEL',font=('times new roman',20,'bold'),bg='yellow',command=fun6)
pdf_excel.place(x=450,y=300)
pdf_jpg=Button(root,text='PDF TO JPG',font=('times new roman',20,'bold'),bg='yellow',command=fun7)
pdf_jpg.place(x=700,y=300)
jpg_pdf=Button(root,text='JPG TO PDF',font=('times new roman',20,'bold'),bg='yellow',command=fun8)
jpg_pdf.place(x=980,y=300)
label=Label(root,height=300,width=300,bg='white')
label.place(x=550,y=400)
icon=Image.new("RGB",(400,400),('white'))
bg=Image.open('pdf__icon.png')
bg=bg.resize((300,300),Image.ANTIALIAS)
icon.paste(bg,(50,50))
icon.save('picon.png')
img=ImageTk.PhotoImage(file='picon.png')
label.config(image=img)
label.image=img
root.mainloop()