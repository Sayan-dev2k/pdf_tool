from tkinter import *
from datetime import date
from tkinter import ttk,messagebox #for creating the treeview
from docxtpl import DocxTemplate
doc=DocxTemplate("Invoice-Template.docx")
invoice_list=[]
ino=100
def gen_invoice():
    
    name=fentry.get()+' '+lentry.get()
    today=date.today()
    datee=str(today.day)+'-'+str(today.month)+'-'+str(today.year)
    
    subtotal=0
    for item in invoice_list:
        subtotal=subtotal+int(item[len(item)-1])
    if(subtotal>=300):
        discount=(10/100)*subtotal
        total=(subtotal-discount)*1.05
        doc.render({'Name':name,'Phone':phentry.get(),'Address':addentry.get(),'Email':emailentry.get(),'d':datee,'ino':ino,
                'invoice_list':invoice_list,'subtotal':subtotal,'discount':'10%','tax':'5%','total':total})
    else:
        total=subtotal*1.05
        doc.render({'Name':name,'Phone':phentry.get(),'Address':addentry.get(),'Email':emailentry.get(),'d':datee,'ino':ino,
                'invoice_list':invoice_list,'subtotal':subtotal,'discount':'0%','tax':'5%','total':total})
    # file=filedialog.asksaveasfilename(defaultextension='.docx')
    file=str(ino)+'.docx'
    doc.save(file)
    messagebox.showinfo('success','invoice generated successfully')
def new_invoice():
    global ino
    ino=ino+1
    invlbl.config(text=ino)
    fentry.delete(0,END)
    lentry.delete(0,END)
    addentry.delete(0,END)
    phentry.delete(0,END)
    emailentry.delete(0,END)
    treeview.delete(*treeview.get_children())
    invoice_list.clear()
def clear_item():
    qty_spinbox.delete(0,END)
    qty_spinbox.insert(0,'1')
    desentry.delete(0,END)
    pricentry.delete(0,END)
    pricentry.insert(0,'0')

def add_item():
    qty=int(qty_spinbox.get())
    desc=desentry.get()
    price=float(pricentry.get())
    total=qty*price
    inv_item=[desc,qty,price,total]
    invoice_list.append(inv_item)
    treeview.insert('',0,values=inv_item)
    clear_item()
root=Tk()
root.title('Invoice form')
root.geometry('1000x500')
root.config(bg='green')
frame=Frame(root,height=300)
frame.pack(padx=20,pady=10)
fname=Label(frame,text='First Name')
fname.grid(row=0,column=0)
lname=Label(frame,text='Last Name')
lname.grid(row=0,column=1)
global fentry
fentry=Entry(frame)
fentry.grid(row=1,column=0)
global lentry
lentry=Entry(frame)
lentry.grid(row=1,column=1)
address=Label(frame,text='Address')
address.grid(row=0,column=2)
global addentry
addentry=Entry(frame)
addentry.grid(row=1,column=2)
phone=Label(frame,text='Phone no.')
phone.grid(row=0,column=3)
global phentry
phentry=Entry(frame)
phentry.grid(row=1,column=3)
email=Label(frame,text='Email ID')
email.grid(row=0,column=4)
global emailentry
emailentry=Entry(frame)
emailentry.grid(row=1,column=4)
invoice_no=Label(frame,text='Invoice no.')
invoice_no.grid(row=2,column=4)
invlbl=Label(frame,bg='white',text=ino)
invlbl.grid(row=3,column=4)
qty=Label(frame,text='Qty')
qty.grid(row=2,column=0)
qty_spinbox=Spinbox(frame,from_=1,to=100)
qty_spinbox.grid(row=3,column=0)
desc=Label(frame,text='Description')
desc.grid(row=2,column=1)
desentry=Entry(frame)
desentry.grid(row=3,column=1)
unit_price=Label(frame,text='Unit Price')
unit_price.grid(row=2,column=2)
pricentry=Entry(frame)
pricentry.grid(row=3,column=2)
butn=Button(frame,text='Add item',bg='skyblue',command=add_item)
butn.grid(row=4,column=3)
columns=('desc','Qty','price','total')
treeview=ttk.Treeview(frame,columns=columns,show='headings')
treeview.heading('Qty',text='Qty')
treeview.heading('desc',text='Description')
treeview.heading('price',text='Unit Price')
treeview.heading('total',text='Total')
treeview.grid(row=5,column=0,columnspan=6,pady=10)
save=Button(frame,text='Generate Invoice',bg='yellow',command=gen_invoice)
save.grid(row=6,column=1,sticky='news')#sticky=news means expand in north east west south
new_invoice=Button(frame,text='New Invoice',bg='yellow',command=new_invoice)
new_invoice.grid(row=6,column=3,sticky='news')
root.mainloop()
