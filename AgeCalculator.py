import tkinter as tk
from tkinter import messagebox
from tkinter import*
import os
import openpyxl

window = tk.Tk()

file=os.path.join(os.getcwd(),'data.xlsx')
from datetime import date

today = date.today()


def click(*args):
    ename.delete(0,END)
    

def clearall():
    ename.delete(0,END)
    ename.insert(0,'Enter your First Name')
    t1.config(state='normal')
    t2.config(state='normal')
    t3.config(state='normal')
    tshow.config(state=NORMAL)
    tshow.delete(0,END)
    e1.delete(0, END)
    e2.delete(0, END)
    e3.delete(0, END)
    t1.delete(0, END)
    t2.delete(0, END)
    t3.delete(0, END)
  

def checkError() :

	if (e1.get() == "" or e2.get() == ""
		or e3.get() == "" ) :

		messagebox.showerror("Input Error")

		clearall()
		
		return 0

def get_age():
    
    
    value = checkError()

    if value == 0 :
        
        return
    
    
    else:
        name=(ename.get())
        ed= int(e1.get())
        em=int(e2.get())
        ey=int(e3.get())
        
        
        cd=today.day
        cm=today.month
        cy=today.year
    
        total_month =[31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        

        if (ed > cd):
            cm = cm - 1
            cd = cd + total_month[em-1]
                
        if (em > cm):
            cy = cy - 1
            cm = cm + 12
            
            
            
        # calculate day, month, year
        cal_day = cd - ed;
        cal_month = cm - em;
        cal_year = cy - ey;
        
        # tshow.config(state=NORMAL)
  
        t1.config(state='normal')
        t2.config(state='normal')
        t3.config(state='normal')
        tshow.config(state='normal')
        tshow.delete(0,END)
        
        t1.insert(tk.END," "+str(cal_year)+' Y')
        t1.config(state='disabled')
        t2.insert(tk.END," "+str(cal_month)+' M')
        t2.config(state='disabled')
        t3.insert(tk.END," "+str(cal_day)+' D')
        t3.config(state='disabled')
        tshow.insert(tk.END,name+", You are now "+ str(cal_year)+" Years Old." )
        tshow.config(state='disabled')
        
        
       
        if not os.path.exists(file):
            print(file)
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(['Name', 'Age', 'DOB', 'MOB', 'YOB', 'Search_Date'])
            sheet.append([name,cal_year,ed,em,ey,date.today()])
            try:
                workbook.save(file)
                print("File created successfully.")
            except Exception as e:
                print(f"Error saving file: {str(e)}")
        else:

            workbook=openpyxl.load_workbook(file)
            sheet = workbook.active
            sheet.append([name,cal_year,ed,em,ey,date.today()])
            workbook.save(file)
            
    
  
def exit():
    window.destroy()

window.geometry("400x300")
window.config(bg="#F7DC6F")
window.resizable(width=False,height=False)
window.title('Age Calculator!')


if __name__ == "__main__" :

    l1 = tk.Label(window,text="Age Calculator!",font=("Arial", 20,"underline"),fg="black",bg="#F7DC6F")

    l2 = tk.Label(window,font=("Arial",12),text="Enter your Date of Birth :",fg="black",bg="#F7DC6F")
    
    l_d=tk.Label(window,text="Date: ",font=('Arial',12,"bold"),fg="darkgreen",bg="#F7DC6F")
    l_m=tk.Label(window,text="Month: ",font=('Arial',12,"bold"),fg="darkgreen",bg="#F7DC6F")
    l_y=tk.Label(window,text="Year: ",font=('Arial',12,"bold"),fg="darkgreen",bg="#F7DC6F")
    
    ename=tk.Entry(window,width=20)
    ename.insert(0,"Enter Your First Name")
    ename.bind("<Button-1>",click)
    
    e1=tk.Entry(window,width=5)
    e2=tk.Entry(window,width=5)
    e3=tk.Entry(window,width=5)
    
    b1=tk.Button(window,text="Calculate Age!",font=("Arial",13),command=get_age)

    b2=tk.Button(window,text="clear",font=("Arial",13),command=clearall)
    
    l3 = tk.Label(window,text="Exact Age is: ",font=('Arial',12,"bold"),fg="darkgreen",bg="#F7DC6F")


    t1=tk.Entry(window,width=5,state="disabled")
    t2=tk.Entry(window,width=5,state="disabled")
    t3=tk.Entry(window,width=5,state="disabled")
    
    tshow = tk.Entry(window,width=55)
    tshow.insert(tk.END,"Welcome")
    tshow.config(state=DISABLED)
    
    b3=tk.Button(window,text="Exit Application!",font=("Arial",13),command=exit)
    ename.place(x=150,y=40)
    l1.place(x=100,y=3)
    l2.place(x=10,y=58)
    l_d.place(x=100,y=78)
    l_m.place(x=100,y=103)
    l_y.place(x=100,y=128)
    e1.place(x=180,y=78)
    e2.place(x=180,y=103)
    e3.place(x=180,y=128)
    b1.place(x=100,y=155)
    b2.place(x=240,y=155)
    l3.place(x=75,y=225)
    t1.place(x=225,y=225)
    t2.place(x=275,y=225)
    t3.place(x=325,y=225)
    tshow.place(x=50,y=203)
    b3.place(x=100,y=250)


    window.mainloop()