from tkinter import *
from tkinter import messagebox,filedialog
from tkhtmlview import HTMLLabel
import os
import pandas as pd 
from PIL.ImageTk import PhotoImage
import code_email
import time
from PIL import ImageTk
from tkinterhtml import TkinterHtml

root = Tk()

class Email:
    def __init__(self, root):
        self.root = root
        self.root.title("Intent Amplify:")
        self.root.state("zoomed")
        self.root.resizable(False, False)  
        self.root.config(bg="lightgrey")

        self.Email_icon = ImageTk.PhotoImage(file="intent.png")
        self.Setting_icon=ImageTk.PhotoImage(file="setting.png")

        self.var_choice = StringVar()

        single = Radiobutton(root, text="Single", value="single", command=self.check_single_OR_bulk,
                             activebackground="skyblue", variable=self.var_choice, font=("Helvetica", 16, "bold"),
                             bg="skyblue", fg="black")
       

        multiple = Radiobutton(root, text="Multiple", value="multiple", command=self.check_single_OR_bulk,
                               variable=self.var_choice, activebackground="skyblue",
                               font=("Helvetica", 16, "bold"), bg="skyblue", fg="black")
        multiple.place(relx=0.5, rely=0.2, anchor=CENTER)
        self.var_choice.set("single")


        title = Label(self.root, text="Amplify Blasting",font=("Goudy Old Style", 48, "bold"), bg="#f68421",
                      fg="black").place(x=0, y=0, relwidth=1)
        
        btn_set1 = Button(self.root, image=self.Email_icon, bg="black", bd='0', command=LEFT,
                          height=self.Email_icon.height(), width=self.Email_icon.width(),
                          activebackground="blue")
        btn_set1.place(x=0, y=0)



        To = Label(self.root, text="To(Email)", font=("Calibri", 18, "bold"), bg="skyblue",
                   fg="black").place(x=50, y=200)
        Subject = Label(self.root, text="Subject", font=("Calibri", 18, "bold"), bg="skyblue",
                        fg="black").place(x=50, y=250)
        Message = Label(self.root, text="Message", font=("Calibri", 18, "bold"), bg="skyblue",
                        fg="black").place(x=50, y=300)

       

        self.Total = Label(self.root, font=("times new roman", 18, "bold"), bg="lightgrey", fg="black")
        self.Total.place(x=50, y=500)

        self.Sent = Label(self.root, font=("times new roman", 18, "bold"), bg="lightgrey", fg="darkgreen")
        self.Sent.place(x=350, y=500)

        self.Left = Label(self.root, font=("times new roman", 18, "bold"), bg="lightgrey", fg="orange")
        self.Left.place(x=450, y=500)

        self.Failed = Label(self.root, font=("times new roman", 18, "bold"), bg="lightgrey", fg="red")
        self.Failed.place(x=550, y=500)




        self.message_entry = HTMLLabel(self.root, font=("times new roman", 15), bg="white")
        self.message_entry.place(x=280, y=300, width=700, height=170)

        self.to_entry = Entry(self.root, font=("times new roman", 15), bg="white")
        self.to_entry.place(x=280, y=200, width=350, height=30)

        self.sub_entry = Entry(self.root, font=("times new roman", 15), bg="white")
        self.sub_entry.place(x=280, y=250, width=450, height=30)

        


        btn1 = Button(root, activebackground="skyblue", command=self._email, text="SEND",
                      font=("times new roman", 20, "bold"), bg="Green",
                      fg="black").place(x=700, y=500, width=130, height=30)
        btn2 = Button(root, activebackground="skyblue", command=self.clear1, text="CLEAR",
                      font=("times new roman", 20, "bold"), bg="red",
                      fg="black").place(x=850, y=500, width=130, height=30)
        self.btn3 = Button(root, activebackground="skyblue", text="Browse", font=("times new roman", 20, "bold"),
                           bg="lightblue", command=self.Browse_button, cursor="hand2", state=DISABLED, fg="black")
        self.btn3.place(x=650, y=200, width=150, height=30)

        self.check_file_exist()
    def Browse_button(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel File for Emails",
                                    filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")))
        if op != None:
            data = pd.read_excel(op.name)

            if 'Email' in data.columns:
                self.EMAIL = list(data['Email'])
                
                c = []
                for i in self.EMAIL:
                    
                    if (pd.isnull(i)) == False:
                        
                        c.append(i)
                self.EMAIL = c
                if len(self.EMAIL) > 0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(op.name.split("/")[-1]))
                    self.to_entry.config(state='readonly')
                    self.Total.config(text="Total: " + str(len(self.EMAIL)) )
                    self.Sent.config(text="Sent: ")
                    self.Left.config(text="Left: ")
                    self.Failed.config(text="Failed: ")
                
            else:
                messagebox.showinfo("Error", "select Email File", parent=self.root)

    def send_email(self):
        x = len(self.message_entry.get('1.0', END))
        if self.to_entry.get() == "" or self.sub_entry.get() == "" or x == 1:
            messagebox.showerror("ERROR", "All feilds are required", parent=self.root)
        else:
            if self.var_choice.get() == "single":
                status=code_email.Email_send_function(self.to_entry.get(),self.sub_entry.get(),self.message_entry.get('1.0',END),self.uname,self.pasw)
                if status=="s":
                    messagebox.showinfo("SUCCESS","Email Sent", parent=self.root)
                if status=="f":
                    messagebox.showerror("Failed","Email Not Sent", parent=self.root)

            if self.var_choice.get()=="multiple":
                self.failed = []
                self.s_count=0
                self.f_count =0
                for x in self.EMAIL:
                    status=code_email.Email_send_function(x,self.sub_entry.get(),self.message_entry.get('1.0',END),self.uname,self.pasw)

                    if status=="s":
                       self.s_count+=1
                    if status=="f":
                       self.f_count+=1
                    self.status_bar()
                    time.sleep(1)


                messagebox.showinfo("Success","Emails are send Thanks for choose Intent-Amplify....", parent=self.root)

    def clear1(self):
        self.to_entry.config(state=NORMAL)
        self.to_entry.delete(0, END)
        self.sub_entry.delete(0, END)
        self.message_entry.delete('1.0', END)
        self.var_choice.set("single")
        self.btn3.config(state=DISABLED)
        self.Total.config(text="")
        self.Sent.config(text="")
        self.Left.config(text="")
        self.Failed.config(text="")


    def status_bar(self):
        self.Total.config(text="Status " + str(len(self.EMAIL))+":-")
        self.Sent.config(text="Sent: "+ str(self.s_count))
        self.Left.config(text="Left: "+ str(len(self.EMAIL)-(self.f_count+self.s_count)))
        self.Failed.config(text="Failed: "+ str(self.f_count))
        self.Total.update()
        self.Sent.update()
        self.Left.update()
        self.Failed.update()

    def check_single_OR_bulk(self):
        if self.var_choice.get() == "single":
            messagebox.showinfo("single","Switch To Single Mail Sender", parent=self.root)
            self.btn3.config(state=DISABLED)
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.clear1()

        if self.var_choice.get() == "multiple":
            messagebox.showinfo("multiple","Switch To Bulk mail sender", parent=self.root)

            self.btn3.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.to_entry.config(state='readonly')

    def setting_clear(self):
        self.uname_entry.delete(0, END)
        self.pasw_entry.delete(0, END)

    def setting_window(self):
        self.check_file_exist()
        self.root2 = Toplevel()
        self.root2.title("Setting")
        self.root2.resizable(False, False)
        self.root2.state("zoomed")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="lightgrey")
        title2 = Label(self.root2, text="Bulk Email Sender", padx=10, compound=LEFT,
                       font=("Goudy Old Style", 48, "bold"), bg="black",
                       fg="white").place(x=0, y=0, relwidth=1)
        REF2 = Label(self.root2, text="Enter your valid Email Id and Password", font=("Calibri (body)", 14,),
                     bg="yellow", fg="black").place(x=0, y=80, relwidth=1)

        uname = Label(self.root2, text="Email Address", font=("times new roman", 18, "bold"), bg="lightgrey",
                      fg="black").place(x=50, y=150)

        pasw = Label(self.root2, text="Password", font=("times new roman", 18, "bold"), bg="lightgrey",
                     fg="black").place(x=50, y=200)

        self.uname_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow")
        self.uname_entry.place(x=250, y=150, width=330, height=30)

        self.pasw_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow", show="*")
        self.pasw_entry.place(x=250, y=200, width=330, height=30)


        btn1 = Button(self.root2, activebackground="skyblue", text="SEND", font=("times new roman", 20, "bold"),
                      bg="black",
                      fg="white", command=self.save_setting).place(x=250, y=250, width=130, height=30)
        btn2 = Button(self.root2, activebackground="skyblue", text="CLEAR", font=("times new roman", 20, "bold"),
                      bg="#ffcccb", command=self.setting_clear,
                      fg="black").place(x=400, y=250, width=130, height=30)

        self.uname_entry.insert(0, self.uname)
        self.pasw_entry.insert(0, self.pasw)


    def check_file_exist(self):
        if os.path.exists("important.txt") == False:
            f = open('important.txt','w')
            f.write(",")
            f.close()
        f2 = open('important.txt', 'r')
        self.credentials = []
        for i in f2:
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        # print(self.credentials)
        self.uname = self.credentials[0][0]
        self.pasw = self.credentials[0][1]
        # print(self.uname,self.pasw)

    def save_setting(self):
        if self.uname_entry.get() == "" or self.pasw_entry.get() == "":
            messagebox.showinfo("ERROR", "All feilds are requiFred", parent=self.root2)

        else:
            f = open('important.txt', 'w')
            f.write(self.uname_entry.get() + "," + self.pasw_entry.get())
            f.close()
            messagebox.showinfo("Sent","Email and password are saved Successfully",parent=self.root2)
            self.check_file_exist()

obj = Email(root)
root.mainloop()
