from tkinter import ttk
from tkinter import *
import tkinter as tk
from tkcalendar import Calendar
import os as os
import re
import win32com.client
from datetime import date, datetime
import time
import pyodbc
from PIL import ImageTk, Image
import argparse, sys
from datetime import timedelta
import traceback
import pandas as pd
import numpy as np
from tkinter import filedialog


#main window properties
root = Tk()
root.title('Claim Automation Tracker')
root.geometry("850x480")

#Frame for the homepage
win = Frame(root,bg="white", highlightbackground="white", highlightthickness=2)
win.place(x=0,y=0,width=850,height=480)



#logo 
global filepath
pathnamelogo = os.path.abspath(os.path.dirname(__file__))
filepath = pathnamelogo
img = ImageTk.PhotoImage(Image.open(pathnamelogo+ "\\logo.png"))
panel = Label(win, image = img)
panel.place(x=215,y=40)
logoalttext = "xx@123"
img2 = ImageTk.PhotoImage(Image.open(pathnamelogo+ "\\xx.png"))
panel2 = Label(win, image = img2)
panel2.place(x=700,y=380)



#connection to the MS Access database
a_driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
dbpath = r''+ pathnamelogo +'\\claimsdb.accdb'
conn = pyodbc.connect(DRIVER=a_driver,DBQ=dbpath, autocommit=True,PWD=logoalttext)
#conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\n1569313\OneDrive - Liberty Mutual\Documents\claims.accdb;')
db = conn.cursor()
global loggedUser
global loggedUserAuth

db.execute("select name,auth from users where nid= ?", (os.getenv('username')))
for rows in db.fetchall():
    loggedUser = rows[0]
    loggedUserAuth = rows[1]


#Production Code
def prod():
    prodwin = Tk()
    prodwin.title('Production - Claim Automation Tracker')
    prodwin.geometry("850x480")
    
    
    
    #left Frame
    frame = Frame(prodwin,bg="#b5edff", highlightbackground="grey", highlightthickness=2)
    frame.place(x=0,y=0,width=280,height=210)
    #right Frame
    frame1 = Frame(prodwin,bg="#d1e1ff", highlightbackground="grey", highlightthickness=2)
    frame1.place(x=280,y=0,width=570,height=210)
    
    
    #defining dropdown events
    def OptionMenu_Select1(event):
        mail_box.config(text=mail.get())
    #def OptionMenu_Select2(event):
    #    wt.config(text=worktype.get())
    def OptionMenu_Select3(event):
        lob_label.config(text=lob.get())
    def OptionMenu_Select4(event):
        status_label.config(text=status.get())
    
    
    # Create the variables
    mail = StringVar()
    worktype = tk.StringVar()
    lob = StringVar()
    status = StringVar()
    def callback(*args):
        wt.config(text=worktype.get())
    
    #worktype dropdown
    wt=Label(frame1,text="Select Worktype",font = "Helvetica 8 bold", bg="#f0f0f0")
    wt.place(x=35, y=55, height=40)
    options2=[]
    db.execute('select distinct worktypes from workTypes order by worktypes')
    for row in db.fetchall():
        options2.append(row[0])
    #options2 = ['ACT BDX','Advise/Update','APC BDX','AWAC Bdx','Claim Closed','Claim Open','COLLEGIATE BDX','DACB Report (New Claims, Reserve Movement and Closure)','Daily Payment Report & Error Report','Decision tree','DWF Report (New Claims, Reserve Movement and Closure)','Escrow Top Up','Gibbs Hartley','GTS Centralised Mailbox','Incomplete Items reporting - 004','IRIS BDX','Loss Runs','MGA Casualty','MGA Property ','MID Loss Runs','MILAN BDX','Monthly Bordereaux (MGA/Motor/Asurion/IRIS)','MOTOL XL BDX','New Claim Setup','Pantaenius - Quarterly Reserve Report','Pantaenius - Settlement/Payments Statements','POINEER BDX','Policy Requests - Documents','Policy Requests - Information','PURGE Report','Recovery','Referrals request','Refund','Reserve Update','RFIB BDX','RI LLR Report ','RPC BDX','Settlement - Fees ','Settlement - Indemnity','Settlement - Queries','TPA Collegiate','TPA Davies','Uploads','Volante BDX']
    wtmenu = tk.OptionMenu(frame1, worktype, *(options2))
    wtmenu.place(x=5, y=55,  width=30, height=40)
    worktype.trace("w", callback)
    
    
    def getInfo():
        db.execute("select count(IIF(status = 'Pending', 1, NULL)) as pending, count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query from mailRecords where mailBox = ?",(mail.get()))
        for data in db.fetchall():
            pendingLb.config(text="Pending: ")
            holdLb.config(text="Hold: ")
            queryLb.config(text="Queried: ")
            pendingVl.config(text=data[0])
            holdVl.config(text=data[1])
            queryVl.config(text=data[2])
            
    def lobWt(mailbox):
    # Reset wt and delete all old options
        db.execute("select lob from maillob where shortmail = ? ",(mailbox))
        lob.set(db.fetchone()[0])
        lob_label.config(text=lob.get())
        
        worktype.set('')
        wtmenu['menu'].delete(0, 'end')
        
        db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailbox))
        for wts in db.fetchall():
            wtmenu['menu'].add_command(label=wts[0], command=tk._setit(worktype, wts[0]))
        wt.config(text="Select Worktype")
    
    #Creating tkinter treeview and scrollbar; Soft refresh function
    def softref():
        try:
            global my_table
            tree_scroll = Scrollbar(prodwin)
            tree_scroll.place(x=830,y=210,height=265)

            my_table = ttk.Treeview(prodwin,show='headings', height=12, yscrollcommand=tree_scroll.set)
            my_table['columns'] = ('messageID','mailstr','from', 'subject', 'recieved','assigned', 'mailbox', 'status')

            my_table.column("#0", width=0,  stretch=NO)
            my_table.column("messageID", width=0,  stretch=NO)
            my_table.column("mailstr", width=0,  stretch=NO)
            my_table.column("from",anchor=W, width=120)
            my_table.column("subject",anchor=W,width=320)
            my_table.column("recieved",anchor=CENTER,width=100)
            my_table.column("assigned",anchor=CENTER,width=70)
            my_table.column("mailbox",anchor=W,width=145)
            my_table.column("status",anchor=W,width=70)

            my_table.heading("#0",text="",anchor=CENTER)
            my_table.heading("messageID",text="",anchor=CENTER)
            my_table.heading("mailstr",text="",anchor=CENTER)
            my_table.heading("from",text="From",anchor=CENTER)
            my_table.heading("subject",text="Subject",anchor=CENTER)
            my_table.heading("recieved",text="Recieved",anchor=CENTER)
            my_table.heading("assigned",text="Assigned To",anchor=CENTER)
            my_table.heading("mailbox",text="Mailbox",anchor=CENTER)
            my_table.heading("status",text="Status",anchor=CENTER)
            db.execute("select * from mailRecords where mailBox= ? and status not in ('Completed','No Action Required') order by reievedAt ASC", (str(mail.get())))
            for row in db.fetchall(): 
                if(row[7]=="Hold"):    
                    my_table.insert(parent='',index='end',iid=row[0],text='',
                    values=(row[0],row[9],row[2],row[3],row[4],row[8],row[5],row[7]), tags = ("hold",) )
                else:
                    my_table.insert(parent='',index='end',iid=row[0],text='',
                    values=(row[0],row[9],row[2],row[3],row[4],row[8],row[5],row[7]))   
            my_table.tag_configure("hold", background="#fcffbd")
            my_table.place(x=0,y=210)
            getInfo()
        except Exception as e:
            print(e,'1')
            Label(prodwin, text="Selected Mailbox not found, Try again!", font="Helvetica 10 bold").place(x=10,y=210)
            
        tree_scroll.config(command=my_table.yview)
    
    
    
    #importing mails from outlook and inserting into database
    def tablet():
        try:
            lobWt(mail.get())
            started_title.config(text='')
            startedAt.config(text='')
            startedAt.config(bg="#b5edff")
            started_title.config(bg="#b5edff")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            folder = outlook.Folders(mail.get())
            inbox = folder.Folders("Inbox")
            print(inbox.Folders[0])
            print(inbox.Folders[1])
            print(inbox.Folders[2])
            print(inbox.Folders[3])
            print(inbox.Folders[4])
            msg = inbox.Items
            
            for mail_item in msg:
                pathname = os.path.abspath(os.path.dirname(__file__))
                pathname = pathname + "\\Temp\\"
                msgname = str(mail_item.Subject)
                if(len(msgname)>150):
                    msgname = msgname[0:149]
                msgnamestr = re.sub('[^A-Za-z0-9]+', '', msgname)
                msgname = msgnamestr + '.msg'
                
                if not(os.path.exists(pathname + msgname)):
                    mail_item.SaveAs(pathname + msgname)
                try:
                    #entryId = str(mail_item.EntryID)
                    entId = str(mail_item.ReceivedTime.strftime("%H%M%S%f%Y%m%d")) + str(mail_item.ReceivedTime.strftime("%Y%m%d%H%M%S%f"))
                    
                    db.execute("select * from mailRecords where entryId = ? ", (entId))
                    if(db.fetchone()):
                        pass
                    else:
                        fromMail = str(mail_item.Sender)
                        subject = str(mail_item.Subject)
                        reievedAt = str(mail_item.ReceivedTime.strftime("%d/%m/%Y %H:%M"))
                        mailBox = str(mail.get())
                        msgname = str(msgnamestr)
                        db.execute("INSERT INTO mailRecords (entryId, fromMail, subject, reievedAt, mailBox, folder, status, msgnamestr, lastsaved) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)", (entId, fromMail, subject, reievedAt, mailBox, 'Inbox', 'Pending', msgname, reievedAt))
                except Exception as e:
                    print(e,'3')
            softref()
            getInfo()
            
        except Exception as e:
            started_title.config(text='Mailbox!')
            startedAt.config(text='Select')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")
            print(e,'3')
    
    
    
    #defining actions after selecting the mail
    #Setting n Number when mail is selected
    def set_cat(mailID, todo):
        try:
            if(todo =="unassign"):
                db.execute("UPDATE mailRecords SET assignedTo = NULL WHERE id = ? ",(mailID))
            else:
                db.execute("select name from users where nid = ?", (os.getenv('username')))
                user = db.fetchone()[0]
                db.execute("UPDATE mailRecords SET assignedTo = ? WHERE id = ? ",(user,mailID))
        except Exception as e:
            print(e," 4")
    
    def checkassign(assigned):
        if(assigned==None):
            return True
        
        db.execute("select name from users where nid = ?", (os.getenv('username')))
        user = db.fetchone()[0]
        if(user==assigned):
            return True
        else:
            return False
    
    #Opening the mail and starting the timer
    def load_mail():
        try:
            global my_table
            selected_data = my_table.selection()[0]
            data_value = my_table.item(selected_data,'values')
            
            db.execute("select assignedTo from mailRecords where id = ? ",(data_value[0]))
            if(checkassign(db.fetchone()[0])):
                pass
            else:
                started_title.config(text="Processing")
                startedAt.config(text="Already")
                startedAt.config(bg="#ffa759")
                started_title.config(bg="#ffa759")
                return None
            
            mailpathname = os.path.abspath(os.path.dirname(__file__))
            mailpathname = mailpathname + "\\Temp\\"
            os.startfile( mailpathname + data_value[1]+ '.msg')
            b1["state"] = "normal"
            b2["state"] = "disabled"
            b3["state"] = "disabled"
            global startTimeObj
            startTimeObj = datetime.now()
            timestampStr = startTimeObj.strftime("%H:%M:%S")
            started_title.config(text=timestampStr)
            startedAt.config(text="Started at:")
            startedAt.config(bg="#ffd000")
            started_title.config(bg="#ffd000")
            set_cat(data_value[0], "assign")
            policy_value.delete(0,END)
            policy_value.insert(0,'-')
            ucr_value.delete(0,END)
            ucr_value.insert(0,'-')
            trans_value.delete(0,END)
            trans_value.insert(0,'-')
            db.execute("update users set lastactive = ? where nid = ?",(datetime.now(),os.getenv('username')))
            
        except Exception as e:
            print(e, '3')
            #Label(prodwin, text="Please select the mail first!", font="Helvetica 10 bold").place(x=10,y=210)
            b2["state"] = "normal"
            b1["state"] = "disabled"
            b3["state"] = "normal"
            started_title.config(text='mail first!')
            startedAt.config(text='Select')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")

    #Clearing fields
    def clear_field():
        policy_value.delete(0, 'end')
        claim_value.delete(0, 'end')
        ucr_value.delete(0, 'end')
        trans_value.delete(0, 'end')
        comment_value.delete(0, 'end')
        wt.config(text='Select Worktype')
        lob_label.config(text='Select LOB')
        status_label.config(text='Status')
    
    
    
    #moving mails to completed after completion
    def myfunc(mailID, status):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folder = outlook.Folders(mail.get())
        inbox = folder.Folders("Inbox")
        completed = inbox.Folders("Completed")
        msg = inbox.Items
        
        db.execute("select subject from mailRecords where id = ? ", (mailID))
        mailID = db.fetchone()[0]
        for mails in msg:
            if(mails.Subject == mailID):
                mails.Move(completed)
                
    def delMail(mailID):
        db.execute('select msgnamestr from mailRecords where id=?',(mailID))
        mailname = db.fetchone()[0]
        mailname = mailname + '.msg'
        mailpname = os.path.abspath(os.path.dirname(__file__))
        mailpname = mailpname + "\\Temp\\"
        if(os.path.exists(mailpname + mailname)):
            os.remove(mailpname+mailname)
        else:
            print("The file does not exist")
    #Submit the response
    def stop_submit():
        b2["state"] = "normal"
        b1["state"] = "disabled"
        b3["state"] = "normal"
        started_title.config(text='Completed')
        startedAt.config(text='Action')
        startedAt.config(bg="#97fc90")
        started_title.config(bg="#97fc90")        
        try:
            selected_data = my_table.selection()[0]
            data_value = my_table.item(selected_data,'values')
            set_cat(data_value[0], "unassign")
            mailID = data_value[0]
            if(status.get()=="Completed" or status.get()=="No Action Required"):
                delMail(mailID)
            global startTimeObj
            endTimeObj = datetime.now()
            timestampstart = startTimeObj.strftime("%Y-%m-%d %H:%M:%S")
            timestampEnd = endTimeObj.strftime("%Y-%m-%d %H:%M:%S")
            timestampEnd = str(timestampEnd)
            timestampstart = str(timestampstart)
            totalTime = endTimeObj - startTimeObj
            totalTime = str(totalTime)
            recievedAt = data_value[4]
            recievedAt = datetime.strptime(recievedAt,'%Y-%m-%d %H:%M:%S' )
            tat = startTimeObj - recievedAt
            tat = str(tat)
            lb = str(lob.get())
            wrktpe = str(worktype.get())
            plcy = str(policy_value.get())
            clm = str(claim_value.get())
            ucr = str(ucr_value.get())
            trns = str(trans_value.get())
            cmnt = str(comment_value.get())
            stts = str(status.get())
            
            
            db.execute("INSERT INTO prodRecords (mailID, status, startedAt, endedAt, timeTaken, lob, worktype, policyNo, claimNo, ucr, trans, comment, user, tat) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (mailID, stts, timestampstart, timestampEnd, totalTime, lb, wrktpe, plcy, clm, ucr, trns, cmnt, os.getenv('username'),tat))
            db.execute("UPDATE mailRecords SET status = ?, lastsaved=?,worktype = ?, policyNo = ?, claimNo = ?, ucr = ?, trans = ?, comment = ?, user = ?, tat = ?,lob = ? WHERE id = ? ", (stts,timestampEnd,wrktpe, plcy, clm, ucr, trns, cmnt, os.getenv('username'),tat, lb, mailID))
            
            softref()
            db.execute("update users set lastactive = ? where nid = ?",(datetime.now(),os.getenv('username')))
            clear_field()
            lobWt(mail.get())
            if(status.get()=="Completed" or status.get()=="No Action Required"):
                myfunc(mailID, stts)

        except Exception as e:
            softref()
            clear_field()
            lobWt(mail.get())
            started_title.config(text='Details')
            startedAt.config(text='Incomplete')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")
            print(e, 10)
        
    
    
    
        
    
    
    #Elements of production window
    mail_box=Label(frame,text="Select Mailbox",font = "Helvetica 10 bold",bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    mail_box.place(x=40, y=15, height=50)
    
    
    
    #lob dropdown
    lob_label=Label(frame1,text="Select LOB",font = "Helvetica 8 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    lob_label.place(x=35, y=5,height=40)
    options3=[]
    db.execute('select distinct lob from maillob order by lob')
    for row in db.fetchall():
        options3.append(row[0])
    OptionMenu(frame1, lob, *(options3),command=OptionMenu_Select3).place(x=5, y=5,  width=30, height=40)    
    
    
    #select mail dropdown
    options1 = []
    db.execute('select shortmail from maillob order by shortmail')
    for row in db.fetchall():
        options1.append(row[0])
    OptionMenu(frame, mail,*(options1),command=OptionMenu_Select1).place(x=10, y=15, width=30, height=50)

    #refresh and initiate buttons
    b3=Button(frame, text="Refresh",font = "Helvetica 10 bold", command=tablet)
    b3.place(x=10,y=130,width=100,height=50)
    
    #initiateButton = Button(frame, text="Initiate selected mail",font = "Helvetica 10 bold", command=load_mail)
    b2 = Button(frame, text="Initiate selected mail",font = "Helvetica 10 bold", command=load_mail)
    b2.place(x=120,y=130,width=150,height=50)
    
    started_title = Label(frame,text = "",font = "Helvetica 10 bold",fg='#1a1446', bg="#b5edff")
    started_title.place(x=10,y=100, width=100)
    startedAt = Label(frame,text = "",font = "Helvetica 10 bold", fg='#1a1446',bg="#b5edff")
    startedAt.place(x=10,y=80, width=100)
    
    b1 = Button(frame, text="Stop",font = "Helvetica 10 bold", state="disabled", command=stop_submit)
    #stopButton = Button(frame, text="Stop",font = "Helvetica 10 bold")
    b1.place(x=120,y=80,width=150,height=40)
    
    pendingLb = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    pendingLb.place(x=5,y=185)
    pendingVl = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    pendingVl.place(x=65,y=185)
    holdLb = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    holdLb.place(x=110,y=185)
    holdVl = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    holdVl.place(x=150,y=185)
    queryLb = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    queryLb.place(x=190,y=185)
    queryVl = Label(frame, text="",bg="#b5edff",font = "Helvetica 8 bold" )
    queryVl.place(x=250,y=185)
    
   

    


    #username
    UserName = Label(frame1, text=loggedUser,font = "Helvetica 8 bold", bg="#d1e1ff")
    UserName.place(x=429,y=0, width=137)
    
    #Policy
    Label(frame1,text="Policy Number",font ="Helvetica 8 bold", bg="#d1e1ff").place(x=5,y=99)
    policy_value = Entry(frame1)
    
    policy_value.place(x=5,y=119,width=170,height=30)

    #Claim
    Label(frame1,text="Claim Number",font ="Helvetica 8 bold", bg="#d1e1ff").place(x=5,y=149)
    claim_value = Entry(frame1)
    claim_value.place(x=5,y=169,width=170,height=30)
    
    #UCR
    Label(frame1,text="UCR",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=220,y=0)
    ucr_value = Entry(frame1)
    ucr_value.place(x=220,y=20,width=170,height=30)

    #Transaction Number
    Label(frame1,text="#Transaction(UCR cases)",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=220,y=60)
    trans_value = Entry(frame1)
    trans_value.place(x=220,y=85,width=170,height=30)

    #comments
    Label(frame1,text="Comment",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=220,y=118)
    comment_value = Entry(frame1)
    comment_value.place(x=220,y=140,width=340,height=60)

    #status
    status_label=Label(frame1,text="Status",font = "Roboto 8 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    status_label.place(x=450, y=80, height=40)
    options4=[]
    db.execute('select distinct status from statusList')
    for row in db.fetchall():
        options4.append(row[0])
    
    #options3 = ['ACT BDX','Advise/Update','APC BDX','AWAC Bdx','Claim Closed','Claim Open','COLLEGIATE BDX','DACB Report (New Claims, Reserve Movement and Closure)','Daily Payment Report & Error Report','Decision tree','DWF Report (New Claims, Reserve Movement and Closure)','Escrow Top Up','Gibbs Hartley','GTS Centralised Mailbox','Incomplete Items reporting - 004','IRIS BDX','Loss Runs','MGA Casualty','MGA Property ','MID Loss Runs','MILAN BDX','Monthly Bordereaux (MGA/Motor/Asurion/IRIS)','MOTOL XL BDX','New Claim Setup','Pantaenius - Quarterly Reserve Report','Pantaenius - Settlement/Payments Statements','POINEER BDX','Policy Requests - Documents','Policy Requests - Information','PURGE Report','Recovery','Referrals request','Refund','Reserve Update','RFIB BDX','RI LLR Report ','RPC BDX','Settlement - Fees ','Settlement - Indemnity','Settlement - Queries','TPA Collegiate','TPA Davies','Uploads','Volante BDX']
    status_drop = OptionMenu(frame1, status, *(options4),command=OptionMenu_Select4).place(x=420, y=80,  width=30, height=40)

    #Label(frame1,text="Comments",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=400,y=60)
    #status_value = Label(frame1,text="-",font="Helvetica 10 bold", bg="white")
    #status_value = Entry(prodwin)
    #status_value.place(x=400,y=85,width=155,height=95)
    #status_value.insert(END, 'default text')

    #tat
    
    #tat_label=Label(frame1,text="TAT",font = "Helvetica 10 bold", bg="#d1e1ff")
    #tat_label.place(x=420, y=30)
    #tat_value = Label(frame1,text="",font="Helvetica 10 bold", bg="red")
    #tat_value.place(x=470,y=25,width=50,height=30)

    #Call for creating production window
    prodwin.mainloop()

def production():
    prodwin = Tk()
    prodwin.title('Production - Claim Automation Tracker')
    prodwin.geometry("1280x720")
    prodwin.state('zoomed')
    
    
    
    #left Frame
    frame = Frame(prodwin,bg="#b5edff", highlightbackground="grey", highlightthickness=2)
    frame.place(x=0,y=0,width=446,height=210)
    #right Frame
    frame1 = Frame(prodwin,bg="#d1e1ff", highlightbackground="grey", highlightthickness=2)
    frame1.place(x=446,y=0,width=918,height=210)
    
    
    
    #defining dropdown events
    def OptionMenu_Select1(event):
        mail_box.config(text=mail.get())
    #def OptionMenu_Select2(event):
    #    wt.config(text=worktype.get())
    def OptionMenu_Select3(event):
        lob_label.config(text=lob.get())
    def OptionMenu_Select4(event):
        status_label.config(text=status.get())
    
    
    # Create the variables
    mail = StringVar()
    
    def mailcallback(*args):
        try:
            folder.set('')
            folderdrop['menu'].delete(0, 'end')
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            mailbox = outlook.Folders(mail.get()).Folders("Inbox")
            mcount = 0
            for sub in mailbox.Folders:
                if(mcount==0):
                    folderdrop['menu'].add_command(label="Inbox", command=tk._setit(folder, "Inbox"))
                mcount = mcount + 1
                data = str(sub)
                folderdrop['menu'].add_command(label=data, command=tk._setit(folder, data))
            folderlabel.config(text="Select Folder")
        except:
            folderlabel.config(text="Mailbox Not Found")
    mail.trace("w", mailcallback)
    
    worktype = tk.StringVar()
    lob = StringVar()
    status = StringVar()
    def callback(*args):
        wt.config(text=worktype.get())
    
    #worktype dropdown
    wt=Label(frame1,text="Select Worktype",font = "Helvetica 8 bold", bg="#f0f0f0")
    wt.place(x=35, y=55, height=40)
    options2=[]
    db.execute('select distinct worktypes from workTypes order by worktypes')
    for row in db.fetchall():
        options2.append(row[0])
    #options2 = ['ACT BDX','Advise/Update','APC BDX','AWAC Bdx','Claim Closed','Claim Open','COLLEGIATE BDX','DACB Report (New Claims, Reserve Movement and Closure)','Daily Payment Report & Error Report','Decision tree','DWF Report (New Claims, Reserve Movement and Closure)','Escrow Top Up','Gibbs Hartley','GTS Centralised Mailbox','Incomplete Items reporting - 004','IRIS BDX','Loss Runs','MGA Casualty','MGA Property ','MID Loss Runs','MILAN BDX','Monthly Bordereaux (MGA/Motor/Asurion/IRIS)','MOTOL XL BDX','New Claim Setup','Pantaenius - Quarterly Reserve Report','Pantaenius - Settlement/Payments Statements','POINEER BDX','Policy Requests - Documents','Policy Requests - Information','PURGE Report','Recovery','Referrals request','Refund','Reserve Update','RFIB BDX','RI LLR Report ','RPC BDX','Settlement - Fees ','Settlement - Indemnity','Settlement - Queries','TPA Collegiate','TPA Davies','Uploads','Volante BDX']
    wtmenu = tk.OptionMenu(frame1, worktype, *(options2))
    wtmenu.place(x=5, y=55,  width=30, height=40)
    worktype.trace("w", callback)
    
    
    def getInfo():
        db.execute("select count(IIF(status = 'Pending', 1, NULL)) as pending, count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query from mailRecords where mailBox = ? and folder=?",(mail.get(),str(folder.get())))
        for data in db.fetchall():
            pendingLb.config(text="Pending: ")
            holdLb.config(text="Hold: ")
            queryLb.config(text="Queried: ")
            pendingVl.config(text=data[0])
            holdVl.config(text=data[1])
            queryVl.config(text=data[2])
            
    def lobWt(mailbox):
    # Reset wt and delete all old options
        db.execute("select lob from maillob where shortmail = ? ",(mailbox))
        lob.set(db.fetchone()[0])
        lob_label.config(text=lob.get())
        
        worktype.set('')
        wtmenu['menu'].delete(0, 'end')
        
        db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailbox))
        for wts in db.fetchall():
            wtmenu['menu'].add_command(label=wts[0], command=tk._setit(worktype, wts[0]))
        wt.config(text="Select Worktype")
    
    #Creating tkinter treeview and scrollbar; Soft refresh function
    def softref():
        try:
            if(mail.get()=="" or folder.get()==""):
                return
            global my_table
            tree_scroll = Scrollbar(prodwin)
            tree_scroll.place(x=1347,y=210,height=485)

            my_table = ttk.Treeview(prodwin,show='headings', height=23, yscrollcommand=tree_scroll.set)
            my_table['columns'] = ('messageID','mailstr','from', 'subject', 'recieved','assigned', 'mailbox', 'status')

            my_table.column("#0", width=0,  stretch=NO)
            my_table.column("messageID", width=0,  stretch=NO)
            my_table.column("mailstr", width=0,  stretch=NO)
            my_table.column("from",anchor=W, width=190)
            my_table.column("subject",anchor=W,width=430)
            my_table.column("recieved",anchor=CENTER,width=200)
            my_table.column("assigned",anchor=CENTER,width=200)
            my_table.column("mailbox",anchor=W,width=203)
            my_table.column("status",anchor=W,width=120)

            my_table.heading("#0",text="",anchor=CENTER)
            my_table.heading("messageID",text="",anchor=CENTER)
            my_table.heading("mailstr",text="",anchor=CENTER)
            my_table.heading("from",text="From",anchor=CENTER)
            my_table.heading("subject",text="Subject",anchor=CENTER)
            my_table.heading("recieved",text="Recieved",anchor=CENTER)
            my_table.heading("assigned",text="Assigned To",anchor=CENTER)
            my_table.heading("mailbox",text="Mailbox",anchor=CENTER)
            my_table.heading("status",text="Status",anchor=CENTER)
            filterdata = '%'+ str(searchBox.get())+ '%'
            if(searchBox.get()==''):
                db.execute("select * from mailRecords where mailBox= ? and folder= ? and status not in ('Completed','No Action Required') order by reievedAt ASC", (str(mail.get()),str(folder.get())))
            else:
                #db.execute("select * from mailRecords where mailBox= ? and subject status not in ('Completed','No Action Required') order by reievedAt ASC", (str(mail.get())))
                db.execute("select * from mailRecords where mailBox= ? and folder= ? and subject like ? and status not in ('Completed','No Action Required') order by reievedAt ASC", (str(mail.get()),str(folder.get()),filterdata))
            for row in db.fetchall():
                
                dateobj = datetime.strptime(str(row[4]),'%Y-%m-%d %H:%M:%S')
                dateob = dateobj.strftime('%d-%b-%Y %H:%M:%S')
                if(row[7]=="Hold"):    
                    my_table.insert(parent='',index='end',iid=row[0],text='',
                    values=(row[0],row[9],row[2],row[3],dateob,row[8],row[5],row[7]), tags = ("hold",) )
                else:
                    my_table.insert(parent='',index='end',iid=row[0],text='',
                    values=(row[0],row[9],row[2],row[3],dateob,row[8],row[5],row[7]))   
            my_table.tag_configure("hold", background="#fcffbd")
            my_table.place(x=0,y=210)
            getInfo()
            lobWt(mail.get())
            amt_label.config(fg="black")
            amt_label.config(bg="#d1e1ff")
            
        except Exception as e:
            print(e,'1')
            Label(prodwin, text="Selected Mailbox not found, Try again!", font="Helvetica 10 bold").place(x=10,y=210)
            
        tree_scroll.config(command=my_table.yview)
    
    
    
    #importing mails from outlook and inserting into database

    def tablet():
        try:
            
            
            started_title.config(text='')
            startedAt.config(text='')
            startedAt.config(bg="#b5edff")
            started_title.config(bg="#b5edff")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            mailfolder = outlook.Folders(mail.get())
            global folder
            if(folder.get()=="Inbox"):
                inbox = mailfolder.Folders("Inbox")
            else:
                inbox = mailfolder.Folders("Inbox").Folders(folder.get())
            msg = inbox.Items
            
            
            for mail_item in msg:
                
                pathname = os.path.abspath(os.path.dirname(__file__))
                pathname = pathname + "\\Temp\\"
                curdate = datetime.now()
                curdate = curdate.strftime("%f")
                try:
                    msgname = str(mail_item.ReceivedTime.strftime("%H%M%S%Y%m%d"))+ str(curdate)+str(mail_item.Subject)
                except:
                    msgname = str(curdate)+str(mail_item.Subject)
                
                if(len(msgname)>110):
                    msgname = msgname[0:109]
                msgnamestr = re.sub('[^A-Za-z0-9]+', '', msgname)
                msgname = msgnamestr + '.msg'
                
                if not(os.path.exists(pathname + msgname)):

                    mail_item.SaveAs(pathname + msgname)
                try:
                    #entryId = str(mail_item.EntryID)
                    try:
                        entId = str(mail_item.ReceivedTime.strftime("%H%M%S%f%Y%m%d")) + str(mail_item.ReceivedTime.strftime("%Y%m%d%H%M%S%f"))
                    except:
                        entId = str(mail_item.ReceivedTime.strftime("%H%M%S%Y%m%d")) + str(mail_item.ReceivedTime.strftime("%Y%m%d%H%M%S"))
                    
                    db.execute("select * from mailRecords where entryId = ? ", (entId))
                    if(db.fetchone()):
                        pass
                    else:
                        fromMail = str(mail_item.Sender)
                        subject = str(mail_item.Subject)
                        reievedAt = str(mail_item.ReceivedTime.strftime("%d/%m/%Y %H:%M"))
                        mailBox = str(mail.get())
                        msgname = str(msgnamestr)
                        db.execute("INSERT INTO mailRecords (entryId, fromMail, subject, reievedAt, mailBox, folder, status, msgnamestr, lastsaved,tat,user,qcstatus) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?,?,?,?)", (entId, fromMail, subject, reievedAt, mailBox, str(folder.get()), 'Pending', msgname, reievedAt,'--',loggedUser,"Pending"))
                except Exception as e:
                    print(e,'3')
                    print(traceback.format_exc())
            softref()
            getInfo()
            
        except Exception as e:
            started_title.config(text='Mailbox!')
            startedAt.config(text='Select')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")
            print(e,'3')
            print(traceback.format_exc())
    
    #defining actions after selecting the mail
    #Setting n Number when mail is selected
    def set_cat(mailID, todo):
        try:
            if(todo =="unassign"):
                db.execute("UPDATE mailRecords SET assignedTo = NULL WHERE id = ? ",(mailID))
            else:
                db.execute("select name from users where nid = ?", (os.getenv('username')))
                user = db.fetchone()[0]
                db.execute("UPDATE mailRecords SET assignedTo = ? WHERE id = ? ",(user,mailID))
        except Exception as e:
            print(e," 4")
            print(traceback.format_exc())
    
    def checkassign(assigned):
        if(assigned==None):
            return True
        
        db.execute("select name from users where nid = ?", (os.getenv('username')))
        user = db.fetchone()[0]
        if(user==assigned):
            return True
        else:
            return False
    
    #Opening the mail and starting the timer
    def load_mail():
        try:
            global my_table
            selected_data = my_table.selection()[0]
            data_value = my_table.item(selected_data,'values')
            
            db.execute("select assignedTo from mailRecords where id = ? ",(data_value[0]))
            if(checkassign(db.fetchone()[0])):
                pass
            else:
                started_title.config(text="Processing")
                startedAt.config(text="Already")
                startedAt.config(bg="#ffa759")
                started_title.config(bg="#ffa759")
                return None
            
            mailpathname = os.path.abspath(os.path.dirname(__file__))
            mailpathname = mailpathname + "\\Temp\\"
            os.startfile( mailpathname + data_value[1]+ '.msg')
            b1["state"] = "normal"
            b2["state"] = "disabled"
            b3["state"] = "disabled"
            global startTimeObj
            startTimeObj = datetime.now()
            timestampStr = startTimeObj.strftime("%H:%M:%S")
            started_title.config(text=timestampStr)
            startedAt.config(text="Started at:")
            startedAt.config(bg="#ffd000")
            started_title.config(bg="#ffd000")
            set_cat(data_value[0], "assign")
            policy_value.delete(0,END)
            policy_value.insert(0,'-')
            ucr_value.delete(0,END)
            ucr_value.insert(0,'-')
            trans_value.delete(0,END)
            trans_value.insert(0,'-')
            try:
                db.execute("update users set lastactive = ? where nid = ?",(datetime.now(),os.getenv('username')))
            except:
                pass
            
        except Exception as e:
            print(e, '3')
            #Label(prodwin, text="Please select the mail first!", font="Helvetica 10 bold").place(x=10,y=210)
            b2["state"] = "normal"
            b1["state"] = "disabled"
            b3["state"] = "normal"
            started_title.config(text='mail first!')
            startedAt.config(text='Select')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")
            print(traceback.format_exc())

    #Clearing fields
    def clear_field():
        policy_value.delete(0, 'end')
        claim_value.delete(0, 'end')
        ucr_value.delete(0, 'end')
        trans_value.delete(0, 'end')
        comment_value.delete(0, 'end')
        wt.config(text='Select Worktype')
        lob_label.config(text='Select LOB')
        status_label.config(text='Status')
    
    
    
    #moving mails to completed after completion
    def myfunc(mailID, status):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mailfolder = outlook.Folders(mail.get())
        
        if(folder.get()=="Inbox"):
            inbox = mailfolder.Folders("Inbox")
        else:
            inbox = mailfolder.Folders("Inbox").Folders(folder.get())
        
        
        # inbox = folder.Folders("Inbox")
        completed = mailfolder.Folders("Inbox").Folders("Completed")
        msg = inbox.Items
        
        db.execute("select entryId,subject from mailRecords where id = ? ", (mailID))
        data = db.fetchone()
        rdate = data[0]
        if(len(rdate)==40):
            rdate = rdate[0:20]
        else:
            rdate = rdate[0:14]
        mailID = data[1]

        for mails in msg:
            try:
                recieved = str(mails.ReceivedTime.strftime("%H%M%S%f%Y%m%d"))
                
            except:
                recieved = str(mails.ReceivedTime.strftime("%H%M%S%Y%m%d"))
            if(mails.Subject == mailID and recieved==rdate):
                mails.Move(completed)
                break
                
    def delMail(mailID):
        try:
            db.execute('select msgnamestr from mailRecords where id=?',(mailID))
            mailname = db.fetchone()[0]
            mailname = mailname + '.msg'
            mailpname = os.path.abspath(os.path.dirname(__file__))
            mailpname = mailpname + "\\Temp\\"
            if(os.path.exists(mailpname + mailname)):
                os.remove(mailpname+mailname)
            else:
                print("The file does not exist")
        except:
            pass
    #Submit the response
    def stop_submit():
        
        b2["state"] = "normal"
        b1["state"] = "disabled"
        b3["state"] = "normal"
        # if not (trans_value.get()=="-" or trans_value.get()==""):
        #     if not (trans_value.get().isdigit()):
        #         softref()
        #         clear_field()
        #         lobWt(mail.get())
        #         amt_label.config(fg="black")
        #         amt_label.config(bg="red")
        #         started_title.config(text='')
        #         startedAt.config(text='')
        #         startedAt.config(bg="#b5edff")
        #         started_title.config(bg="#b5edff") 
        #         return
        started_title.config(text='Completed')
        startedAt.config(text='Action')
        startedAt.config(bg="#97fc90")
        started_title.config(bg="#97fc90")        
        try:
            selected_data = my_table.selection()[0]
            data_value = my_table.item(selected_data,'values')
            set_cat(data_value[0], "unassign")
            mailID = data_value[0]
            
            global startTimeObj
            endTimeObj = datetime.now()
            timestampstart = startTimeObj.strftime("%d-%b-%Y %H:%M:%S")
            timestampEnd = endTimeObj.strftime("%Y-%m-%d %H:%M:%S")
            timestampEnd = str(timestampEnd)
            timestampstart = str(timestampstart)
            totalTime = endTimeObj - startTimeObj
            totalTime = str(totalTime)
            recievedAt = data_value[4]
            recievedAt = datetime.strptime(recievedAt,'%d-%b-%Y %H:%M:%S' )
            tat = startTimeObj - recievedAt
            #tat = tat.strftime('%M')
            tat = tat.total_seconds()
            tat = int(tat / 3600)
            if(tat>23):
                tatval = "Missed"
            else:
                tatval = "Met"
            #tat = str(tat)
            lb = str(lob.get())
            wrktpe = str(worktype.get())
            plcy = str(policy_value.get())
            clm = str(claim_value.get())
            ucr = str(ucr_value.get())
            trns = str(trans_value.get())
            cmnt = str(comment_value.get())
            stts = str(status.get())
            cur_value = str(cur_dd.get())
            
            #db.execute("select name from users where nid = ?",(str(os.getenv('username'))))
            #user = db.fetchone()
            global loggedUser
            db.execute("INSERT INTO prodRecords (mailID, status, startedAt, endedAt, timeTaken, lob, worktype, policyNo, claimNo, ucr, trans, comment, user, tat) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (mailID, stts, timestampstart, timestampEnd, totalTime, lb, wrktpe, plcy, clm, ucr, trns, cmnt, loggedUser,tatval))
            db.execute("UPDATE mailRecords SET status = ?, lastsaved=?,worktype = ?, policyNo = ?, claimNo = ?, ucr = ?, trans = ?, comment = ?, user = ?, tat = ?,lob = ?, curren=? WHERE id = ? ", (stts,timestampEnd,wrktpe, plcy, clm, ucr, trns, cmnt, loggedUser,tatval, lb, cur_value,mailID))
            
            softref()
            try:
                db.execute("update users set lastactive = ? where nid = ?",(datetime.now(),os.getenv('username')))
            except:
                pass
            clear_field()
            lobWt(mail.get())
            if(status.get()=="Completed" or status.get()=="No Action Required"):
                myfunc(mailID, stts)
                delMail(mailID)
            
                

        except Exception as e:
            softref()
            clear_field()
            lobWt(mail.get())
            started_title.config(text='Details')
            startedAt.config(text='Incomplete')
            startedAt.config(bg="#fa9a70")
            started_title.config(bg="#fa9a70")
            print(e, 10)
            print(traceback.format_exc())
    
    
    
        
    
    
    #Elements of production window
    mail_box=Label(frame,text="Select Mailbox",font = "Helvetica 9 bold",bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    mail_box.place(x=50, y=10, height=40)
    
    #inbox folder
    def folderchange(event):
        folderlabel.config(text=folder.get())
    def foldercb(*args):
        folderlabel.config(text=folder.get())
    global folder
    folder = StringVar()
    folderoptions = []
    folderoptions.append("Inbox")
    folderdrop = OptionMenu(frame, folder, *(folderoptions), command=folderchange)
    folderdrop.place(x=250,y=10, width=30, height=40)
    folderlabel = Label(frame, text="Folder",font ="Helvetica 8 bold", fg="#1a1446", bg="#f0f0f0")
    folderlabel.place(x=280,y=10,height=40)
    folder.trace("w",foldercb)
    
    #lob dropdown
    lob_label=Label(frame1,text="Select LOB",font = "Helvetica 8 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    lob_label.place(x=35, y=5,height=40)
    options3=[]
    db.execute('select distinct lob from maillob order by lob')
    for row in db.fetchall():
        options3.append(row[0])
    OptionMenu(frame1, lob, *(options3),command=OptionMenu_Select3).place(x=5, y=5,  width=30, height=40)    
    
    
    #select mail dropdown
    options1 = []
    db.execute('select shortmail from maillob order by shortmail')
    for row in db.fetchall():
        options1.append(row[0])
    OptionMenu(frame, mail, *(options1), command=OptionMenu_Select1).place(x=10, y=10, width=40, height=40)

    #refresh and initiate buttons
    b3=Button(frame, text="Refresh",font = "Helvetica 10 bold", command=tablet)
    b3.place(x=10,y=120,width=100,height=40)
    
    #initiateButton = Button(frame, text="Initiate selected mail",font = "Helvetica 10 bold", command=load_mail)
    b2 = Button(frame, text="Initiate selected mail",font = "Helvetica 10 bold", command=load_mail)
    b2.place(x=120,y=120,width=150,height=40)
    
    started_title = Label(frame,text = "",font = "Helvetica 10 bold",fg='#1a1446', bg="#b5edff")
    started_title.place(x=10,y=85, width=100)
    startedAt = Label(frame,text = "",font = "Helvetica 10 bold", fg='#1a1446',bg="#b5edff")
    startedAt.place(x=10,y=65, width=100)
    
    b1 = Button(frame, text="Stop",font = "Helvetica 10 bold", state="disabled", command=stop_submit)
    #stopButton = Button(frame, text="Stop",font = "Helvetica 10 bold")
    b1.place(x=120,y=65,width=150,height=40)
    
    searchBox = Entry(frame)
    searchBox.place(x=10, y=170, width=325,height=25)
    filterButton = Button(frame,text="Filter", font="Helvetica 9 bold", command=softref)
    filterButton.place(x=345, y=170, width=75,height=25)
    
    
    stbg = Label(frame, text=" ",bg="#002663")
    stbg.place(x=300,y=65,width=120, height=95)
    pendingLb = Label(frame, text="Pending: ",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    pendingLb.place(x=310,y=72)
    pendingVl = Label(frame, text="-",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    pendingVl.place(x=370,y=72)
    holdLb = Label(frame, text="Hold: ",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    holdLb.place(x=310,y=101)
    holdVl = Label(frame, text="-",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    holdVl.place(x=370,y=101)
    queryLb = Label(frame, text="Queried: ",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    queryLb.place(x=310,y=130)
    queryVl = Label(frame, text="-",bg="#002663",fg="white",font = "Helvetica 9 bold" )
    queryVl.place(x=370,y=130)
    
   

    


    #username
    UserName = Label(frame1, text=loggedUser,font = "Helvetica 8 bold", bg="#d1e1ff")
    UserName.place(x=759,y=0, width=137)
    
    #Policy
    Label(frame1,text="Policy Number",font ="Helvetica 8 bold", bg="#d1e1ff").place(x=5,y=99)
    policy_value = Entry(frame1)
    policy_value.place(x=5,y=119,width=250,height=30)

    #Claim
    Label(frame1,text="Claim Number",font ="Helvetica 8 bold", bg="#d1e1ff").place(x=5,y=149)
    claim_value = Entry(frame1)
    claim_value.place(x=5,y=169,width=250,height=30)

    #UCR
    Label(frame1,text="UCR",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=320,y=0)
    ucr_value = Entry(frame1)
    ucr_value.place(x=320,y=20,width=250,height=30)

    #Amount Number
    amt_label = Label(frame1,text="Amount (In Numeric only)",font ="Helvetica 10 bold", bg="#d1e1ff")
    amt_label.place(x=320,y=60)
    cur_list = ('AED  د.إ','AMD Դ','ANG ƒ','ARS $','ATS','AUD','AZN ₼','BBD $','BDT ৳','BEF','BHD ب.د','BND $','BRL R$','BSD $','BZD $','CAD $','CHF','CHP','CLF','CNY ¥','COP $','CRC ₡','CYP','CZK Kč','DEM','DKK kr','DOP $','DZD د.ج','ESP','EUR €','EZP','FIM','FJD $','FRF','GBP £','GRD','GTQ Q','HKD $','HNL L','HRK Kn','HUF Ft','IDR Rp','INR ₨','ISK Kr','ISS','ITL','JOD د.ا','JPY ¥','KES Sh','KHR ៛','KPW ₩','KRW ₩','KWD د.ك','KYD $','KZT 〒','LAK ₭','LUF','LUR','LYD ل.د','MAD د.م.','MMK','MNT ₮','MOP P','MTP','MUR ₨','MXN $','MYR RM','NGN ₦','NIO C$','NLG','NOK kr','NZD $','OMR ﷼','PAB B/.','PEN S/.','PGK K','PHP ₱','PKR ₨','PLN zł','PTE','PYG ₲','QAR ر.ق','RON L','RUB ₽','SAR ر.س','SEK kr','SGD $','SKK','THB ฿','TND د.ت','TRL','TRY ₺','TTD $','TWD NT$','USD $','UYP','VEB','VEF Bs','VND ₫','XCD $','XPF ₣','ZAR R','ZWD Z$')
    currency = StringVar()
    cur_dd = ttk.Combobox(frame1, width = 10, textvariable = currency)

    cur_dd['values'] = cur_list
    cur_dd.place(x=320,y=85,height=30)
    cur_dd.current(94)
    amt_value = StringVar()
    trans_value = Entry(frame1, textvariable=amt_value)
    trans_value.place(x=420,y=85,width=150,height=30)
    #comments
    Label(frame1,text="Comment",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=320,y=118)
    comment_value = Entry(frame1)
    comment_value.place(x=320,y=140,width=340,height=60)

    #status
    status_label=Label(frame1,text="Select Status",font = "Helvetica 9 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
    status_label.place(x=740, y=140, height=60)
    options4=[]
    db.execute('select distinct status from statusList')
    for row in db.fetchall():
        options4.append(row[0])
    #options3 = ['ACT BDX','Advise/Update','APC BDX','AWAC Bdx','Claim Closed','Claim Open','COLLEGIATE BDX','DACB Report (New Claims, Reserve Movement and Closure)','Daily Payment Report & Error Report','Decision tree','DWF Report (New Claims, Reserve Movement and Closure)','Escrow Top Up','Gibbs Hartley','GTS Centralised Mailbox','Incomplete Items reporting - 004','IRIS BDX','Loss Runs','MGA Casualty','MGA Property ','MID Loss Runs','MILAN BDX','Monthly Bordereaux (MGA/Motor/Asurion/IRIS)','MOTOL XL BDX','New Claim Setup','Pantaenius - Quarterly Reserve Report','Pantaenius - Settlement/Payments Statements','POINEER BDX','Policy Requests - Documents','Policy Requests - Information','PURGE Report','Recovery','Referrals request','Refund','Reserve Update','RFIB BDX','RI LLR Report ','RPC BDX','Settlement - Fees ','Settlement - Indemnity','Settlement - Queries','TPA Collegiate','TPA Davies','Uploads','Volante BDX']
    status_drop = OptionMenu(frame1, status, *(options4),command=OptionMenu_Select4).place(x=700, y=140,  width=40, height=60)
    #Label(frame1,text="Comments",font ="Helvetica 10 bold", bg="#d1e1ff").place(x=400,y=60)
    #status_value = Label(frame1,text="-",font="Helvetica 10 bold", bg="white")
    #status_value = Entry(prodwin)
    #status_value.place(x=400,y=85,width=155,height=95)
    #status_value.insert(END, 'default text')

    #tat
    #tat_label=Label(frame1,text="TAT",font = "Helvetica 10 bold", bg="#d1e1ff")
    #tat_label.place(x=420, y=30)
    #tat_value = Label(frame1,text="",font="Helvetica 10 bold", bg="red")
    #tat_value.place(x=470,y=25,width=50,height=30)

    def ecfwin():
        ecfwindow = Tk()
        ecfwindow.title("Submit BDX/ECF Record - Claim Automation Tracker")
        ecfwindow.geometry("430x620")
        frame = Frame(ecfwindow, bg="#c2c2c2", highlightbackground="grey", highlightthickness=2)
        frame.place(x=0,y=0,width=430,height=620)
        
        
        def mail_callback(*args):
            db.execute("select lob from maillob where shortmail = ? ",(mail_ecf.get()))
            lob_ecf.set(db.fetchone()[0])
            lob_label.config(text=lob_ecf.get())
        
            wt_ecf.set('')
            wt_drop['menu'].delete(0, 'end')
        
            db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mail_ecf.get()))
            for wts in db.fetchall():
                wt_drop['menu'].add_command(label=wts[0], command=tk._setit(wt_ecf, wts[0]))
            wt_label.config(text="Select Worktype")
        
        
        #Mailbox dropdown
        def mail_menu(event):
            mail_label.config(text=mail_ecf.get())
        
        mail_ecf = StringVar()
        
        mail_label=Label(frame,text="Select MailBox",font = "Helvetica 9 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
        mail_label.place(x=50, y=80, height=40)
        options4=[]
        db.execute('select distinct shortmail from maillob order by shortmail')
        for row in db.fetchall():
            options4.append(row[0])
        mail_drop = OptionMenu(frame, mail_ecf, *(options4),command=mail_menu).place(x=10, y=80,  width=40, height=40)
        mail_ecf.trace("w", mail_callback)
        
        
        #LOB dropdown
        
        def lob_menu(event):
            lob_label.config(text=lob_ecf.get())
        
        lob_ecf = StringVar()
        
        lob_label=Label(frame,text="Select LOB",font = "Helvetica 9 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
        lob_label.place(x=50, y=140, height=40)
        options1=[]
        db.execute('select distinct lob from maillob order by lob')
        for row in db.fetchall():
            options1.append(row[0])
        lob_drop = OptionMenu(frame, lob_ecf, *(options1),command=lob_menu).place(x=10, y=140,  width=40, height=40)
        
        
        #Worktype dropdown
        
        def wt_menu(*args):
            wt_label.config(text=wt_ecf.get())
        
        wt_ecf = StringVar()
        
        wt_label=Label(frame,text="Select Worktype",font = "Helvetica 9 bold", bg="#f0f0f0",highlightbackground="grey", highlightthickness=2)
        wt_label.place(x=50, y=200, height=40)
        options2=[]
        db.execute('select distinct worktypes from workTypes order by worktypes')
        for row in db.fetchall():
            options2.append(row[0])
        wt_drop = OptionMenu(frame, wt_ecf, *(options2))
        wt_drop.place(x=10, y=200,  width=40, height=40)
        wt_ecf.trace("w", wt_menu)
        
        
        Label(frame,text="Mail Subject", font="Helvetica 10 bold", bg="#c2c2c2").place(x=10,y=280, height=30)
        sub_input = Entry(frame)
        sub_input.place(x=150,y=280, width=250,height=30)
        
        Label(frame,text="UCR", font="Helvetica 10 bold", bg="#c2c2c2").place(x=10,y=330, height=30)
        ucr_input = Entry(frame)
        ucr_input.place(x=150,y=330, width=250,height=30)
        
        Label(frame,text="Policy/Claim No.", font="Helvetica 10 bold", bg="#c2c2c2").place(x=10,y=380, height=30)
        pol_input = Entry(frame)
        pol_input.place(x=150,y=380, width=250,height=30)
        
        Label(frame,text="Comments", font="Helvetica 10 bold", bg="#c2c2c2").place(x=10,y=430, height=30)
        com_input = Entry(frame)
        com_input.place(x=150,y=430, width=250,height=30)
        
        
        messageBox = Label(frame, text="Submit BDX / ECF Record",bg="#000000",fg="#c2c2c2", font="Helvetica 12 bold")
        messageBox.place(x=10,y=25,height=30)
        
        def submit_ecf():
            
            if(mail_ecf.get()=="" or lob_ecf.get()=="" or wt_ecf.get()=="" or sub_input.get()=="" or ucr_input.get()=="" or pol_input.get()=="" or com_input.get()==""):
                messageBox.config(text="Incomplete Details")
                messageBox.config(bg="#fa9a70")
                messageBox.config(fg="#000000")
                
                return
            else:
                try:
                    timestamp = datetime.now()
                    timestamp_str = timestamp.strftime("%Y%m%d%H%M%S")
                
                    entID = "ECF_" + str(timestamp_str)
                    default_ecf = "ECF"
                    timestampEnd = timestamp.strftime("%Y-%m-%d %H:%M:%S")
                    timestampEnd = str(timestampEnd)
                    global loggedUser
                    mailbx = str(mail_ecf.get())
                    lb = str(lob_ecf.get())
                    wrktpe = str(wt_ecf.get())
                    subjct = str(sub_input.get())
                    ucrval = str(ucr_input.get())
                    plcy = str(pol_input.get())
                    comnt = str(com_input.get())

                    
                    db.execute("INSERT INTO mailRecords (entryId, fromMail, subject, reievedAt, mailBox, folder, status, msgnamestr, lastsaved, worktype, policyNo, claimNo, ucr, comment, user, tat, lob, qcstatus) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (entID,default_ecf,subjct,timestampEnd,mailbx,default_ecf,"Completed",default_ecf,timestampEnd,wrktpe,plcy,plcy,ucrval,comnt,loggedUser,"Met",lb,"Pending"))
                    messageBox.config(text="Record Submitted!")
                    messageBox.config(bg="#97fc90")
                    messageBox.config(fg="#000000")
                except:
                    messageBox.config(text="Error: Please try again!")
                    messageBox.config(bg="#fa9a70")
                    messageBox.config(fg="#000000")
        
        Button(frame,text="Submit",font = "Helvetica 9 bold",fg="#ffffff", bg="#002663",command=submit_ecf).place(x=250,y=500,width=150,height=40)
        
        
        ecfwindow.mainloop()
    
    ecf = Button(frame1,text="Submit BDX / ECF", font="Helvetica 10 bold", fg="#ecac00", bg="#002663", command=ecfwin)
    ecf.place(x=720,y=50, width=160,height=50)


    #Call for creating production window
    prodwin.mainloop()


def supervisor():
    superwin = Tk()
    superwin.title("Supervisor - Claim Automation Tracker")
    superwin.geometry("1280x720")
    superwin.state('zoomed')
    
    
    frameref = Frame(superwin, bg="#ffea8c", highlightbackground="grey", highlightthickness=2)
    frame1 = Frame(superwin, bg="#b5edff", highlightbackground="grey", highlightthickness=2)
    frame2a = Frame(superwin, bg="#fff0c7", highlightbackground="grey", highlightthickness=2)
    frame2 = Frame(superwin, bg="#e3fffa", highlightbackground="grey", highlightthickness=2)
    frame3 = Frame(superwin, bg="#b5edff", highlightbackground="grey", highlightthickness=2)
    frame4 = Frame(superwin, bg="#e3fffa", highlightbackground="grey", highlightthickness=2)
    frameref.place(x=0,y=0,width=360,height=80)
    frame1.place(x=0,y=81,width=360,height=264)
    frame2a.place(x=361,y=0,width=450,height=345)
    frame2.place(x=812,y=0,width=551,height=345)
    frame3.place(x=0,y=346,width=360,height=355)
    frame4.place(x=361,y=346,width=1002,height=355)
    fromdate = StringVar()
    todate = StringVar()
    
    
    
    

    
    
    mainfrom = StringVar()
    mainto = StringVar()
    secondfrom = StringVar()
    secondto = StringVar()
    
    def datepick(mod, bname):
        datepicker = Tk()
 
        # Set geometry
        datepicker.title("Select Date")
        datepicker.geometry("300x270")
        
        # Add Calendar
        cal = Calendar(datepicker, selectmode = 'day',date_pattern='dd/mm/yyyy')
        
        cal.pack(pady = 20)
        
        def grad_date():
                bname.config(text=cal.get_date())
                #print(cal.get_date())
                mod.set(cal.get_date())
                
                datepicker.destroy()
        # Add Button and Label
        Button(datepicker, text = "Select Date",
            command = grad_date).pack()
        
        date = Label(datepicker, text = "")
        date.pack()
        
        # Execute Tkinter
        datepicker.mainloop()
    def getMailStatus():
        try:
            mail_scroll = Scrollbar(frame1)
            mail_scroll.place(x=339,y=30,height=226)
            global mailStatusTable
            mailStatusTable = ttk.Treeview(frame1,show='headings', height=9, yscrollcommand=mail_scroll.set)
            mailStatusTable['columns'] = ('ID','mailbox', 'pending','hold', 'query', 'completed','total')

            mailStatusTable.column("#0", width=0,  stretch=NO)
            mailStatusTable.column("ID", width=0,  stretch=NO)
            mailStatusTable.column("mailbox",anchor=W,width=95)
            mailStatusTable.column("pending",anchor=N,width=50)
            mailStatusTable.column("hold",anchor=N,width=35)
            mailStatusTable.column("query",anchor=N,width=45)
            mailStatusTable.column("completed",anchor=N,width=60)
            mailStatusTable.column("total",anchor=N,width=50)

            mailStatusTable.heading("#0",text="",anchor=CENTER)
            mailStatusTable.heading("ID",text="",anchor=CENTER)
            mailStatusTable.heading("mailbox",text="Mailbox",anchor=CENTER)
            mailStatusTable.heading("pending",text="Pending",anchor=CENTER)
            mailStatusTable.heading("hold",text="Hold",anchor=CENTER)
            mailStatusTable.heading("query",text="Queried",anchor=CENTER)
            mailStatusTable.heading("completed",text="Completed",anchor=CENTER)
            mailStatusTable.heading("total",text="Total",anchor=CENTER)
            db.execute('select distinct shortmail from maillob order by shortmail asc')
            increment = 1
            global todatestr
            mailpend = 0
            mailhold = 0
            mailquery = 0
            mailcomp = 0
            mailtot = 0
            global maildatalist
            maildatalist = []
            maildatalist.clear()
            for mb in db.fetchall():
                #querym = "select count(IIF(status = 'Pending', 1, NULL)) as pending, count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query, count(IIF(status = 'Completed', 1, NULL)) as completed from mailRecords where mailBox = ? AND lastsaved >= CDate(?) AND lastsaved <= CDate(?)"
                #db.execute(querym,mb, secondfrom.get(),todatestr)
                #query4 = ("select count(IIF(status = 'Pending', 1, NULL)) as pending, count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query, count(IIF(status = 'Completed', 1, NULL)) as completed from mailRecords"
                #          " WHERE mailBox = '{}' ").format(mb[0])
                temp = []
                #db.execute(query4)
                query4 = ("select count(IIF(status = 'Pending', 1, NULL)) as pending, count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query, count(IIF(status = 'Completed', 1, NULL)) as completed from mailRecords"
                         " WHERE mailBox = '{}' "
                         " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?)").format(mb[0])
                #print(todatestr)
                #print(secondfrom.get())
                
                db.execute(query4, secondfrom.get(), todatestr)
                
                
                dataval = db.fetchone()
                total = 0
                for totvalue in dataval:
                    total = total+totvalue
                mailStatusTable.insert(parent='',index='end',iid=increment,text='', values=(increment,mb[0],dataval[0],dataval[1],dataval[2],dataval[3],total))
                temp.append(mb[0])
                temp.append(dataval[0])
                temp.append(dataval[1])
                temp.append(dataval[2])
                temp.append(dataval[3])
                temp.append(total)
                maildatalist.append(temp)
                mailpend = mailpend + dataval[0]
                mailhold = mailhold + dataval[1]
                mailquery = mailquery + dataval[2]
                mailcomp = mailcomp + dataval[3]
                mailtot = mailtot + total
                increment = increment+1
            temp1 = []
            temp1.append("Total")
            temp1.append(mailpend)
            temp1.append(mailhold)
            temp1.append(mailquery)
            temp1.append(mailcomp)
            temp1.append(mailtot)
            maildatalist.append(temp1)
            Label(frame1,text="Total",font = "Helvetica 8 bold", bg="#b5edff").place(x=5,y=238)
            Label(frame1,text=mailpend,font = "Helvetica 8 bold", bg="#b5edff").place(x=110,y=238)
            Label(frame1,text=mailhold,font = "Helvetica 8 bold", bg="#b5edff").place(x=160,y=238)
            Label(frame1,text=mailquery,font = "Helvetica 8 bold", bg="#b5edff").place(x=200,y=238)
            Label(frame1,text=mailcomp,font = "Helvetica 8 bold", bg="#b5edff").place(x=250,y=238)
            Label(frame1,text=mailtot,font = "Helvetica 8 bold", bg="#b5edff").place(x=300,y=238)
            mailStatusTable.place(x=0,y=30)
            mail_scroll.config(command=mailStatusTable.yview)
            Label(frame1,text="Mailbox-Wise Overview",font = "Helvetica 9 bold", bg="#b5edff").place(x=5,y=2,height=25)
            def exportmail(maillist):
                df = pd.DataFrame (maillist, columns = ['MailBox','Pending','Hold','Query','Completed','Total'])
                today = datetime.now()
                today = today.strftime('%H%M%S%f%Y%m%d')
                global filepath  
                pathnamemailwise = filepath + "\\Output\\Mailbox-wise_"+ str(today)+".csv"
                df.to_csv(r''+ pathnamemailwise, encoding='utf-8-sig',index=False)
                exportMailStatus.config(bg="#44ad3e")
                exportMailStatus.config(text="Exported")
                
            exportMailStatus = Button(frame1,text="Export Table",font = "Helvetica 9 bold", bg="grey",fg="#fbff00",highlightbackground="grey", highlightthickness=2,command=lambda: exportmail(maildatalist)  )
            exportMailStatus.place(x=256,y=2,width=100, height=25)

        except Exception as e:
            print(e)
            print(traceback.format_exc())
    def getUserData():
        try:
            userd_scroll = Scrollbar(frame3)
            userd_scroll.place(x=339,y=35,height=306)
            global mailStatusTable
            userdTable = ttk.Treeview(frame3,show='headings', height=13, yscrollcommand=userd_scroll.set)
            userdTable['columns'] = ('ID','mailbox','hold', 'query', 'completed','nar','total')

            userdTable.column("#0", width=0,  stretch=NO)
            userdTable.column("ID", width=0,  stretch=NO)
            userdTable.column("mailbox",anchor=W,width=125)
            userdTable.column("hold",anchor=N,width=35)
            userdTable.column("query",anchor=N,width=50)
            userdTable.column("completed",anchor=N,width=60)
            userdTable.column("nar",anchor=N,width=30)
            userdTable.column("total",anchor=N,width=35)

            userdTable.heading("#0",text="",anchor=CENTER)
            userdTable.heading("ID",text="",anchor=CENTER)
            userdTable.heading("mailbox",text="User Name",anchor=CENTER)
            userdTable.heading("hold",text="Hold",anchor=CENTER)
            userdTable.heading("query",text="Queried",anchor=CENTER)
            userdTable.heading("completed",text="Completed",anchor=CENTER)
            userdTable.heading("nar",text="NAR",anchor=CENTER)
            userdTable.heading("total",text="Total",anchor=CENTER)
            
            userhold = 0
            userquery = 0
            usercomp = 0
            usernar = 0
            usertot = 0
            global userdatalist
            userdatalist = []
            userdatalist.clear()
            db.execute("select nid,name from users order by name asc")
            increment = 1
            for un in db.fetchall():
                temp = []
                #db.execute("select count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query, count(IIF(status = 'Completed', 1, NULL)) as completed from prodRecords where user = ?",(un[0]))
                query3 = ("select count(IIF(status = 'Hold', 1, NULL)) as hold,count(IIF(status = 'Queried-Onshore', 1, NULL)) as query, count(IIF(status = 'Completed', 1, NULL)) as completed, count(IIF(status = 'No Action Required', 1, NULL)) as nar from mailRecords"
                         " WHERE user = '{}' "
                         " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?)").format(un[1])
                db.execute(query3, secondfrom.get(), todatestr)
                dataval = db.fetchone()
                total = 0
                for totvalue in dataval:
                    total = total+totvalue
                userdTable.insert(parent='',index='end',iid=increment,text='', values=(increment,un[1],dataval[0],dataval[1],dataval[2],dataval[3],total))
                temp.append(un[1])
                temp.append(dataval[0])
                temp.append(dataval[1])
                temp.append(dataval[2])
                temp.append(dataval[3])
                temp.append(total)
                userdatalist.append(temp)
                increment = increment+1
                userhold = userhold + dataval[0]
                userquery = userquery + dataval[1]
                usercomp = usercomp + dataval[2]
                usernar = usernar + dataval[3]
                usertot = usertot + total
            temp1 = []
            temp1.append("Total")
            temp1.append(userhold)
            temp1.append(userquery)
            temp1.append(usercomp)
            temp1.append(usernar)
            temp1.append(usertot)
            userdatalist.append(temp1)    
                
            Label(frame3,text="Total",font = "Helvetica 8 bold", bg="#b5edff").place(x=5,y=325)
            Label(frame3,text=userhold,font = "Helvetica 8 bold", bg="#b5edff").place(x=138,y=325)
            Label(frame3,text=userquery,font = "Helvetica 8 bold", bg="#b5edff").place(x=180,y=325)
            Label(frame3,text=usercomp,font = "Helvetica 8 bold", bg="#b5edff").place(x=237,y=325)
            Label(frame3,text=usernar,font = "Helvetica 8 bold", bg="#b5edff").place(x=280,y=325)
            Label(frame3,text=usertot,font = "Helvetica 8 bold", bg="#b5edff").place(x=317,y=325)
            userdTable.place(x=0,y=35)
            userd_scroll.config(command=userdTable.yview)
            Label(frame3,text="User-Wise Overview",font = "Helvetica 9 bold", bg="#b5edff").place(x=5,y=7,height=25)
            def exportuser(userlist):
                df = pd.DataFrame (userlist, columns = ['MailBox','Hold','Query','Completed','NAR','Total'])
                today = datetime.now()
                today = today.strftime('%H%M%S%f%Y%m%d')
                global filepath
                pathnameuserwise = filepath + "\\Output\\User-wise_"+ str(today)+".csv"
                df.to_csv(r'' + pathnameuserwise, encoding='utf-8-sig',index=False)
                exportuserdStatus.config(bg="#44ad3e")
                exportuserdStatus.config(text="Exported")
            exportuserdStatus = Button(frame3,text="Export Table",font = "Helvetica 9 bold", bg="grey",fg="#fbff00",highlightbackground="grey", highlightthickness=2,command=lambda: exportuser(userdatalist)  )
            exportuserdStatus.place(x=256,y=7,width=100, height=25)        
        except Exception as e:
            print(e)    
    
    def mailDataStatus():
        #Mailbox Dropdown
        def changemailbox(event):
            sel_mail_label.config(text=shortmailbox.get())
        shortmailbox = StringVar()
        shortmailbox.set("All")
        mailoptions = []
        mailoptions.append("All")
        db.execute("select distinct shortmail from maillob order by shortmail asc")
        for shortmail in db.fetchall():
            mailoptions.append(shortmail[0])
        sel_mailbox = OptionMenu(frame2, shortmailbox, *(mailoptions), command=changemailbox)
        sel_mailbox.place(x=5,y=25, width=30)
        Label(frame2, text="Select Mailbox",bg="#e3fffa", font = "Helvetica 10 bold").place(x=5,y=5)
        sel_mail_label = Label(frame2, text="All", font = "Helvetica 10 bold",highlightbackground="grey", highlightthickness=2)
        sel_mail_label.place(x=34,y=25,height=30)
        
        Label(frame2, text="Search",bg="#e3fffa", font = "Helvetica 10 bold").place(x=280,y=5)
        searchbox = Entry(frame2)
        searchbox.place(x=280,y=25, width=220, height=30)
        
        #callback for setting the radio values
        def setData(var,val):
            var.set(val)
        
        #radio buttons for status
        radioval = StringVar()
        radioval.set("all")
        Label(frame2, text="Select Status",bg="#e3fffa", font = "Helvetica 10 bold").place(x=5,y=80)
        R1 = Radiobutton(frame2, text="All",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="all",command= lambda: setData(radioval, "all"))
        R2 = Radiobutton(frame2, text="WIP",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="wip",command= lambda: setData(radioval, "wip"))
        R3 = Radiobutton(frame2, text="Pending",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="Pending",command= lambda: setData(radioval, "Pending"))
        R4 = Radiobutton(frame2, text="Hold",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="Hold",command= lambda: setData(radioval, "Hold"))
        R5 = Radiobutton(frame2, text="Queried",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="Queried-Onshore",command= lambda: setData(radioval, "Queried-Onshore"))
        R6 = Radiobutton(frame2, text="Completed",bg="#e3fffa",font = "Helvetica 10 bold", variable=radioval, value="Completed",command= lambda: setData(radioval, "Completed"))
        R1.place(x=5,y=105)
        R2.place(x=65,y=105)
        R3.place(x=135,y=105)
        R4.place(x=225,y=105)
        R5.place(x=295,y=105)
        R6.place(x=385,y=105)
        R1.select()

        #assign selected user
        assignuser = Button(frame2,text="Assign Selected to User",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2)
        assignuser.place(x=350,y=290)
        
        def changeuser(event):
            sel_user_label.config(text=userval.get())
        userval = StringVar()
        useroptions = []
        db.execute("select name from users where auth='normal' order by name asc")
        for shortname in db.fetchall():
            useroptions.append(shortname[0])
        Label(frame2, text="Assign User",bg="#fff0c7", font = "Helvetica 9 bold").place(x=700,y=0)
        sel_user = OptionMenu(frame2, userval, *(useroptions), command=changeuser)
        sel_user.config(bg="#a6a6a6", fg="black", activebackground="#a6a6a6", activeforeground="black")
        sel_user["menu"].config(bg="#a6a6a6", fg="black", activebackground="#a6a6a6", activeforeground="black")
        sel_user.place(x=350,y=249, width=30)
        sel_user_label = Label(frame2, text="Select User",bg="#a6a6a6", fg="black", font = "Helvetica 9 bold",highlightbackground="grey", highlightthickness=2)
        sel_user_label.place(x=379,y=251,height=27)
        
        #radio buttons for tat
        tatval = StringVar()
        tatval.set("all")
        Label(frame2, text="Select TAT",bg="#e3fffa", font = "Helvetica 10 bold").place(x=5,y=150)
        R7 = Radiobutton(frame2, text="All",bg="#e3fffa",font = "Helvetica 10 bold", variable=tatval, value="all",command= lambda: setData(tatval, "all"))
        R8 = Radiobutton(frame2, text="Met",bg="#e3fffa",font = "Helvetica 10 bold", variable=tatval, value="Met",command= lambda: setData(tatval, "Met"))
        R9 = Radiobutton(frame2, text="Missed",bg="#e3fffa",font = "Helvetica 10 bold", variable=tatval, value="Missed",command= lambda: setData(tatval, "Missed"))
        R7.place(x=5,y=175)
        R8.place(x=65,y=175)
        R9.place(x=135,y=175)
        R7.select()
#frame4 main table
        mlist = []
        slist = []
        tlist = []
        def maintable():
            user_scroll = Scrollbar(frame4)
            user_scroll.place(x=979,y=2,height=345)
            global userStatusTable
            userStatusTable = ttk.Treeview(frame4,show='headings', height=16, yscrollcommand=user_scroll.set)
            userStatusTable['columns'] = ('ID','hash','mailbox','subject','policy', 'claim', 'status','user','comment','tat','last')

            userStatusTable.column("#0", width=0,  stretch=NO)
            userStatusTable.column("ID", width=0,  stretch=NO)
            userStatusTable.column("hash",anchor=W,width=10)
            userStatusTable.column("mailbox",anchor=W,width=100)
            userStatusTable.column("subject",anchor=W,width=180)
            userStatusTable.column("policy",anchor=W,width=100)
            userStatusTable.column("claim",anchor=W,width=100)
            userStatusTable.column("status",anchor=W,width=70)
            userStatusTable.column("user",anchor=W,width=90)
            userStatusTable.column("comment",anchor=W,width=140)
            userStatusTable.column("tat",anchor=W,width=60)
            userStatusTable.column("last",anchor=W,width=120)
            
            userStatusTable.heading("#0",text="",anchor=CENTER)
            userStatusTable.heading("ID",text="",anchor=CENTER)
            userStatusTable.heading("hash",text="#",anchor=W)
            userStatusTable.heading("mailbox",text="Mailbox",anchor=W)
            userStatusTable.heading("subject",text="Subject",anchor=W)
            userStatusTable.heading("policy",text="Policy#",anchor=W)
            userStatusTable.heading("claim",text="Claim#",anchor=W)
            userStatusTable.heading("status",text="Status",anchor=W)
            userStatusTable.heading("user",text="Username",anchor=W)
            userStatusTable.heading("comment",text="Comments",anchor=W)
            userStatusTable.heading("tat",text="TAT",anchor=W)
            userStatusTable.heading("last",text="Last Saved",anchor=W)
            if(searchbox.get()==""):
                mlist.clear()
                slist.clear()
                tlist.clear()
                
                if(shortmailbox.get()=="All"):
                    db.execute("select distinct shortmail from maillob")
                    for i in db.fetchall():
                        mlist.append(i[0])
                else:
                    mlist.append(shortmailbox.get())
                
                if(radioval.get()=="all"):
                    db.execute("select distinct status from statusList")
                    for i in db.fetchall():
                        slist.append(i[0])
                    slist.append("wip")
                    slist.append("Pending")
                else:
                    slist.append(radioval.get())            
                
                if(tatval.get()=="all"):
                    tlist.append("Met")
                    tlist.append("Missed")
                    tlist.append("--")
                    tlist.append("")
                else:
                    tlist.append(tatval.get())
                global todatestr               
                todatestr = secondto.get()
                todateobj = datetime.strptime(todatestr,'%d/%m/%Y')
                todateobj += timedelta(days=1)
                todatestr = datetime.strftime(todateobj, '%d/%m/%Y')
                todatestr = str(todatestr)
                
                
                query = ("select * from mailRecords"
                        " WHERE mailBox IN {} "
                        " AND status IN {} AND tat IN {} AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) order by lastsaved asc").format(tuple(mlist),tuple(slist),tuple(tlist))
                
                db.execute(query, secondfrom.get(), todatestr)
            
            else:
                searchstr = "%" + searchbox.get() + "%"
                
                db.execute("select * from mailRecords where subject = ? or mailBox = ? or assignedTo = ? or policyNo = ? or claimNo = ? or comment = ?", (searchstr,searchstr,searchstr,searchstr,searchstr,searchstr))
                
            
                
            maintabledatalist = []
            maintabledatalist.clear()
            countid = 1
            for records in db.fetchall():
                #print(records[17])
                if(records[17]==None):
                    uservalue = "-"
                else:
                    #db.execute("select name from users where nid=?",(records[17]))
                    uservalue = records[17]
                
                userStatusTable.insert(parent='',index='end',iid=countid,text='', values=(records[0],countid,records[5],records[3],records[12],records[13],records[7],uservalue,records[15],records[18],records[10]))
                
                tempp = []
                tempp.clear()
                tempp.append(countid)
                tempp.append(records[5])
                tempp.append(records[3])
                tempp.append(records[12])
                tempp.append(records[13])
                tempp.append(records[7])
                tempp.append(uservalue)
                tempp.append(records[15])
                tempp.append(records[18])
                tempp.append(records[10])
                maintabledatalist.append(tempp)
                countid = countid + 1
            #userStatusTable.insert(parent='',index='end',iid=1,text='', values=("1","1","EPCClaims","Mail Subject 1","12345","54321","Pending","Kapil Sharma","No Action Required","Missed",'14/06/2022 14:00 pm','3'))
            userStatusTable.place(x=5,y=2)
            user_scroll.config(command=userStatusTable.yview)
            def exportmaintable(maintabledatalist):
                df = pd.DataFrame (maintabledatalist, columns = ['#','Mailbox','Subject','Policy#','claim#','Status','Username','Comments','Tat','Last Saved'])
                today = datetime.now()
                today = today.strftime('%H%M%S%f%Y%m%d') 
                global filepath
                pathnamemain = filepath + "\\Output\\Maintable_"+ str(today)+".csv"
                df.to_csv(r''+ pathnamemain, encoding='utf-8-sig',index=False)
                exportuserdStatus.config(bg="#44ad3e")
                exportuserdStatus.config(text="Exported")
            exportuserdStatus = Button(frame2,text="Export Table",font = "Helvetica 9 bold", bg="grey",fg="#fbff00",highlightbackground="grey", highlightthickness=2,command=lambda: exportmaintable(maintabledatalist)  )
            exportuserdStatus.place(x=170,y=290,width=140, height=40)
        globals()['maintable']=maintable
        maintable()
        
        searchmail = Button(frame2,text="Filter Results",font = "Helvetica 10 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command=maintable)
        searchmail.place(x=5,y=290, width=140, height=40)        
        
    def useractive():
        try:
            usera_scroll = Scrollbar(frame2a)
            usera_scroll.place(x=428,y=35,height=306)
            
            useraTable = ttk.Treeview(frame2a,show='headings', height=14, yscrollcommand=usera_scroll.set)
            useraTable['columns'] = ('ID','user','lasta', 'lastl')
            
            useraTable.column("#0", width=0,  stretch=NO)
            useraTable.column("ID", width=0,  stretch=NO)
            useraTable.column("user",anchor=W,width=145)
            useraTable.column("lasta",anchor=N,width=134)
            useraTable.column("lastl",anchor=N,width=145)
            useraTable.heading("#0",text="",anchor=CENTER)
            useraTable.heading("ID",text="",anchor=CENTER)
            useraTable.heading("user",text="User Name",anchor=CENTER)
            useraTable.heading("lasta",text="Last Active",anchor=CENTER)
            useraTable.heading("lastl",text="Last LoggedIn",anchor=CENTER)
            global useractivelist
            useractivelist = []
            useractivelist.clear()
            db.execute("select name,lastactive,lastloggedin from users order by name asc")
            increment = 1
            for un in db.fetchall():
                temp=[]
                la = un[1]
                ll = un[2]
                if(un[1]==None):
                    la = "--"
                if(un[2]==None):
                    ll = "--"
                useraTable.insert(parent='',index='end',iid=increment,text='', values=(increment,un[0],la,ll))
                temp.append(un[0])
                temp.append(la)
                temp.append(ll)
                useractivelist.append(temp)
                increment = increment+1
            useraTable.place(x=0,y=35)
            usera_scroll.config(command=useraTable.yview)
            Label(frame2a,text="User Activity Overview",font = "Helvetica 9 bold", bg="#fff0c7").place(x=5,y=7,height=25)
            def exportuseractive(useralist):
                df = pd.DataFrame (useralist, columns = ['User Name','Last Active','Last Loggedin'])
                today = datetime.now()
                today = today.strftime('%H%M%S%f%Y%m%d')
                global filepath
                pathnameuseract = filepath + "\\Output\\User-activity_"+ str(today)+".csv"
                df.to_csv(r''+ pathnameuseract, encoding='utf-8-sig',index=False)
                exportuseraStatus.config(bg="#44ad3e")
                exportuseraStatus.config(text="Exported")
            exportuseraStatus = Button(frame2a,text="Export Table",font = "Helvetica 9 bold", bg="grey",fg="#fbff00",highlightbackground="grey", highlightthickness=2,command=lambda: exportuseractive(useractivelist) )
            exportuseraStatus.place(x=345,y=5,width=100, height=25)        
        except Exception as e:
            print(e)    

    #getUserStatus()

    def refreshData():
        getMailStatus()
        getUserData()
        maintable()
        useractive()
    #frame1 Elements
    mainButton = Button(frameref, text="Refresh\nData",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command=refreshData )
    mainButton.place(x=5,y=13, width=90, height = 50)
    #Label(frameref, text="From",font = "Helvetica 9 bold",bg="#ffea8c",fg="#1a1446").place(x=120,y=5, width=90, height = 20)
    #fromd = Button(frameref, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command= lambda: datepick(mainfrom,fromd))
    #fromd.place(x=120,y=25, width=90, height = 30)
    #Label(frameref, text="To",font = "Helvetica 9 bold",bg="#ffea8c",fg="#1a1446").place(x=240,y=5, width=90, height = 20)
    #Label(frameref, text="-",font = "Helvetica 12 bold",bg="#ffea8c",fg="#1a1446").place(x=220,y=28)
    #tod = Button(frameref, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2,command= lambda: datepick(mainto,tod) )
    #tod.place(x=240,y=25, width=90, height = 30)

    
    #userName = Label(frameref,bg="#ffd000", text="Supervisor: "+loggedUser, font = "Helvetica 10 ")
    #userName.place(x=110,y=5, width=245, height = 20)
    
    #db.execute("select count(IIF(auth = 'normal', 1, NULL)) as pending from users")
    #userCount = db.fetchone()[0]
    #userCount = str(userCount)
    #totalUser = Label(frameref,bg="#b5edff", text="Total Users: "+userCount, font = "Helvetica 10")
    #totalUser.place(x=110,y=25, width=245, height = 20)    
    
    #frame2 elements
    Label(frameref, text="From",font = "Helvetica 9 bold",bg="#ffea8c",fg="#1a1446").place(x=120,y=5, width=90, height = 20)
    secfromd = Button(frameref, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command= lambda: datepick(secondfrom,secfromd))
    secfromd.place(x=120,y=25, width=90, height = 30)
    Label(frameref, text="To",font = "Helvetica 9 bold",bg="#ffea8c",fg="#1a1446").place(x=240,y=5, width=90, height = 20)
    Label(frameref, text="-",font = "Helvetica 12 bold",bg="#ffea8c",fg="#1a1446").place(x=220,y=28)
    sectod = Button(frameref, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2,command= lambda: datepick(secondto, sectod) )
    sectod.place(x=240,y=25, width=90, height = 30)
    
    #setting today's date
    today = datetime.now()
    today = today.strftime('%d/%m/%Y')    
    #mainfrom.set(today)
    #mainto.set(today)
    secondfrom.set(today)
    secondto.set(today)
    #fromd.config(text=today)
    #tod.config(text=today)
    secfromd.config(text=today)
    sectod.config(text=today)
    
    mailDataStatus()
    getMailStatus()
    getUserData()
    useractive()
    superwin.mainloop()

##################################!!!!!


def access():
    accwin = Tk()
    accwin.title('Access - Claim Automation Tracker')
    accwin.geometry("850x480")
    
    frame1 = Frame(accwin,bg="#404040", highlightbackground="grey", highlightthickness=2)
    frame1.place(x=0,y=0,width=283,height=240)
    
    frame2 = Frame(accwin,bg="#404040", highlightbackground="grey", highlightthickness=2)
    frame2.place(x=284,y=0,width=283,height=240)
    
    frame3 = Frame(accwin,bg="#404040", highlightbackground="grey", highlightthickness=2)
    frame3.place(x=567,y=0,width=284,height=240)
    
    frame4 = Frame(accwin,bg="#404040", highlightbackground="grey", highlightthickness=2)
    frame4.place(x=0,y=241,width=425,height=240)
    
    frame5 = Frame(accwin,bg="#404040", highlightbackground="grey", highlightthickness=2)
    frame5.place(x=426,y=241,width=425,height=240)
    
    Label(frame1,text="Manage Mailbox",font ="Helvetica 10 bold", bg="#404040", fg="white").place(x=0,y=0)
    Label(frame2,text="Manage Worktypes",font ="Helvetica 10 bold", bg="#404040", fg="white").place(x=0,y=0)
    Label(frame3,text="Manage Status",font ="Helvetica 10 bold", bg="#404040", fg="white").place(x=0,y=0)
    Label(frame4,text="User/Supervisor",font ="Helvetica 10 bold", bg="#404040", fg="white").place(x=0,y=0)
    Label(frame5,text="LOB Mapping",font ="Helvetica 10 bold", bg="#404040", fg="white").place(x=0,y=0)
    
    #Manage Mailbox
    
    def getmails():
        t_scroll = Scrollbar(frame1)
        t_scroll.place(x=264,y=25,height=146)
        global mailtable
        mailtable = ttk.Treeview(frame1,show='headings', height=6, yscrollcommand=t_scroll.set)
        mailtable['columns'] = ('messageID','mailstr','id','mailbox')

        mailtable.column("#0", width=0,  stretch=NO)
        mailtable.column("messageID", width=0,  stretch=NO)
        mailtable.column("mailstr", width=0,  stretch=NO)
        mailtable.column("id",anchor=W,width=30)
        mailtable.column("mailbox",anchor=W,width=228)

        mailtable.heading("#0",text="",anchor=CENTER)
        mailtable.heading("messageID",text="",anchor=CENTER)
        mailtable.heading("mailstr",text="",anchor=CENTER)
        mailtable.heading("id",text="#",anchor=CENTER)
        mailtable.heading("mailbox",text="Mailbox",anchor=CENTER)
        db.execute("select ID,shortmail from maillob")
        count = 0
        for rows in db.fetchall():
            mailtable.insert(parent='',index='end',iid=count,text='', values=(rows[0],"msgnamestr",count+1,rows[1]))
            count+=1
        mailtable.place(x=0,y=22)
        t_scroll.config(command=mailtable.yview)
    
    def addbut():
        try:
            value = data_entry.get()
            if(value==""):
                mailbox_label.config(text="Enter Mailbox")
                mailbox_label.config(bg="#fa9a70")
                return
            db.execute("insert into maillob (shortmail) values (?)", (value))
            mailbox_label.config(text="Added")
            mailbox_label.config(bg="#97fc90")
            getmails()
            data_entry.delete(0,END)
        except Exception as e:
            print(e)
            mailbox_label.config(text="Error!")
            mailbox_label.config(bg="#fa9a70")
    def savbut(idd):
        data = data_entry.get()
        db.execute("update maillob set shortmail = ? where ID = ?", (data, idd))
        getmails()
        delete_button['command'] = delbut
        update_button['command'] = upbut
        data_entry.delete(0,END)
        mailbox_label.config(bg="#97fc90")
        mailbox_label.config(text="Updated")
        data_entry.delete(0,END)
        add_button['state']= "normal"
        update_button.config(text="Update")
        delete_button.config(text="Delete")
        
    def delbut():
        try:
            
            selected_data = mailtable.selection()[0]
            data_value = mailtable.item(selected_data,'values')
            db.execute("delete from maillob where ID = ?", (data_value[0]))
            getmails()
            mailbox_label.config(text="Deleted!")
            mailbox_label.config(bg="#ff4242")
        except Exception as e:
            print(e)
            mailbox_label.config(text="Select First!")
            mailbox_label.config(bg="#fa9a70")
            
    def canbut():
        delete_button['command'] = delbut
        update_button['command'] = upbut
        data_entry.delete(0,END)
        add_button['state']= "normal"
        update_button.config(text="Update")
        delete_button.config(text="Delete")
        
    def upbut():
        try:
            selected_data = mailtable.selection()[0]
            data_value = mailtable.item(selected_data,'values')
            data_entry.delete(0,END)
            data_entry.insert(0,data_value[3])
            update_button.config(text="Save")
            delete_button.config(text="Cancel")
            add_button['state']= "disabled"
            update_button['command'] = lambda: savbut(data_value[0])
            delete_button['command'] = canbut
            
        except Exception as e:
            print(e)
            mailbox_label.config(text="Select First!")
            mailbox_label.config(bg="#fa9a70")
            
    
    getmails() 
    add_button = Button(frame1, text="Add", command=addbut)
    add_button.place(x=10,y=207,height=25,width=80)
    update_button = Button(frame1, text="Update", command=upbut)
    update_button.place(x=100,y=207,height=25,width=80)
    delete_button = Button(frame1, text="Delete", fg="red", command=delbut)
    delete_button.place(x=190,y=207,height=25,width=80)
    mailbox_label = Label(frame1, text='',bg="#404040")
    mailbox_label.place(x=10,y=175,height=25,width=80)
    data_entry = Entry(frame1)
    data_entry.place(x=100,y=175,height=25,width=168)
    
    #Manage Status
    
    def getstatus():
        status_scroll = Scrollbar(frame3)
        status_scroll.place(x=264,y=25,height=146)
        global statustable
        statustable = ttk.Treeview(frame3,show='headings', height=6, yscrollcommand=status_scroll.set)
        statustable['columns'] = ('ID','id','mailbox','comment')

        statustable.column("#0", width=0,  stretch=NO)
        statustable.column("ID", width=0,  stretch=NO)
        statustable.column("id",anchor=W,width=30)
        statustable.column("mailbox",anchor=W,width=154)
        statustable.column("comment",anchor=W,width=74)

        statustable.heading("#0",text="",anchor=CENTER)
        statustable.heading("ID",text="",anchor=CENTER)
        statustable.heading("id",text="#",anchor=CENTER)
        statustable.heading("comment",text="Comment Required",anchor=CENTER)
        db.execute("select * from statusList")
        count = 0
        for rows in db.fetchall():
            statustable.insert(parent='',index='end',iid=count,text='', values=(rows[0],count+1,rows[1],rows[2]))
            count+=1
        statustable.place(x=0,y=22)
        status_scroll.config(command=statustable.yview)
    
    def addstatus():
        try:
            value = status_entry.get()
            commentvalue = comment_entry.get()
            if(value=="" or commentvalue==""):
                status_label.config(text="Missing Data")
                status_label.config(bg="#fa9a70")
                return
            db.execute("insert into statusList (status, comment) values (?, ?)", (value, commentvalue))
            status_label.config(text="Added")
            status_label.config(bg="#97fc90")
            getstatus()
            status_entry.delete(0,END)
            comment_entry.delete(0,END)
        except Exception as e:
            print(e, ": status add")
            status_label.config(text="Error!")
            status_label.config(bg="#fa9a70")
    def savstatus(idd):
        data = status_entry.get()
        commentdata = comment_entry.get()
        db.execute("update statusList set status = ?, comment= ? where ID = ?", (data,commentdata, idd))
        getstatus()
        delete_status['command'] = delstatus
        update_status['command'] = upstatus
        status_label.config(bg="#97fc90")
        status_label.config(text="Updated")
        status_entry.delete(0,END)
        comment_entry.delete(0,END)
        add_status['state']= "normal"
        update_status.config(text="Update")
        delete_status.config(text="Delete")
        
    def delstatus():
        try:
            
            selected_status = statustable.selection()[0]
            status_value = statustable.item(selected_status,'values')
            db.execute("delete from statusList where ID = ?", (status_value[0]))
            getstatus()
            status_label.config(text="Deleted!")
            status_label.config(bg="#ff4242")
        except Exception as e:
            print(e, ": status delete")
            status_label.config(text="Select First!")
            status_label.config(bg="#fa9a70")
            
    def canstatus():
        try:
            
            delete_status['command'] = delstatus
            update_status['command'] = upstatus
            status_entry.delete(0,END)
            add_status['state']= "normal"
            update_status.config(text="Update")
            delete_status.config(text="Delete")
        except Exception as e:
            print(e, ": status cancel")
    def upstatus():
        try:
            selected_data = statustable.selection()[0]
            data_value = statustable.item(selected_data,'values')
            status_entry.delete(0,END)
            status_entry.insert(0,data_value[2])
            comment_entry.delete(0,END)
            comment_entry.insert(0,data_value[3])
            update_status.config(text="Save")
            delete_status.config(text="Cancel")
            add_status['state']= "disabled"
            update_status['command'] = lambda: savstatus(data_value[0])
            delete_status['command'] = canstatus
            
        except Exception as e:
            print(e, ": status update")
            status_label.config(text="Select First!")
            status_label.config(bg="#fa9a70")
            
    
    getstatus() 
    add_status = Button(frame3, text="Add", command=addstatus)
    add_status.place(x=10,y=207,height=25,width=80)
    update_status = Button(frame3, text="Update", command=upstatus)
    update_status.place(x=100,y=207,height=25,width=80)
    delete_status = Button(frame3, text="Delete", fg="red", command=delstatus)
    delete_status.place(x=190,y=207,height=25,width=80)
    status_label = Label(frame3, text='',bg="#404040")
    status_label.place(x=10,y=175,height=25,width=80)
    status_entry = Entry(frame3)
    status_entry.place(x=100,y=175,height=25,width=140)
    comment_entry = Entry(frame3)
    comment_entry.place(x=242,y=175,height=25,width=26)
    
    
    
    accwin.mainloop()


def qualityBak():
    qualwin = Tk()
    qualwin.title('QC - Claim Automation Tracker')
    qualwin.geometry("1280x720")
    qualwin.state('zoomed')
    
    frame1 = Frame(qualwin, bg="#ffffff", highlightbackground="grey", highlightthickness=1)
    frame2 = Frame(qualwin, bg="#ffd000", highlightbackground="grey", highlightthickness=1)
    frame3 = Frame(qualwin, bg="#f2ffff", highlightbackground="grey", highlightthickness=1)
    frame4 = Frame(qualwin, bg="#f2f4f7", highlightbackground="grey", highlightthickness=1)
    frame1.place(x=0,y=0,width=845,height=200)
    frame2.place(x=845,y=0,width=435,height=200)
    frame3.place(x=0,y=200,width=845,height=520)
    frame4.place(x=845,y=200,width=435,height=520)
    
    
    
    def getcount():
        oplist = []
        if(op.get()=="All"):
            db.execute("select name from users")
            for names in db.fetchall():
                oplist.append(names[0])
        else:
            oplist.append(op.get())
        query = ("select id from mailRecords"
                  " WHERE lob = '{}' AND worktype='{}' AND user IN {} "
                  " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) ").format(lob.get(),wt.get(),tuple(oplist))
        db.execute(query, secondfrom.get(), secondto.get())
        count = 0
        for tet in db.fetchall():
            count = count + 1
        reftotal.config(text=count,font="Helvetica 9 bold")

        
        
    def searchR():
        
        
        if(int(range_.get())>100 or int(range_.get())<1):
            print("Please enter range between 1 to 100")
            return
        try:
            oplist = []
            if(op.get()=="All"):
                db.execute("select name from users")
                for names in db.fetchall():
                    oplist.append(names[0])
            else:
                oplist.append(op.get())
            query = ("select id from mailRecords"
                      " WHERE lob = '{}' AND worktype='{}' AND user IN {} "
                      " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) ").format(lob.get(),wt.get(),tuple(oplist))
            db.execute(query, secondfrom.get(), secondto.get())
            count = 0
            for tet in db.fetchall():
                count = count + 1
            range_limit = round(count/100 * int(range_.get()))
            if(range_limit ==0):
                range_limit = 1
        except:
            pass
        
        try:
            qc_scroll = Scrollbar(frame3)
            qc_scroll.place(x=823,y=3,height=466)
            global qctable
            qctable = ttk.Treeview(frame3,show='headings', height=23, yscrollcommand=qc_scroll.set)
            qctable['columns'] = ('ID','claim','claimdate','operator','subject','comments')

            qctable.column("#0", width=0,  stretch=NO)
            qctable.column("ID", width=0,  stretch=NO)
            qctable.column("claim",anchor=W,width=140)
            qctable.column("claimdate",anchor=N,width=130)
            qctable.column("operator",anchor=N,width=124)
            qctable.column("subject",anchor=W,width=238)
            qctable.column("comments",anchor=W,width=184)

            qctable.heading("#0",text="",anchor=CENTER)
            qctable.heading("ID",text="",anchor=CENTER)
            qctable.heading("claim",text="Claim #",anchor=CENTER)
            qctable.heading("claimdate",text="Process Date",anchor=CENTER)
            qctable.heading("operator",text="Processor",anchor=CENTER)
            qctable.heading("subject",text="Mail Subject",anchor=CENTER)
            qctable.heading("comments",text="xx Comments",anchor=CENTER)
        
            query4 = ("select TOP {} id,claimNo,lastsaved,user,subject,comment from mailRecords"
                      " WHERE lob = '{}' AND worktype='{}' AND user IN {} "
                      " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) ORDER BY entryId").format(range_limit,lob.get(),wt.get(),tuple(oplist))
            db.execute(query4, secondfrom.get(), secondto.get())
            idc = 0
            for tdata in db.fetchall():
                qctable.insert(parent='',index='end',iid=idc,text='', values=(tdata[0],tdata[1],tdata[2],tdata[3],tdata[4],tdata[5]))
                idc+=1
            qctable.place(x=3,y=3)
            qc_scroll.config(command=qctable.yview)
            selectRec["state"] = "normal"
        except:
            pass
        
    def setval(var,val):
        var.set(val)
        print(var.get())
        
    def questionList(lob,wt):
        
        db.execute("select ID,question,positive,negetive from qcques where lob = ? AND worktype = ?",(lob,wt))
        varlist = []
        for ss in db.fetchall():
            temp = []
            temp.append(str(ss[0]))
            temp.append(str(ss[1]))
            temp.append(str(ss[2]))
            temp.append(str(ss[3]))
            temp.append("rvar"+ str(ss[0]))
            temp.append("r1"+ str(ss[0]))
            temp.append("r2"+ str(ss[0]))
            varlist.append(temp)
        
        yaxis = 50
        
        global dlist
        global buttonlist
        dlist = []
        buttonlist = []
        dlist.clear()
        buttonlist.clear()
        for qc in varlist:
            templist = []
            question = Label(frame4,text=qc[1], font="Helvetica 9", bg="#f2f4f7")
            question.place(x=5,y=yaxis)
            var = qc[4]
            templist.append(qc[5])
            templist.append(qc[6])
            var = StringVar()
            dlist.append(var)
            var.set(qc[2])
            qc[5] = Radiobutton(frame4, text=qc[2],bg="#f2f4f7",font = "Helvetica 9", variable=var, value=qc[2], command=lambda: setval(var,qc[2]))
            qc[6] = Radiobutton(frame4, text=qc[3],bg="#f2f4f7",font = "Helvetica 9", variable=var, value=qc[3], command=lambda: setval(var,qc[3]))
            qc[5].place(x=310,y=yaxis)
            qc[6].place(x=370,y=yaxis)
            qc[5].select()
            buttonlist.append(templist)
            yaxis = yaxis + 25
        
    
        
        def markall(lob,wt,action):
            pass
        
        Button(frame4, text="Mark All", font = "Helvetica 7",bg="#f2f4f7",command = lambda: markall(lob,wt,1)).place(x=310,y=10,height=25)
        Button(frame4, text="Mark All", font = "Helvetica 7",bg="#f2f4f7",command = lambda: markall(lob,wt,0)).place(x=370,y=10,height=25)
        Label(frame4,text="Auditor Comments",font="Helvetica 10 bold",bg="#f2f4f7").place(x=10,y=452)
        acomments = Entry(frame4)
        acomments.place(x=140,y=450,width=280,height=30)

    
    
    def savecase():
        selectRec.config(text="Select Case")
        selectRec["command"] = selectcase
        searchResult["state"] = "normal"
        claimentry.delete(0,END)
        cdateentry.delete(0,END)
        opentry.delete(0,END)
        subentry.delete(0,END)
        commententry.delete(0,END)
        for widget in frame4.winfo_children():
            widget.destroy()
    
    def selectcase():
        global qctable
        selected_data = qctable.selection()[0]
        data_value = qctable.item(selected_data,'values')
        selectRec.config(text="Save/Stop")
        selectRec["command"] = savecase
        searchResult["state"] = "disabled"
        claimentry.delete(0,END)
        claimentry.insert(0,data_value[1])
        cdateentry.delete(0,END)
        cdateentry.insert(0,data_value[2])
        opentry.delete(0,END)
        opentry.insert(0,data_value[3])
        subentry.delete(0,END)
        subentry.insert(0,data_value[4])
        commententry.delete(0,END)
        commententry.insert(0,data_value[5])

        #call questionList function
        
        questionList(lob.get(),wt.get())






    secondfrom = StringVar()
    secondto = StringVar()
    
    def datepick(mod, bname):
        datepicker = Tk()
 
        # Set geometry
        datepicker.title("Select Date")
        datepicker.geometry("300x270")
        
        # Add Calendar
        cal = Calendar(datepicker, selectmode = 'day',date_pattern='dd/mm/yyyy')
        
        cal.pack(pady = 20)
        
        def grad_date():
                bname.config(text=cal.get_date())
                #print(cal.get_date())
                mod.set(cal.get_date())
                
                datepicker.destroy()
        # Add Button and Label
        Button(datepicker, text = "Select Date",
            command = grad_date).pack()
        
        date = Label(datepicker, text = "")
        date.pack()
        
        # Execute Tkinter
        datepicker.mainloop()
    
    
    #mailbox dropdown
    # def mailchange(event):
    #     maillabel.config(text=mail.get())
    #     lobWt(mail.get())
    # mail = StringVar()
    # mailoptions = []
    # mailoptions.append('All')
    # db.execute('select distinct shortmail from maillob')
    # for mails in db.fetchall():
    #     mailoptions.append(mails[0])
    # drop1 = OptionMenu(frame1, mail, *(mailoptions), command=mailchange)
    # drop1.place(x=20,y=40, width=60, height=40)
    # maillabel = Label(frame1, text="Select Mailbox",font ="Helvetica 10 bold", fg="#1a1446", bg="#d6d6d6")
    # maillabel.place(x=80,y=40,height=40)
    
    #Operator dropdown
    def operatorchange(event):
        oplabel.config(text=op.get())
    op = StringVar()
    opoptions = []
    opoptions.append('All')
    db.execute('select distinct name from users')
    for opname in db.fetchall():
        opoptions.append(opname[0])
    drop2 = OptionMenu(frame1, op, *(opoptions), command=operatorchange)
    drop2.place(x=20,y=40, width=60, height=40)
    oplabel = Label(frame1, text="Select Operator",font ="Helvetica 10 bold", fg="#1a1446", bg="#d6d6d6")
    oplabel.place(x=80,y=40,height=40)
    
    def lobchange(event):
        loblabel.config(text=lob.get())
    
    def lobcallback(*args):
        
        wt.set('')
        drop4['menu'].delete(0, 'end')
        
        if(lob.get()=="All"):
            db.execute("select distinct worktypes from workTypes")
            for wotype in db.fetchall():
                drop4['menu'].add_command(label=wotype[0], command=tk._setit(wt, wotype[0]))
        else:
            thisset = set()
            db.execute("select distinct shortmail from maillob where lob = ? order by shortmail",(lob.get()))
            for mailb in db.fetchall():
                db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailb[0]))
                for wtype in db.fetchall():    
                    thisset.add(wtype[0])
            for woty in thisset:
                
                drop4['menu'].add_command(label=woty, command=tk._setit(wt, woty))
        wtlabel.config(text="Select Worktype")
    
    
    lob = StringVar()
    loboptions = []
    #loboptions.append('All')
    db.execute('select distinct lob from maillob')
    for lobname in db.fetchall():
        loboptions.append(lobname[0])
    drop3 = OptionMenu(frame1, lob, *(loboptions), command=lobchange)
    drop3.place(x=270,y=40, width=60, height=40)
    loblabel = Label(frame1, text="Select LOB",font ="Helvetica 10 bold", fg="#1a1446", bg="#d6d6d6")
    loblabel.place(x=330,y=40,height=40)
    lob.trace("w", lobcallback)
    
    
    def wtcallback(*args):
        wtlabel.config(text=wt.get())
    def wtchange(event):
        wtlabel.config(text=wt.get())
    wt = StringVar()
    wtoptions = []
    wtoptions.append('All')
    db.execute('select distinct worktypes from workTypes')
    for wtname in db.fetchall():
        wtoptions.append(wtname[0])
    drop4 = OptionMenu(frame1, wt, *(wtoptions), command=wtchange)
    drop4.place(x=560,y=40, width=60, height=40)
    wtlabel = Label(frame1, text="Select Worktype",font ="Helvetica 10 bold", fg="#1a1446", bg="#d6d6d6")
    wtlabel.place(x=620,y=40,height=40)
    wt.trace("w", wtcallback)
    
    
    range_ = Entry(frame1, background="#ebebeb", font=('Helvetica 15 bold'))
    range_.place(x=130,y=120,height=40,width=50)
    range_label = Label(frame1, text="Range %(1-100)", font="Helvetica 10 bold", fg="#1a1446")
    range_label.place(x=20,y=120, width=110, height=40)
    
    Label(frame1, text="From",font = "Helvetica 9 bold",bg="#ffffff",fg="#1a1446").place(x=230,y=110, width=90, height = 20)
    secfromd = Button(frame1, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command= lambda: datepick(secondfrom,secfromd))
    secfromd.place(x=230,y=130, width=90, height = 30)
    Label(frame1, text="To",font = "Helvetica 9 bold",bg="#ffffff",fg="#1a1446").place(x=330,y=110, width=90, height = 20)
    sectod = Button(frame1, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2,command= lambda: datepick(secondto, sectod) )
    sectod.place(x=330,y=130, width=90, height = 30)
    
    reftotal = Button(frame1,text="🔁", font="Helvetica 20 bold", command=getcount)
    reftotal.place(x=460,y=120,width=80,height=40)
    searchResult = Button(frame1,text="Search Results", font="Helvetica 10 bold", command=searchR)
    searchResult.place(x=560,y=120,width=140,height=40)
    selectRec = Button(frame1,text="Select Case", font="Helvetica 10 bold", state="disabled", command=selectcase)
    selectRec.place(x=730,y=120,width=90,height=40)
    
    #frame2
    Label(frame2,text="Claim #",bg="#1a1446",fg="white",font="Helvetica 10 bold").place(x=5,y=40,width=130)
    Label(frame2,text="Processed Date",bg="#1a1446",fg="white",font="Helvetica 10 bold").place(x=5,y=70,width=130)
    Label(frame2,text="Processor Name",bg="#1a1446",fg="white",font="Helvetica 10 bold").place(x=5,y=100,width=130)
    Label(frame2,text="Email Subject",bg="#1a1446",fg="white",font="Helvetica 10 bold").place(x=5,y=130,width=130)
    Label(frame2,text="xx Comments",bg="#1a1446",fg="white",font="Helvetica 10 bold").place(x=5,y=160,width=130)
    claimentry = Entry(frame2)
    cdateentry = Entry(frame2)
    opentry = Entry(frame2)
    subentry = Entry(frame2)
    commententry = Entry(frame2)
    claimentry.place(x=135,y=40,width=285, height=22)
    cdateentry.place(x=135,y=70,width=285, height=22)
    opentry.place(x=135,y=100,width=285, height=22)
    subentry.place(x=135,y=130,width=285, height=22)
    commententry.place(x=135,y=160,width=285, height=22)
    
    #loggeduser
    global loggedUser
    Label(frame2,text=loggedUser,bg="#ffd000",fg="#1a1446",font="Helvetica 9 bold").place(x=5,y=2)
    
    # def lobWt(mailbox):
    # # Reset wt and delete all old options
    #     if(mailbox=="All"):
    #         db.execute("select distinct lob from maillob")
    #         loblabel.config(text="Select LOB")
    #     else:
    #         db.execute("select lob from maillob where shortmail = ? ",(mailbox))
    #         lob.set(db.fetchone()[0])
    #         loblabel.config(text=lob.get())
        
    #     wt.set('')
    #     drop4['menu'].delete(0, 'end')
        
    #     if(mailbox=="All"):
    #         db.execute("select distinct worktypes from workTypes")
    #     else:
    #         db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailbox))
    #     for wts in db.fetchall():
    #         drop4['menu'].add_command(label=wts[0], command=tk._setit(wt, wts[0]))
    #     wtlabel.config(text="Select Worktype")
    
    
    
    qualwin.mainloop()


def quality():
    qualwin = Tk()
    qualwin.title('QC - Claim Automation Tracker')
    qualwin.geometry("1280x720")
    qualwin.state('zoomed')
    
    frame1 = Frame(qualwin, bg="#ffffff", highlightbackground="grey", highlightthickness=1)
    frame2 = Frame(qualwin, bg="#ffd000", highlightbackground="grey", highlightthickness=1)
    frame3 = Frame(qualwin, bg="#f2ffff", highlightbackground="grey", highlightthickness=1)
    frame4 = Frame(qualwin, bg="#f2f4f7", highlightbackground="grey", highlightthickness=1)
    frame1.place(x=0,y=0,width=845,height=250)
    frame2.place(x=845,y=0,width=435,height=200)
    frame3.place(x=0,y=250,width=845,height=470)
    frame4.place(x=845,y=200,width=435,height=520)
    
    
    
    def getcount():
        
        oplist = []
        if(op.get()=="All"):
            db.execute("select distinct name from users")
            for names in db.fetchall():
                oplist.append(names[0])
        else:
            oplist.append(op.get())
        query = ("select id from mailRecords"
                  " WHERE lob = '{}' AND worktype='{}' AND user IN {} AND qcstatus = '{}'  AND status IN ('No Action Required','Completed')"
                  " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?)").format(lob.get(),wt.get(),tuple(oplist),qcstat.get())
        
        db.execute(query, secondfrom.get(), secondto.get())
        count = 0
        for tet in db.fetchall():
            count = count + 1
        reftotal.config(text=count,font="Helvetica 9 bold")
    
    global boo
    boo = 0
    def sortcol(event):
        xs = qctable.identify_column(event.x)
        xd = qctable.identify_row(event.y)
        xs = xs[1:]
        if(xd==""):
            if(xs=="3" or xs=="6"):
                global boo
                if(boo==0):
                    boo=1
                else:
                    boo=0
                searchR(int(xs),boo)
            
        
    def searchR(value,boo):
        if(value==3):
            cols = "lastsaved"
        if(value==6):
            cols = "trans"
        if not(value==3 or value==6):
            cols="entryId"
        if(boo==0):
            asc = "ASC"
        if(boo==1):
            asc = "DESC"
        if(qcstat.get()=="Pending"):
            if(range_.get()=='' or int(range_.get())>100 or int(range_.get())<1):
                print("Please enter range between 1 to 100")
                return
        try:
            oplist = []
            if(op.get()=="All"):
                db.execute("select distinct name from users")
                for names in db.fetchall():
                    oplist.append(names[0])
            else:
                oplist.append(op.get())
            query = ("select id from mailRecords"
                      " WHERE lob = '{}' AND worktype='{}' AND user IN {} AND qcstatus = '{}' AND status IN ('No Action Required','Completed')"
                      " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?)").format(lob.get(),wt.get(),tuple(oplist),qcstat.get())
            db.execute(query, secondfrom.get(), secondto.get())
            global range_limit
            count = 0
            for tet in db.fetchall():
                count = count + 1
            if(qcstat.get()=="Pending"):
                range_limit = round(count/100 * int(range_.get()))
            else:
                range_limit = count
            if(range_limit ==0):
                range_limit = 1
                
        except:
            print(traceback.format_exc())
            pass
            
        
        try:
            qc_scroll = Scrollbar(frame3,orient='vertical')
            qc_xscroll = Scrollbar(frame3,orient='horizontal')
            qc_scroll.place(x=823,y=3,height=445)
            qc_xscroll.place(x=0,y=420,width=810)
            qc_scroll = Scrollbar(frame3)
            qc_scroll.place(x=823,y=3,height=466)
            global qctable
            qctable = ttk.Treeview(frame3,show='headings', height=19, yscrollcommand=qc_scroll.set,xscrollcommand=qc_xscroll.set)
            qctable['columns'] = ('ID','claim','claimdate','operator','subject','amount','comments','auditor', 'qcdate','qcscore')

            qctable.column("#0", width=0,  stretch=NO)
            qctable.column("ID", width=0,  stretch=NO)
            qctable.column("claim",anchor=W,width=105)
            qctable.column("claimdate",anchor=N,width=110)
            qctable.column("operator",anchor=N,width=100)
            qctable.column("subject",anchor=W,width=208)
            qctable.column("amount",anchor=W,width=100)
            qctable.column("comments",anchor=W,width=114)
            qctable.column("auditor",anchor=W,width=80)
            qctable.column("qcdate",anchor=W,width=110)
            qctable.column("qcscore",anchor=W,width=40)

            qctable.heading("#0",text="",anchor=CENTER)
            qctable.heading("ID",text="",anchor=CENTER)
            qctable.heading("claim",text="Claim #",anchor=CENTER)
            qctable.heading("claimdate",text="Process Date",anchor=CENTER)
            qctable.heading("operator",text="Processor",anchor=CENTER)
            qctable.heading("subject",text="Mail Subject",anchor=CENTER)
            qctable.heading("amount",text="Amount",anchor=CENTER)
            qctable.heading("comments",text="xx Comments",anchor=CENTER)
            qctable.heading("auditor",text="Auditor",anchor=CENTER)
            qctable.heading("qcdate",text="QC Date",anchor=CENTER)
            qctable.heading("qcscore",text="Score",anchor=CENTER)
            
            qctodate = datetime.strptime(secondto.get(),'%d/%m/%Y')
            qctodate = qctodate + timedelta(days=1)
            
            if(searchQc.get()==''):
                query4 = ("select TOP {} id,claimNo,lastsaved,user,subject,comment,qcdate,qcscore,auditor,trans,curren from mailRecords"
                        " WHERE lob = '{}' AND worktype='{}' AND user IN {} AND qcstatus = '{}' AND status IN ('No Action Required','Completed')"
                        " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) ORDER BY {} {}").format(range_limit,lob.get(),wt.get(),tuple(oplist),qcstat.get(),cols,asc)
                db.execute(query4, secondfrom.get(), qctodate)
            else:
                filterdata = '%'+ str(searchQc.get())+ '%'
                range_limit = count
                query4 = ("select TOP {} id,claimNo,lastsaved,user,subject,comment,qcdate,qcscore,auditor,trans,curren from mailRecords"
                        " WHERE lob = '{}' AND worktype='{}' AND user IN {} AND qcstatus = '{}' AND status IN ('No Action Required','Completed')"
                        " AND lastsaved >= CDate(?) AND lastsaved <= CDate(?) AND claimNo LIKE ? ORDER BY {} {}").format(range_limit,lob.get(),wt.get(),tuple(oplist),qcstat.get(),cols,asc)
                db.execute(query4, secondfrom.get(), qctodate, filterdata)
                
            
            idc = 0
            for tdata in db.fetchall():
                claimdatee = datetime.strptime(str(tdata[2]),'%Y-%m-%d %H:%M:%S')
                claimdatee = claimdatee.strftime('%d-%b-%Y %H:%M:%S')
                qcdatee = tdata[6]
                qcscoree = tdata[7]
                if(qcdatee==None):
                    qcdatee = '--'
                else:
                    qcdatee = datetime.strptime(str(qcdatee),'%Y-%m-%d %H:%M:%S')
                    qcdatee = qcdatee.strftime('%d-%b-%Y %H:%M:%S')
                if(qcscoree==None):
                    qcscoree = '--'
                if(tdata[9]=="-" or tdata[9]==""):
                    amountt = "-"
                else:
                    amountt = str(tdata[10]+" "+ str(tdata[9]))
                qctable.insert(parent='',index='end',iid=idc,text='', values=(tdata[0],tdata[1],claimdatee,tdata[3],tdata[4],amountt,tdata[5],tdata[8],qcdatee,qcscoree))
                qctable.bind("<Double-1>", selectcase)
                idc+=1
            qctable.bind("<Button-1>", sortcol)
            qctable.place(x=3,y=3,width=810)
            qc_scroll.config(command=qctable.yview)
            qc_xscroll.config(command=qctable.xview)
            selectRec["state"] = "normal"
        except:
            print(traceback.format_exc())
            pass
        
        
    def questionList(lob,wt):
        global qctable
        selected_data = qctable.selection()[0]
        data_value = qctable.item(selected_data,'values')
        yaxis = 50
        global variables
        variables = []
        variables.clear()
        global values
        values = []
        values.clear()
        db.execute("select ID,question,positive,negetive,point,isnum from qcques where lob = ? AND worktype = ? order by ID",(lob,wt))
        for idd, ques, pos, neg, point,isnum in db.fetchall():
            Label(frame4,text=ques, font="Helvetica 9", bg="#f2f4f7").place(x=5,y=yaxis)
            Label(frame4,text=point, font="Helvetica 8", bg="#f2f4f7").place(x=310,y=yaxis)
            ent = Entry(frame4)
            ent.place(x=385,y=yaxis, width=35)
            yaxis = yaxis + 22
            variables.append(ent)
            temp = []
            temp.append(data_value[0])
            temp.append(idd)
            temp.append(ques)
            temp.append(pos)
            temp.append(neg)
            temp.append(isnum)
            
            values.append(temp)

        def markall(var):
            for i in range(len(variables)):
                variables[i].delete(0,END)
                if(var==1):
                    variables[i].insert(0,values[i][3])
                else:
                    variables[i].insert(0,values[i][4])
            
        Button(frame4, text="Mark All Postive", font = "Helvetica 7",bg="#f2f4f7",command = lambda: markall(1)).place(x=250,y=4,height=25)
        Button(frame4, text="Mark All Negetive", font = "Helvetica 7",bg="#f2f4f7",command = lambda: markall(0)).place(x=340,y=4,height=25)
        Label(frame4, text="*Responses are Case-Sensitive!", font = "Helvetica 7",bg="#f2f4f7",fg="red").place(x=260,y=30)
        
        
        def change_sendmail(event):
            sendmail_label.config(text=sendmail.get())
            
        global sendmail
        sendmail = StringVar()
        sendmail_op = []
        sendmail_op.append("No")
        sendmail_op.append("Yes")
        sendmail.set("No")
        OptionMenu(frame4, sendmail,*(sendmail_op),command = change_sendmail).place(x=140,y=10,width=30,height=25)
        Label(frame4,text="Send email to Processor?",font="Helvetica 8").place(x=10,y=10,height=25)
        sendmail_label = Label(frame4,text="No",font="Helvetica 8",fg="#1a1446", bg="#d6d6d6")
        sendmail_label.place(x=170,y=10, height=25)
        
        Label(frame4,text="Auditor Comments",font="Helvetica 9 bold",bg="#f2f4f7").place(x=5,y=412)
        global acomments
        acomments = Entry(frame4)
        acomments.place(x=5,y=440,width=320,height=25)
        saveRec = Button(frame4,text="Save/Stop", font="Helvetica 9 bold", command=savecase)
        saveRec.place(x=350,y=440, height=25)

    def aftersave():
        selectRec.config(text="Select Case")
        selectRec["command"] = selectcase
        searchResult["state"] = "normal"
        qcstatdrop2.configure(state="normal")
        claimentry.delete(0,END)
        cdateentry.delete(0,END)
        opentry.delete(0,END)
        subentry.delete(0,END)
        amountentry.delete(0,END)
        commententry.delete(0,END)
        for widget in frame4.winfo_children():
            widget.destroy()
        searchR(3,0)
    
    def sendmail_func(setvalues,values,aucom,score):
        global claim_
        global pdate_
        global pname_
        global mailsub_
        global amount_
        global gencom_
        global loggedUser
        qdata = ""
        for i in range(len(setvalues)):
            tempp = str(values[i][2]) + ": " + str(setvalues[i]) + "\n"
            qdata = qdata + tempp
        maildata = "Claim #: "+str(claim_)+ "\nProcessing Date: " + str(pdate_)+ "\nSubject: " + str(mailsub_)+ "\nAmount: " + str(amount_)+ "\nxx Comments: " + str(gencom_)+ "\n"
        db.execute("select mail from users where name = ? ",(pname_))
        usermail = db.fetchone()
        usermail = usermail[0]
        
        db.execute("select mail from users where cc='yes'")
        cclist = ""
        for cc in db.fetchall():
            cclist = cclist + cc[0]
            cclist = cclist + ";"
        cclist = str(cclist)
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        mail = outlook.CreateItem(0)
        mail.To = usermail
        mail.CC = cclist
        mail.Subject = "QC Observations - Claim #: " + claim_ + ", Score: "+ str(score)
        mail.Body = "Hi "+str(pname_)+",\n\n"+ maildata +"\nQC Observations -\n"+ qdata +"\nTotal Score: "+str(score)+"\nAuditor Comments: "+aucom+"\n\nRegards\n"+ str(loggedUser)
        mail.Send()
    
    
    def savecase():
        global variables
        global values
        global acomments
        global loggedUser
        setvalues = []
        try:
            score = 0
            for i in range(len(variables)):
                if(variables[i].get()==""):
                    aftersave()
                    messageBox.config(text="Incomplete\nDetails!")
                    messageBox.config(bg="#cc2f2f")
                    messageBox.config(fg="#ffffff")
                    return
                setvalues.append(variables[i].get())
                if(values[i][5]=="no"):
                    if(variables[i].get()=="No"):
                        score = 0
                        break
                    score = 100
                else:
                    score = score + float(variables[i].get())
            aucomm = str(acomments.get())
            qcdate = datetime.now()
            qcdate = qcdate.strftime("%Y-%m-%d %H:%M:%S")
            qcdate = str(qcdate)
            db.execute("SELECT ID FROM qcRecords WHERE mailID = ?",(values[0][0]))
            questionIds = []
            questionIds.clear()
            for ids in db.fetchall():
                questionIds.append(ids[0])
            for j in range(len(variables)):
                mailqid = str(values[j][0]) + str(values[j][1])
                if(qcstat.get()=="Pending"):
                    db.execute("INSERT INTO qcRecords (mailID, ques, setvalue,qcdate, mailquesID) VALUES(?,?,?,?,?)",(str(values[j][0]),str(values[j][2]),str(variables[j].get()),qcdate,mailqid))
                else:
                    db.execute("UPDATE qcRecords SET setvalue = ?,qcdate = ? WHERE ID = ? ",(str(variables[j].get()),qcdate, questionIds[j]))
        
            db.execute("UPDATE mailRecords SET qcstatus = ?, qcscore = ?, qccomment = ?,qcdate= ?,auditor = ? WHERE id = ? ",("Completed",str(score),aucomm,qcdate,loggedUser,values[0][0]))
            #save the data in database
            aftersave()
            global sendmail
            if(sendmail.get()=="Yes"):
                sendmail_func(setvalues,values,aucomm,score)
            messageBox.config(text="Action\nCompleted!")
            messageBox.config(bg="#71e04c")
            messageBox.config(fg="#ffffff")
        except:
            aftersave()
            messageBox.config(text="Something\nwent wrong!")
            messageBox.config(bg="#cc2f2f")
            messageBox.config(fg="#ffffff")
            print(traceback.format_exc())
        
    def selectcase(event):
        global qctable
        global variables
        global claim_
        global pdate_
        global pname_
        global mailsub_
        global gencom_
        global amount_
        try:
            selected_data = qctable.selection()[0]
            data_value = qctable.item(selected_data,'values')
            if(qcstat.get()=="Pending"):
                selectRec.config(text="Save/Stop")
                selectRec["command"] = savecase
            else:
                selectRec.config(text="Update/Stop")
                selectRec["command"] = savecase
            searchResult["state"] = "disabled"
            qcstatdrop2.configure(state="disabled")
            claimentry.delete(0,END)
            claimentry.insert(0,data_value[1])
            claim_ = data_value[1]
            cdateentry.delete(0,END)
            cdateentry.insert(0,data_value[2])
            pdate_ = data_value[2]
            opentry.delete(0,END)
            opentry.insert(0,data_value[3])
            pname_ = data_value[3]
            subentry.delete(0,END)
            subentry.insert(0,data_value[4])
            mailsub_ = data_value[4]
            amountentry.delete(0,END)
            amountentry.insert(0,data_value[5])
            amount_ = data_value[5]
            commententry.delete(0,END)
            commententry.insert(0,data_value[6])
            gencom_ = data_value[6]
            messageBox.config(text="")
            messageBox.config(bg="#ffffff")
            messageBox.config(fg="#000000")

            #call questionList function
            questionList(lob.get(),wt.get())
            if(qcstat.get()=="Completed"):
                db.execute("SELECT setvalue FROM qcRecords WHERE mailID = ?", (data_value[0]))
                qcdata = db.fetchall()
                for i in range(len(qcdata)):
                    variables[i].delete(0,END)
                    variables[i].insert(0,qcdata[i][0])
                db.execute("SELECT qccomment FROM mailRecords WHERE id = ?", (data_value[0]))    
                audcom = db.fetchone()
                acomments.delete(0,END)
                acomments.insert(0,audcom[0])
        except:
            messageBox.config(text="Something\nwent wrong!")
            messageBox.config(bg="#cc2f2f")
            messageBox.config(fg="#ffffff")
            print(traceback.format_exc())


    secondfrom = StringVar()
    secondto = StringVar()
    
    def datepick(mod, bname):
        datepicker = Tk()
 
        # Set geometry
        datepicker.title("Select Date")
        datepicker.geometry("300x270")
        
        # Add Calendar
        cal = Calendar(datepicker, selectmode = 'day',date_pattern='dd/mm/yyyy')
        
        cal.pack(pady = 20)
        
        def grad_date():
                bname.config(text=cal.get_date())
                #print(cal.get_date())
                mod.set(cal.get_date())
                
                datepicker.destroy()
        # Add Button and Label
        Button(datepicker, text = "Select Date",
            command = grad_date).pack()
        
        date = Label(datepicker, text = "")
        date.pack()
        
        # Execute Tkinter
        datepicker.mainloop()
    
    
    #mailbox dropdown
    # def mailchange(event):
    #     maillabel.config(text=mail.get())
    #     lobWt(mail.get())
    # mail = StringVar()
    # mailoptions = []
    # mailoptions.append('All')
    # db.execute('select distinct shortmail from maillob')
    # for mails in db.fetchall():
    #     mailoptions.append(mails[0])
    # drop1 = OptionMenu(frame1, mail, *(mailoptions), command=mailchange)
    # drop1.place(x=20,y=40, width=60, height=40)
    # maillabel = Label(frame1, text="Select Mailbox",font ="Helvetica 10 bold", fg="#1a1446", bg="#d6d6d6")
    # maillabel.place(x=80,y=40,height=40)
    
    #Operator dropdown
    def operatorchange(event):
        oplabel.config(text=op.get())
    op = StringVar()
    opoptions = []
    opoptions.append('All')
    db.execute('select distinct name from users')
    for opname in db.fetchall():
        opoptions.append(opname[0])
    drop2 = OptionMenu(frame1, op, *(opoptions), command=operatorchange)
    drop2.place(x=20,y=60, width=60, height=25)
    oplabel = Label(frame1, text="Select Operator",font ="Helvetica 9 bold", fg="#1a1446", bg="#d6d6d6")
    oplabel.place(x=80,y=60,height=25)
    
    def qcstatchange(event):
        qcstatlabel.config(text=qcstat.get())
    qcstat = StringVar()
    qcstatoptions = []
    qcstatoptions.append('Pending')
    qcstatoptions.append('Completed')
    qcstatdrop2 = OptionMenu(frame1, qcstat, *(qcstatoptions), command=qcstatchange)
    qcstatdrop2.place(x=20,y=15, width=60, height=25)
    qcstatlabel = Label(frame1, text="Select Status",font ="Helvetica 9 bold", fg="#1a1446", bg="#d6d6d6")
    qcstatlabel.place(x=80,y=15,height=25)
    
    def lobchange(event):
        loblabel.config(text=lob.get())
    
    def lobcallback(*args):
        
        wt.set('')
        drop4['menu'].delete(0, 'end')
        
        if(lob.get()=="All"):
            db.execute("select distinct worktypes from workTypes")
            for wotype in db.fetchall():
                drop4['menu'].add_command(label=wotype[0], command=tk._setit(wt, wotype[0]))
        else:
            thisset = set()
            db.execute("select distinct shortmail from maillob where lob = ? order by shortmail",(lob.get()))
            for mailb in db.fetchall():
                db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailb[0]))
                for wtype in db.fetchall():    
                    thisset.add(wtype[0])
            for woty in thisset:
                
                drop4['menu'].add_command(label=woty, command=tk._setit(wt, woty))
        wtlabel.config(text="Select Worktype")
    
    
    lob = StringVar()
    loboptions = []
    #loboptions.append('All')
    db.execute('select distinct lob from maillob')
    for lobname in db.fetchall():
        loboptions.append(lobname[0])
    drop3 = OptionMenu(frame1, lob, *(loboptions), command=lobchange)
    drop3.place(x=20,y=105, width=60, height=25)
    loblabel = Label(frame1, text="Select LOB",font ="Helvetica 9 bold", fg="#1a1446", bg="#d6d6d6")
    loblabel.place(x=80,y=105,height=25)
    lob.trace("w", lobcallback)
    
    
    def wtcallback(*args):
        wtlabel.config(text=wt.get())
    def wtchange(event):
        wtlabel.config(text=wt.get())
    wt = StringVar()
    wtoptions = []
    wtoptions.append('All')
    db.execute('select distinct worktypes from workTypes')
    for wtname in db.fetchall():
        wtoptions.append(wtname[0])
    drop4 = OptionMenu(frame1, wt, *(wtoptions), command=wtchange)
    drop4.place(x=20,y=150, width=60, height=25)
    wtlabel = Label(frame1, text="Select Worktype",font ="Helvetica 9 bold", fg="#1a1446", bg="#d6d6d6")
    wtlabel.place(x=80,y=150,height=25)
    wt.trace("w", wtcallback)
    
    
    range_ = Entry(frame1, background="#ebebeb", font=('Helvetica 12 bold'))
    range_.place(x=500,y=120,height=30,width=65)
    range_label = Label(frame1, text="Range %(1-100)", font="Helvetica 10 bold", fg="#1a1446",bg="#ffffff")
    range_label.place(x=500,y=95, height=25)
    
    Label(frame1, text="From",font = "Helvetica 9 bold",bg="#ffffff",fg="#1a1446").place(x=500,y=15, width=90, height = 20)
    secfromd = Button(frame1, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2, command= lambda: datepick(secondfrom,secfromd))
    secfromd.place(x=500,y=35, width=90, height = 30)
    Label(frame1, text="To",font = "Helvetica 9 bold",bg="#ffffff",fg="#1a1446").place(x=600,y=15, width=90, height = 20)
    sectod = Button(frame1, text="Select",font = "Helvetica 9 bold", bg="#1a1446",fg="white",highlightbackground="#e2e602", highlightthickness=2,command= lambda: datepick(secondto, sectod) )
    sectod.place(x=600,y=35, width=90, height = 30)
    
    reftotal = Button(frame1,text="Check Count", font="Helvetica 9 bold", command=getcount)
    reftotal.place(x=570,y=120,height=30,width=85)
    searchResult = Button(frame1,text="Search Results", font="Helvetica 10 bold", command= lambda: searchR(3,0))
    searchResult.place(x=500,y=190,width=158,height=47)
    selectRec = Button(frame1,text="Select Case", font="Helvetica 10 bold", state="disabled", command=selectcase)
    #selectRec.place(x=730,y=120,width=90,height=40)
    
    messageBox = Label(frame1,text="",font="Helvetica 8 bold",fg="#000000",bg="#ffffff")
    messageBox.place(x=700,y=95,width="120",height="80")
    
    search_label = Label(frame1, text="Search Claim Number", font="Helvetica 10 bold", fg="#1a1446",bg="#ffffff")
    search_label.place(x=20,y=190, height=25)
    searchQc = Entry(frame1,background="#ebebeb", font=('Helvetica 12 bold'))
    searchQc.place(x=20,y=210,width=460,height=30)

    
    cases_label = Label(frame1,text="", font = "Helvetica 8 bold",bg="#ffffff",fg="#1a1446")
    cases_label.place(x=20,y=175)

    
    #frame2
    Label(frame2,text="Claim #",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=32,width=130)
    Label(frame2,text="Processed Date",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=60,width=130)
    Label(frame2,text="Processor Name",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=88,width=130)
    Label(frame2,text="Email Subject",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=116,width=130)
    Label(frame2,text="Amount",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=144,width=130)
    Label(frame2,text="xx Comments",bg="#1a1446",fg="white",font="Helvetica 9 bold").place(x=5,y=172,width=130)
    claimentry = Entry(frame2)
    cdateentry = Entry(frame2)
    opentry = Entry(frame2)
    subentry = Entry(frame2)
    amountentry = Entry(frame2)
    commententry = Entry(frame2)
    claimentry.place(x=135,y=32,width=285, height=21)
    cdateentry.place(x=135,y=60,width=285, height=21)
    opentry.place(x=135,y=88,width=285, height=21)
    subentry.place(x=135,y=116,width=285, height=21)
    amountentry.place(x=135,y=144,width=285, height=21)
    commententry.place(x=135,y=172,width=285, height=21)
    
    #loggeduser
    global loggedUser
    Label(frame2,text=loggedUser,bg="#ffd000",fg="#1a1446",font="Helvetica 9 bold").place(x=5,y=2)
    
    # def lobWt(mailbox):
    # # Reset wt and delete all old options
    #     if(mailbox=="All"):
    #         db.execute("select distinct lob from maillob")
    #         loblabel.config(text="Select LOB")
    #     else:
    #         db.execute("select lob from maillob where shortmail = ? ",(mailbox))
    #         lob.set(db.fetchone()[0])
    #         loblabel.config(text=lob.get())
        
    #     wt.set('')
    #     drop4['menu'].delete(0, 'end')
        
    #     if(mailbox=="All"):
    #         db.execute("select distinct worktypes from workTypes")
    #     else:
    #         db.execute("select distinct worktypes from workTypes where mailbox = ? order by worktypes",(mailbox))
    #     for wts in db.fetchall():
    #         drop4['menu'].add_command(label=wts[0], command=tk._setit(wt, wts[0]))
    #     wtlabel.config(text="Select Worktype")
    
    
    
    qualwin.mainloop()

def reportsBAK():
    repwin = Tk()
    repwin.title('QC - Claim Automation Tracker')
    repwin.geometry("1280x720")
    repwin.state('zoomed')
    
    repwin.mainloop()

def reports():
    repwin = Tk()
    repwin.title('Reports - Claim Automation Tracker')
    repwin.geometry("1280x720")
    repwin.state('zoomed')
    frame1 = Frame(repwin, bg="#ffffff", highlightbackground="grey", highlightthickness=2)
    frame2 = Frame(repwin, bg="#ffffff", highlightbackground="grey", highlightthickness=1)
    frame3 = Frame(repwin, bg="#f2ffff", highlightbackground="grey", highlightthickness=2)
    frame4 = Frame(repwin, bg="#f2ffff", highlightbackground="grey", highlightthickness=1)

    frame1.place(x=0,y=1,width=1280,height=50)
    frame2.place(x=0,y=51,width=1280,height=279)
    frame3.place(x=0,y=331,width=1280,height=50)
    frame4.place(x=0,y=381,width=1280,height=339)
    
    #Frame1 components
    
    datevar = StringVar()
    fromvar = StringVar()
    tovar = StringVar()
    
    def datepick(mod, bname):
        exportDump.config(bg="#aaaef2")
        exportDump.config(text="Export Dump 💾")
        datepicker = Tk()
 
        # Set geometry
        datepicker.title("Select Date")
        datepicker.geometry("300x270")
        
        # Add Calendar
        cal = Calendar(datepicker, selectmode = 'day',date_pattern='dd-mm-yyyy')
        
        cal.pack(pady = 20)
        
        def grad_date():
                bname.config(text=cal.get_date())
                #print(cal.get_date())
                mod.set(cal.get_date())
                
                datepicker.destroy()
        # Add Button and Label
        Button(datepicker, text = "Select Date",
            command = grad_date).pack()
        
        date = Label(datepicker, text = "")
        date.pack()
        
        # Execute Tkinter
        datepicker.mainloop()
        
    
    Label(frame1, text="EOD Status Report                                 |",bg="#ffffff",font = "Helvetica 14 bold").place(x=10,y=0,height=45)
    
    Label(frame1, text="Select Day",bg="#ffffff",font = "Helvetica 12").place(x=400,y=0,height=45)
    
    todatestr = datetime.now()
    todatestr = todatestr.strftime('%d-%m-%Y')
    
    def OptionMenu_Select3(event):
        typelabel.config(text=typevar.get())
        exportEod["state"] = "disabled"

    typevar = StringVar()
    typelist =  []
    typelist.append("LOB-Worktype-Wise")
    typelist.append("Mailbox-Wise")
    typelist.append("User-Wise")
    OptionMenu(frame1, typevar, *(typelist),command=OptionMenu_Select3).place(x=615, y=4,  width=40, height=38)
    typelabel = Label(frame1,text="Select Report Type",font="Helvetica 10")
    typelabel.place(x=655,y=4,height=38)
    
    dateBut = Button(frame1, text=todatestr,bg="#0c007a",fg="#ffffff", font="Helvetica 12", command= lambda: datepick(datevar, dateBut))
    dateBut.place(x=490,y=4,height=38)
    
    dateSub = Button(frame1, text="Search 🔎", font="Helvetica 12", command=lambda: showEodTable(typevar.get()))
    dateSub.place(x=835,y=4,height=38)
    
    Label(frame1, text="|",bg="#ffffff",font = "Helvetica 14 bold").place(x=960,y=0,height=45)
    def printlist(typee):
        global exportlist
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
        if len(tempdir) > 0:
            today = datetime.now()
            today = today.strftime('%H%M%S%f%Y%m%d')
            if(typevar.get()=="LOB-Worktype-Wise"):
                df = pd.DataFrame (exportlist, columns = ['Activity','Sub-Activity','Volume Recieved','Task Touched','Task Completed','Completed %','Query/hold Task','Query %','TAT','Comments'])
                pathnamemailwise = tempdir + "\\EOD_LOBWT_"+ str(datevar.get())+".csv"
            if(typevar.get()=="Mailbox-Wise"):
                df = pd.DataFrame (exportlist, columns = ['Mailbox','Backlog','Volume Recieved','Task Touched','Task Completed','Completed %','Query/hold Task','Query %','Pending','Pending %','TAT','Comments'])
                pathnamemailwise = tempdir + "\\EOD_Mailbox_"+ str(datevar.get())+".csv"
            if(typevar.get()=="User-Wise"):
                df = pd.DataFrame (exportlist, columns = ['User','Volume Recieved','Task Touched','Task Completed','Completed %','Query/hold Task','Query %','TAT','Comments'])
                pathnamemailwise = tempdir + "\\EOD_User_"+ str(datevar.get())+".csv"
            df.to_csv(r''+ pathnamemailwise, encoding='utf-8-sig',index=False)
            exportEod.config(bg="#b4ff94")
            exportEod.config(text="File Exported!")
        else:
            exportEod.config(bg="#ff6969")
            exportEod.config(text="Please try again!")
    
        #print(exportlist)
    exportEod = Button(frame1, text="Export Data 💾", font="Helvetica 12",state="disabled", command=lambda: printlist(typevar.get()))
    exportEod.place(x=1065,y=4,height=38)
    
    #frame2 Components
    def showEodTable(type):
        datevarfrom = datetime.strptime(datevar.get(),"%d-%m-%Y")
        datevarto = datevarfrom + timedelta(days=1)
        global exportlist
        exportlist = []
        exportlist.clear()
        totalc = 0
        db.execute("Select distinct lob from maillob")
        for lobs in db.fetchall():
            db.execute("Select distinct worktypes from workTypes where lob = ?",(lobs[0]))
            for mailIds in db.fetchall():
                totalc = totalc+1        
        try:
            if(type=="LOB-Worktype-Wise"):
            
                eod_scroll = Scrollbar(frame2)
                eod_scroll.place(x=1260,y=1,height=276)
                global eodtable
                eodtable = ttk.Treeview(frame2,show='headings', height=12, yscrollcommand=eod_scroll.set)
                eodtable['columns'] = ('ID','lob','wt','totvol','touch','completed','completeper','queryhold','queryper','tat','comments')
                eodtable.column("#0", width=0,  stretch=NO)
                eodtable.column("ID", width=0,  stretch=NO)
                eodtable.column("lob",anchor=W,width=215)
                eodtable.column("wt",anchor=W,width=225)
                eodtable.column("totvol",anchor=N,width=110)
                eodtable.column("touch",anchor=N,width=100)
                eodtable.column("completed",anchor=N,width=100)
                eodtable.column("completeper",anchor=N,width=90)
                eodtable.column("queryhold",anchor=N,width=110)
                eodtable.column("queryper",anchor=N,width=90)
                eodtable.column("tat",anchor=N,width=100)
                eodtable.column("comments",anchor=N,width=110)

                eodtable.heading("#0",text="",anchor=CENTER)
                eodtable.heading("ID",text="",anchor=CENTER)
                eodtable.heading("lob",text="Activity",anchor=CENTER)
                eodtable.heading("wt",text="Sub-Activity",anchor=CENTER)
                eodtable.heading("totvol",text="Volume Recieved",anchor=CENTER)
                eodtable.heading("touch",text="Task Touched",anchor=CENTER)
                eodtable.heading("completed",text="Task Completed",anchor=CENTER)
                eodtable.heading("completeper",text="Completed %",anchor=CENTER)
                eodtable.heading("queryhold",text="Query/Hold Task",anchor=CENTER)
                eodtable.heading("queryper",text="Query %",anchor=CENTER)
                eodtable.heading("tat",text="TAT",anchor=CENTER)
                eodtable.heading("comments",text="Comments",anchor=CENTER)
                ids=0
                #print(datevar.get())
                
                
                try:
                    db.execute("Select distinct lob from maillob")
                    for lobs in db.fetchall():
                        db.execute("Select distinct worktypes from workTypes where lob = ?",(lobs[0]))
                        for mailIds in db.fetchall():
                            ids=ids+1
                            #volume recieved
                            db.execute("Select count(entryId) from mailRecords where lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datevarfrom,datevarto))
                            totalvoll = db.fetchone()
                            #touch tasks
                            db.execute("Select count(entryId) from mailRecords where NOT status='Pending' AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datevarfrom,datevarto))
                            touchtask = db.fetchone()
                            #taskcompleted
                            db.execute("Select count(entryId) from mailRecords where status IN ('Completed','No Action Required') AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datevarfrom,datevarto))
                            comptask = db.fetchone()
                            #completed%
                            if(int(comptask[0])==0):
                                compperc = "0.0 %"
                            else:
                                compperc = round((int(comptask[0])/int(totalvoll[0]))*100,1)
                                compperc = str(compperc) + " %"
                            #query/hold
                            db.execute("Select count(entryId) from mailRecords where status IN ('Hold','Queried-Onshore') AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datevarfrom,datevarto))
                            qhtask = db.fetchone()
                            #query%
                            if(int(qhtask[0])==0):
                                qhperc = "0.0 %"
                            else:
                                qhperc = round((int(qhtask[0])/int(totalvoll[0]))*100,1)
                                qhperc = str(qhperc) + " %"
                            #tat
                            db.execute("select tat from mailRecords where lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datevarfrom,datevarto))
                            tatval = "-"
                            for xxx in db.fetchall():
                                if(xxx[0]=="Missed"):
                                    tatval="Missed"
                                    break
                                tatval = "Met"
                            tempdata1 = []
                            tempdata1.clear()
                            tempdata1.append(lobs[0])
                            tempdata1.append(mailIds[0])
                            tempdata1.append(totalvoll[0])
                            tempdata1.append(touchtask[0])
                            tempdata1.append(comptask[0])
                            tempdata1.append(compperc)
                            tempdata1.append(qhtask[0])
                            tempdata1.append(qhperc)
                            tempdata1.append(tatval)
                            tempdata1.append("-")
                            exportlist.append(tempdata1)
                            #table Rows
                            perc = round((int(ids)/int(totalc))*100,0)
                            perc = "Loading "+str(perc)+ "%"
                            print(perc)
                            eodtable.insert(parent='',index='end',iid=ids,text='', values=(ids,lobs[0],mailIds[0],totalvoll[0],touchtask[0],comptask[0],compperc,qhtask[0],qhperc,tatval,"-"))

                except:
                    print("Please try again!")
                    print(traceback.format_exc())

                eodtable.place(x=3,y=4)
                eod_scroll.config(command=eodtable.yview)
                exportEod["state"] = "normal"
                exportEod.config(bg="#f0f0f0")
                exportEod.config(text="Export Data 💾")
            if(type=="User-Wise"):
            
                eod_scroll = Scrollbar(frame2)
                eod_scroll.place(x=1260,y=1,height=276)
                global eodtable3
                eodtable3 = ttk.Treeview(frame2,show='headings', height=12, yscrollcommand=eod_scroll.set)
                eodtable3['columns'] = ('ID','name','totvol','touch','completed','completeper','queryhold','queryper','tat','comments')
                eodtable3.column("#0", width=0,  stretch=NO)
                eodtable3.column("ID", width=0,  stretch=NO)
                eodtable3.column("name",anchor=W,width=240)
                eodtable3.column("totvol",anchor=N,width=135)
                eodtable3.column("touch",anchor=N,width=125)
                eodtable3.column("completed",anchor=N,width=125)
                eodtable3.column("completeper",anchor=N,width=115)
                eodtable3.column("queryhold",anchor=N,width=135)
                eodtable3.column("queryper",anchor=N,width=115)
                eodtable3.column("tat",anchor=N,width=125)
                eodtable3.column("comments",anchor=N,width=135)

                eodtable3.heading("#0",text="",anchor=CENTER)
                eodtable3.heading("ID",text="",anchor=CENTER)
                eodtable3.heading("name",text="User",anchor=CENTER)
                eodtable3.heading("totvol",text="Volume Recieved",anchor=CENTER)
                eodtable3.heading("touch",text="Task Touched",anchor=CENTER)
                eodtable3.heading("completed",text="Task Completed",anchor=CENTER)
                eodtable3.heading("completeper",text="Completed %",anchor=CENTER)
                eodtable3.heading("queryhold",text="Query/Hold Task",anchor=CENTER)
                eodtable3.heading("queryper",text="Query %",anchor=CENTER)
                eodtable3.heading("tat",text="TAT",anchor=CENTER)
                eodtable3.heading("comments",text="Comments",anchor=CENTER)
                ids=0
                #print(datevar.get())
                
                
                try:
                    db.execute("Select distinct name from users order by name")
                    for users in db.fetchall():
                        ids=ids+1
                        #volume recieved
                        db.execute("Select count(entryId) from mailRecords where user=? AND status IN ('Completed','Hold','No Action Required','Queried-Onshore') AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(users[0]),datevarfrom,datevarto))
                        totalvoll = db.fetchone()
                        #touch tasks
                        db.execute("Select count(entryId) from mailRecords where user=? AND status IN ('Completed','Hold','No Action Required','Queried-Onshore') AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(users[0]),datevarfrom,datevarto))
                        touchtask = db.fetchone()
                        #taskcompleted
                        db.execute("Select count(entryId) from mailRecords where user=? AND status IN ('Completed','No Action Required') AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(users[0]),datevarfrom,datevarto))
                        comptask = db.fetchone()
                        #completed%
                        if(int(comptask[0])==0):
                            compperc = "0.0 %"
                        else:
                            compperc = round((int(comptask[0])/int(totalvoll[0]))*100,1)
                            compperc = str(compperc) + " %"
                        #query/hold
                        db.execute("Select count(entryId) from mailRecords where user=? AND status IN ('Hold','Queried-Onshore') AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(users[0]),datevarfrom,datevarto))
                        qhtask = db.fetchone()
                        #query%
                        if(int(qhtask[0])==0):
                            qhperc = "0.0 %"
                        else:
                            qhperc = round((int(qhtask[0])/int(totalvoll[0]))*100,1)
                            qhperc = str(qhperc) + " %"
                        #tat
                        db.execute("select tat from mailRecords where user=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(users[0]),datevarfrom,datevarto))
                        tatval = "-"
                        for xxx in db.fetchall():
                            if(xxx[0]=="Missed"):
                                tatval="Missed"
                                break
                            tatval = "Met"
                        tempdata1 = []
                        tempdata1.clear()
                        tempdata1.append(users[0])
                        tempdata1.append(totalvoll[0])
                        tempdata1.append(touchtask[0])
                        tempdata1.append(comptask[0])
                        tempdata1.append(compperc)
                        tempdata1.append(qhtask[0])
                        tempdata1.append(qhperc)
                        tempdata1.append(tatval)
                        tempdata1.append("-")
                        exportlist.append(tempdata1)
                        #table Rows
                        perc = round((int(ids)/int(totalc))*100,0)
                        perc = "Loading "+str(perc)+ "%"
                        print(perc)
                        eodtable3.insert(parent='',index='end',iid=ids,text='', values=(ids,users[0],totalvoll[0],touchtask[0],comptask[0],compperc,qhtask[0],qhperc,tatval,"-"))

                except:
                    print("Please try again!")
                    print(traceback.format_exc())

                eodtable3.place(x=3,y=4)
                eod_scroll.config(command=eodtable3.yview)
                exportEod["state"] = "normal"
                exportEod.config(bg="#f0f0f0")
                exportEod.config(text="Export Data 💾")
            if(type=="Mailbox-Wise"):
                eod_scroll = Scrollbar(frame2)
                eod_scroll.place(x=1260,y=1,height=276)
                global eodtable2
                eodtable2 = ttk.Treeview(frame2,show='headings', height=12, yscrollcommand=eod_scroll.set)
                eodtable2['columns'] = ('ID','mail','curback','totvol','touch','completed','completeper','queryhold','queryper','pending','pendper','tat','comments')
                eodtable2.column("#0", width=0,  stretch=NO)
                eodtable2.column("ID", width=0,  stretch=NO)
                eodtable2.column("mail",anchor=W,width=220)
                eodtable2.column("curback",anchor=N,width=70)
                eodtable2.column("totvol",anchor=N,width=110)
                eodtable2.column("touch",anchor=N,width=80)
                eodtable2.column("completed",anchor=N,width=100)
                eodtable2.column("completeper",anchor=N,width=90)
                eodtable2.column("queryhold",anchor=N,width=110)
                eodtable2.column("queryper",anchor=N,width=90)
                eodtable2.column("pending",anchor=N,width=80)
                eodtable2.column("pendper",anchor=N,width=80)
                eodtable2.column("tat",anchor=N,width=110)
                eodtable2.column("comments",anchor=N,width=110)

                eodtable2.heading("#0",text="",anchor=CENTER)
                eodtable2.heading("ID",text="",anchor=CENTER)
                eodtable2.heading("mail",text="MailBox",anchor=CENTER)
                eodtable2.heading("curback",text="Backlog",anchor=CENTER)
                eodtable2.heading("totvol",text="Volume Recieved",anchor=CENTER)
                eodtable2.heading("touch",text="Task Touched",anchor=CENTER)
                eodtable2.heading("completed",text="Task Completed",anchor=CENTER)
                eodtable2.heading("completeper",text="Completed %",anchor=CENTER)
                eodtable2.heading("queryhold",text="Query/Hold Task",anchor=CENTER)
                eodtable2.heading("queryper",text="Query %",anchor=CENTER)
                eodtable2.heading("pending",text="Pending",anchor=CENTER)
                eodtable2.heading("pendper",text="Pending %",anchor=CENTER)
                eodtable2.heading("tat",text="TAT",anchor=CENTER)
                eodtable2.heading("comments",text="Comments",anchor=CENTER)
                ids=0
                try:
                    db.execute("Select distinct shortmail from maillob order by shortmail")
                    for mails in db.fetchall():
                        #backlog
                        db.execute("Select count(entryId) from mailRecords where mailBox=? AND status='Pending' AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom))
                        backlogvol = db.fetchone()
                        #volume recieved
                        db.execute("Select count(entryId) from mailRecords where mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        totalvoll = db.fetchone()
                        #touch tasks
                        db.execute("Select count(entryId) from mailRecords where NOT status='Pending' AND mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        touchtask = db.fetchone()
                        #taskcompleted
                        db.execute("Select count(entryId) from mailRecords where status IN ('Completed','No Action Required') AND mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        comptask = db.fetchone()
                        #completed%
                        if(int(comptask[0])==0):
                            compperc = "0.0 %"
                        else:
                            compperc = round((int(comptask[0])/int(totalvoll[0]))*100,1)
                            compperc = str(compperc) + " %"
                        #query/hold
                        db.execute("Select count(entryId) from mailRecords where status IN ('Hold','Queried-Onshore') AND mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        qhtask = db.fetchone()
                        #query%
                        if(int(qhtask[0])==0):
                            qhperc = "0.0 %"
                        else:
                            qhperc = round((int(qhtask[0])/int(totalvoll[0]))*100,1)
                            qhperc = str(qhperc) + " %"
                        #tat
                        db.execute("select tat from mailRecords where mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        tatval = "-"
                        for xxx in db.fetchall():
                            if(xxx[0]=="Missed"):
                                tatval="Missed"
                                break
                            tatval = "Met"
                        #Pending
                        db.execute("Select count(entryId) from mailRecords where status='Pending' AND mailBox=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(mails[0]),datevarfrom,datevarto))
                        ptask = db.fetchone()
                        #Pending %
                        if(int(ptask[0])==0):
                            pperc = "0.0 %"
                        else:
                            pperc = round((int(ptask[0])/int(totalvoll[0]))*100,1)
                            pperc = str(pperc) + " %"
                        ids=ids+1
                        perc = round((int(ids)/int(totalc))*100,0)
                        perc = "Loading "+str(perc)+ "%"
                        print(perc)
                        eodtable2.insert(parent='',index='end',iid=ids,text='', values=(ids,mails[0],backlogvol[0],totalvoll[0],touchtask[0],comptask[0],compperc,qhtask[0],qhperc,ptask[0],pperc,tatval,"-"))

                except:
                    print("Please try again!")
                    print(traceback.format_exc())

                eodtable2.place(x=3,y=4)
                eod_scroll.config(command=eodtable2.yview)
                exportEod["state"] = "normal"
        except:
            pass
            print(traceback.format_exc())
    
    def exportdumplist():
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
        datefromvar = datetime.strptime(fromvar.get(),"%d-%m-%Y")
        datetovar = datetime.strptime(tovar.get(),"%d-%m-%Y")
        datetovar = datetovar + timedelta(days=1)
        if len(tempdir) > 0:
            dumplist = []
            dumplist.clear()
            dumpcount = 0
            db.execute("select * from mailRecords where lastsaved >= CDate(?) AND lastsaved < CDate(?)", (datefromvar, datetovar))
            for datalist in db.fetchall():
                templist = []
                templist.clear()
                if(datalist[20]=="Pending"):
                    dumpcount = dumpcount + 1
                    templist.append(dumpcount)
                    templist.append(datalist[0])
                    templist.append(datalist[2])
                    templist.append(datalist[3])
                    templist.append(datalist[4])
                    templist.append(datalist[5])
                    templist.append(datalist[6])
                    templist.append(datalist[7])
                    templist.append(datalist[10])
                    templist.append(datalist[11])
                    templist.append(datalist[19])
                    templist.append(datalist[12])
                    templist.append(datalist[13])
                    templist.append(datalist[14])
                    templist.append(datalist[15])
                    templist.append(datalist[25])
                    templist.append(datalist[16])
                    templist.append(datalist[17])
                    db.execute("select top 1 startedAt, endedAt, timeTaken from prodRecords where mailID=?",(datalist[0]))
                    ahtdata = db.fetchone()
                    if(ahtdata==None):
                        templist.append("-")
                        templist.append("-")
                        templist.append("-")
                    else:
                        templist.append(ahtdata[0])
                        templist.append(ahtdata[1])
                        templist.append(ahtdata[2])
                    templist.append(datalist[18])
                    templist.append(datalist[20])
                    templist.append("")
                    templist.append(datalist[21])
                    templist.append(datalist[22])
                    templist.append(datalist[23])
                    templist.append(datalist[24])
                    dumplist.append(templist)
                
                if(datalist[20]=="Completed"):
                    db.execute("Select ques, setvalue from qcRecords where mailID=?", (datalist[0]))
                    for qcd in db.fetchall():
                        #print(qcd[0])
                        templist1 = []
                        templist1.clear()
                        dumpcount = dumpcount + 1
                        templist1.append(dumpcount)
                        templist1.append(datalist[0])
                        templist1.append(datalist[2])
                        templist1.append(datalist[3])
                        templist1.append(datalist[4])
                        templist1.append(datalist[5])
                        templist1.append(datalist[6])
                        templist1.append(datalist[7])
                        templist1.append(datalist[10])
                        templist1.append(datalist[11])
                        templist1.append(datalist[19])
                        templist1.append(datalist[12])
                        templist1.append(datalist[13])
                        templist1.append(datalist[14])
                        templist1.append(datalist[15])
                        templist1.append(datalist[25])
                        templist1.append(datalist[16])
                        templist1.append(datalist[17])
                        db.execute("select top 1 startedAt, endedAt, timeTaken from prodRecords where mailID=?",(datalist[0]))
                        ahtdata1 = db.fetchone()
                        
                        if(ahtdata1==None):
                            templist1.append("-")
                            templist1.append("-")
                            templist1.append("-")
                        else:
                            templist1.append(ahtdata1[0])
                            templist1.append(ahtdata1[1])
                            templist1.append(ahtdata1[2])
                        templist1.append(datalist[18])
                        templist1.append(datalist[20])
                        templist1.append(qcd[0])
                        templist1.append(qcd[1])
                        templist1.append(datalist[22])
                        templist1.append(datalist[23])
                        templist1.append(datalist[24])
                        dumplist.append(templist1)
                        
            today = datetime.now()
            today = today.strftime('%H%M%S%f%m%d')
            df = pd.DataFrame (dumplist, columns = ['#','Mail ID','From Mail','Subject','Recieved At','MailBox','Folder','Status','Last Saved At','WorkType','LOB','Policy Number','Claim Number','UCR','Processor Comments','Currency','Amount','Processor Name','Started At','Ended At','Time Taken[Mins]','TAT','QC Status','QC Question','QC Score','Auditor Comments','QC Date','Auditor Name'])
            pathnamemailwise = tempdir + "\\Claims_Dump_"+ str(fromvar.get()) + "_to_"+ str(tovar.get()) +".csv"
            df.to_csv(r''+ pathnamemailwise, encoding='utf-8-sig',index=False)
            exportDump.config(bg="#b4ff94")
            exportDump.config(text="File Exported!")
        else:
            exportDump.config(bg="#ff6969")
            exportDump.config(text="Please try again!")
    
    def exportprodlist():
        global prodlist
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
        if len(tempdir) > 0:
            today = datetime.now()
            today = today.strftime('%H%M%S%f%m%d')
            df = pd.DataFrame (prodlist, columns = ['LOB','Sub Activity','Production Target','Production Met','Production Met %','TAT Target','TAT Met','Query #','Query %','Accuracy Target','Accuracy Audit %','Accuracy Internal QC %','Accuracy Met','Comments','Files Audited','Files Processed','Opportunities Achieved','Total Opportunities'])
            pathnamemailwise = tempdir + "\\ProdReport_"+ str(fromvar.get()) + "_to_"+ str(tovar.get()) +".csv"
            df.to_csv(r''+ pathnamemailwise, encoding='utf-8-sig',index=False)
            exportProd.config(bg="#b4ff94")
            exportProd.config(text="File Exported!")
        else:
            exportProd.config(bg="#ff6969")
            exportProd.config(text="Please try again!")
    #Frame4 components
    def showProdTable():
        datefromvar = datetime.strptime(fromvar.get(),"%d-%m-%Y")
        datetovar = datetime.strptime(tovar.get(),"%d-%m-%Y")
        datetovar = datetovar + timedelta(days=1)
        global prodlist
        prodlist = []
        prodlist.clear()
        totalc = 0
        db.execute("Select distinct lob from maillob")
        for lobs in db.fetchall():
            db.execute("Select distinct worktypes from workTypes where lob = ?",(lobs[0]))
            for mailIds in db.fetchall():
                totalc = totalc+1  
        try:
            eod_scroll = Scrollbar(frame4)
            eod_scroll.place(x=1260,y=1,height=305)
            global prodtable
            prodtable = ttk.Treeview(frame4,show='headings', height=14, yscrollcommand=eod_scroll.set)
            prodtable['columns'] = ('ID','lob','wt','prodtar','prodmet','prodmetper','tattar','tatmet','query',
                                    'queryper','acctar','accaudit','accint','accmet','comment','fileaud','filepro',
                                    'oppachi','totopp')
            prodtable.column("#0", width=0,  stretch=NO)
            prodtable.column("ID", width=0,  stretch=NO)
            prodtable.column("lob",anchor=W,width=120)
            prodtable.column("wt",anchor=W,width=115)
            prodtable.column("prodtar",anchor=N,width=70)
            prodtable.column("prodmet",anchor=N,width=70)
            prodtable.column("prodmetper",anchor=N,width=70)
            prodtable.column("tattar",anchor=N,width=65)
            prodtable.column("tatmet",anchor=N,width=60)
            prodtable.column("query",anchor=N,width=50)
            prodtable.column("queryper",anchor=N,width=50)
            prodtable.column("acctar",anchor=N,width=80)
            prodtable.column("accaudit",anchor=N,width=50)
            prodtable.column("accint",anchor=N,width=50)
            prodtable.column("accmet",anchor=N,width=65)
            prodtable.column("comment",anchor=N,width=60)
            prodtable.column("fileaud",anchor=N,width=65)
            prodtable.column("filepro",anchor=N,width=65)
            prodtable.column("oppachi",anchor=N,width=80)
            prodtable.column("totopp",anchor=N,width=65)

            prodtable.heading("#0",text="",anchor=CENTER)
            prodtable.heading("ID",text="",anchor=CENTER)
            prodtable.heading("lob",text="LOB\n",anchor=CENTER)
            prodtable.heading("wt",text="Sub-Activity",anchor=CENTER)
            prodtable.heading("prodtar",text="Prod Target",anchor=CENTER)
            prodtable.heading("prodmet",text="Prod Met",anchor=CENTER)
            prodtable.heading("prodmetper",text="Prod Met%",anchor=CENTER)
            prodtable.heading("tattar",text="TAT Target",anchor=CENTER)
            prodtable.heading("tatmet",text="TAT Met",anchor=CENTER)
            prodtable.heading("query",text="Query",anchor=CENTER)
            prodtable.heading("queryper",text="Query%",anchor=CENTER)
            prodtable.heading("acctar",text="Accu. Target",anchor=CENTER)
            prodtable.heading("accaudit",text="Audit %",anchor=CENTER)
            prodtable.heading("accint",text="QC%",anchor=CENTER)
            prodtable.heading("accmet",text="Accu. Met",anchor=CENTER)
            prodtable.heading("comment",text="Comments",anchor=CENTER)
            prodtable.heading("fileaud",text="Audited#",anchor=CENTER)
            prodtable.heading("filepro",text="Processed#",anchor=CENTER)
            prodtable.heading("oppachi",text="Opp. Achieved",anchor=CENTER)
            prodtable.heading("totopp",text="Total Opp.",anchor=CENTER)
            style = ttk.Style()
            style.configure('Treeview.Heading', foreground='black') 
            ids=0
            db.execute("Select distinct lob from maillob")
            for lobs in db.fetchall():
                db.execute("Select distinct worktypes from workTypes where lob = ?",(lobs[0]))
                for mailIds in db.fetchall():
                    #prod target
                    db.execute("Select count(entryId) from mailRecords where lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    totalvoll = db.fetchone()
                    #prod target
                    db.execute("Select count(entryId) from mailRecords where lob=? AND worktype=? AND tat='Met' AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    prodmet = db.fetchone()
                    #prod met%
                    if(int(prodmet[0])==0):
                        prodmetper = 0.0
                        prodmetper = str(prodmetper)+ " %"
                    else:
                        prodmetper = round((int(prodmet[0])/int(totalvoll[0]))*100,1)
                        prodmetper = str(prodmetper)+ " %"
                    #query
                    db.execute("Select count(entryId) from mailRecords where status IN ('Hold','Queried-Onshore') AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    qhtask = db.fetchone()
                    #query%
                    if(int(qhtask[0])==0):
                        qhtaskper = 0.0
                        qhtaskper = str(qhtaskper)+ " %"
                    else:
                        qhtaskper = round((int(qhtask[0])/int(totalvoll[0]))*100,1)
                        qhtaskper = str(qhtaskper)+ " %"
                    #qcaudited
                    db.execute("Select count(entryId) from mailRecords where qcstatus='Completed' AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    qhaudit = db.fetchone()
                    #print(datefromvar,datetovar)
                    #qcaudited%
                    if(int(qhaudit[0])==0):
                        qhauditper = 0.0
                        qhauditper = str(qhauditper)+ " %"
                    else:
                        qhauditper = round((int(qhaudit[0])/int(totalvoll[0]))*100,1)
                        qhauditper = str(qhauditper)+ " %"
                    #accqcmet
                    db.execute("Select qcscore from mailRecords where qcstatus='Completed' AND lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    totalscore = 0
                    qcid = 0
                    accumet = 0
                    for xxd in db.fetchall():
                        qcid = qcid + 1
                        totalscore = float(xxd[0]) + totalscore
                    if(qcid<1):
                        accumets = "0 %"
                    else:
                        accumet = totalscore/qcid
                        accumets = str(accumet) + " %"
                    #processed
                    db.execute("Select count(entryId) from mailRecords where lob=? AND worktype=? AND lastsaved >= CDate(?) AND lastsaved < CDate(?)", (str(lobs[0]),str(mailIds[0]),datefromvar,datetovar))
                    processed = db.fetchone()
                    #oppac
                    oppachieve = int(qhaudit[0])*accumet
                    opptotal = int(qhaudit[0])*100
                    
                    tempdata2 = []
                    tempdata2.clear()
                    tempdata2.append(lobs[0])
                    tempdata2.append(mailIds[0])
                    tempdata2.append(totalvoll[0])
                    tempdata2.append(prodmet[0])
                    tempdata2.append(prodmetper)
                    tempdata2.append('95%')
                    tempdata2.append(prodmetper)
                    tempdata2.append(qhtask[0])
                    tempdata2.append(qhtaskper)
                    tempdata2.append('95%')
                    tempdata2.append(qhauditper)
                    tempdata2.append('-')
                    tempdata2.append(accumets)
                    tempdata2.append('-')
                    tempdata2.append(qhaudit[0])
                    tempdata2.append(processed[0])
                    tempdata2.append(oppachieve)
                    tempdata2.append(opptotal)
                    prodlist.append(tempdata2)
                    
                    ids = ids + 1
                    perc = round((int(ids)/int(totalc))*100,0)
                    perc = "Loading "+str(perc)+ "%"
                    print(perc)
                    prodtable.insert(parent='',index='end',iid=ids,text='', values=(ids,lobs[0],mailIds[0],totalvoll[0],prodmet[0],prodmetper,'95%',prodmetper,qhtask[0],qhtaskper,'95%',qhauditper,'-',accumets,'-',qhaudit[0],processed[0],oppachieve,opptotal))
            
            prodtable.place(x=3,y=4)
            eod_scroll.config(command=prodtable.yview)
            exportProd["state"] = "normal"
            exportProd.config(bg="#f0f0f0")
            exportProd.config(text="Export Data 💾")
            
        except:
            print(traceback.format_exc())
            

    #frame3 Components
    Label(frame3, text="Claims Production Report                 |",bg="#f2ffff",font = "Helvetica 14 bold").place(x=10,y=0,height=45)
    
    Label(frame3, text="Select Range",bg="#f2ffff",font = "Helvetica 12").place(x=420,y=0,height=45)
    
    dateButFrom = Button(frame3, text=todatestr,bg="#0c007a",fg="#ffffff", font="Helvetica 12", command= lambda: datepick(fromvar, dateButFrom))
    dateButFrom.place(x=530,y=4,height=38)
    
    Label(frame3, text="To",bg="#f2ffff",font = "Helvetica 12").place(x=640,y=0,height=45)
    
    dateButTo = Button(frame3, text=todatestr,bg="#0c007a",fg="#ffffff", font="Helvetica 12", command= lambda: datepick(tovar, dateButTo))
    dateButTo.place(x=675,y=4,height=38)
    
    dateSubProd = Button(frame3, text="Search 🔎", font="Helvetica 12", command=showProdTable)
    dateSubProd.place(x=795,y=4,height=38)
    
    Label(frame3, text="|",bg="#f2ffff",font = "Helvetica 14 bold").place(x=915,y=0,height=45)
    
    exportProd = Button(frame3, text="Export Table 💾", font="Helvetica 12",state="disabled", command=exportprodlist)
    exportProd.place(x=955,y=4,height=38)
    
    exportDump = Button(frame3, text="Export Dump 💾", font="Helvetica 12",bg="#aaaef2", command=exportdumplist)
    exportDump.place(x=1115,y=4,height=38)
    
    datevar.set(todatestr)
    tovar.set(todatestr)
    fromvar.set(todatestr)
    
    
    
    repwin.mainloop()

# db.execute("Select distinct mailBox from mailRecords")
# for cvb in db.fetchall():
#     db.execute("select lob from maillob where shortmail = ? ",(cvb[0]))
#     lobb = db.fetchone()
#     #db.execute("update mailRecords set lob=? where mailBox = ? and status in ('Completed','No Action Required','Queried-Onshore','Hold')",(lobb[0],cvb[0]))

#buttons and Label
label = Label(win,bg="white", text="Welcome"+" " +loggedUser, font = "Helvetica 14 bold")
label.pack(pady=10)
b1 = Button(win,text="Supervisor", font = "Helvetica 14 bold", command=supervisor)
b1.place(x=80, y=240)
b2 = Button(win,text="Production", font = "Helvetica 14 bold", command=production)
b2.place(x=260, y=240)
b3 = Button(win,text="     Reports    ", font = "Helvetica 14 bold",command=reports)
b3.place(x=440, y=240)
b4 = Button(win,text="        QC        ", font = "Helvetica 14 bold", command=quality)
b4.place(x=630, y=240)
b5 = Button(win,text="     Access    ", font = "Helvetica 14 bold", command=access)
b5.place(x=340, y=310)

if(loggedUserAuth!="super"):
    b1["state"] = "disabled"
    b3["state"] = "disabled"
    b4["state"] = "disabled"
    b5["state"] = "disabled"

try:
    db.execute("update users set lastloggedin = ? where nid = ?",(datetime.now(),os.getenv('username')))
except:
    pass
root.mainloop()
