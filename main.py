import sqlite3
import tkinter as tk
import tkinter.messagebox
from tkinter import font
from tkinter import *
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import tkinter.filedialog
import sqlite3 as sl
from tkinter import ttk
from imbox import Imbox

root = Tk()
root.title('Mail_Client_Ver_0.1')
root.eval('tk::PlaceWindow . center')
root.resizable(False, False)


class Main_Page:
    def __init__(self, root=None):
        self.root = root
        self.frame = tk.Frame(self.root)
        self.frame.grid()
        self.address_page = Address_book(master=self.root, app=self)  # Important, create frames from classes here
        self.mail_getter = Mail_getter(master=self.root, app=self)  # Important, create frames from classes here

        try:
            con = sl.connect('Contact_List.db')
            zaqvka = "SELECT * FROM Contacts"
            data = con.execute(zaqvka)
            my_list = [r for r in data]
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)

        self.label_mail = Label(self.frame, text="E-Mail Address:")
        self.label_mail.grid(row=0, column=0, padx=5, sticky=E, ipady=3)

        self.entry_mail = Entry(self.frame)
        self.entry_mail.grid(row=0, column=1, padx=10, ipadx=150, ipady=3, sticky=W)

        self.label_pass = Label(self.frame, text="Password:")
        self.label_pass.grid(row=0, column=1, sticky=E)

        self.google_password = StringVar()
        self.google_password.set("") # you can insert your pass here
        self.entry_password = Entry(self.frame, show="*", textvariable=self.google_password)
        self.entry_password.grid(row=0, column=2, padx=10, ipadx=70, ipady=3, sticky=EW)

        self.label_smtp = Label(self.frame, text="SMTP Server:")
        self.label_smtp.grid(row=1, column=0, padx=5, sticky=W, pady=5, ipady=3)

        self.google_smtp = StringVar()
        self.google_smtp.set("smtp.gmail.com")
        self.entry_smtp = Entry(self.frame, textvariable=self.google_smtp)
        self.entry_smtp.grid(row=1, column=1, padx=10, ipadx=150, sticky=W, ipady=3)

        self.label_port = Label(self.frame, text="Port:")
        self.label_port.grid(row=1, column=1, sticky=E)

        self.google_port = StringVar()
        self.google_port.set(587)
        self.entry_port = Entry(self.frame, textvariable=self.google_port)
        self.entry_port.grid(row=1, column=2, padx=10, sticky=W, ipady=3)

        self.button_login = Button(self.frame, text="  Login  ", command=self.login)
        self.button_login.grid(row=1, column=1, ipadx=31, pady=5, columnspan=3, sticky=E)

        self.label_to = Label(self.frame, text="To:")
        self.label_to.grid(row=2, column=0, sticky=E, padx=10, pady=0)

        self.combo_to = ttk.Combobox(self.frame, values=my_list)
        self.combo_to.grid(row=2, column=1, padx=10, sticky=W, pady=10, ipadx=329, columnspan=3, ipady=3)
        self.combo_to.config(state="disabled")

        self.button_address_book = Button(self.frame, text="   To Contact List  ", command=self.go_to_Address_book)
        self.button_address_book.grid(row=2, column=1, padx=10, sticky=E, columnspan=3)
        self.button_address_book.config(state="disabled")

        self.label_subject = Label(self.frame, text="Subject:")
        self.label_subject.grid(row=3, column=0, sticky=SE, padx=10, pady=0)

        self.entry_subject = Entry(self.frame)
        self.entry_subject.grid(row=3, column=1, padx=10, sticky=W, ipadx=339, columnspan=3, ipady=3)
        self.entry_subject.config(state="disabled")

        self.button_attachments = Button(self.frame, text="Add Attachments", command=self.attach_to_mail)
        self.button_attachments.grid(row=3, column=1, padx=10, sticky=E, columnspan=3)
        self.button_attachments.config(state="disabled")

        self.label_mail_text = Label(self.frame, text="Mail Text:")
        self.label_mail_text.grid(row=4, column=1, pady=4, sticky=W)

        myfont = font.Font(family="Arial", size=16)
        self.my_text = Text(self.frame)
        self.my_text.configure(font=myfont)
        self.my_text.grid(row=5, column=1, sticky=EW, columnspan=2)
        self.my_text.config(state="disabled")

       # self.button_bold = Button(self.frame, text="Bold Text", command=self.do_bold)
       # self.button_bold.grid(row=5, column=0, sticky=N)

       # self.button_italic = Button(self.frame, text="Italic Text", command=self.do_italic)
       # self.button_italic.grid(row=5, column=0)

        self.label_attachments = Label(self.frame, text="Attachments:")
        self.label_attachments.grid(row=6, column=1, sticky=W, pady=10)

        self.send_button = Button(self.frame, text="Send Mail", command=self.send_mail)
        self.send_button.grid(row=7, column=1, sticky=NSEW, ipady=3, pady=5, padx=5, columnspan=2)
        self.send_button.config(state="disabled")

        self.get_mails = Button(self.frame, text="Get Mails", command=self.go_to_email_getter)
        self.get_mails.grid(row=8, column=1, sticky=EW, pady=10, columnspan=2, padx=5, ipady=3)
        self.get_mails.config(state="disabled")

        self.msg = MIMEMultipart()

    def main_page(self):
        self.frame.grid()
        self.update_form_main()
        self.combo_to.set("")

    def go_to_Address_book(self):
        self.frame.grid_forget()
        self.address_page.start_page()

    def go_to_email_getter(self):
        self.frame.grid_forget()
        self.mail_getter.start_page()

    def login(self):
        try:
            self.SMTP = self.entry_smtp.get()
            self.port = self.entry_port.get()
            self.mail = self.entry_mail.get()
            self.password = self.entry_password.get()

            self.server = smtplib.SMTP(self.SMTP, self.port)
            self.server.ehlo()
            self.server.starttls()
            self.server.ehlo()
            self.server.login(self.mail, self.password)
            self.entry_smtp.config(state="disabled")
            self.entry_mail.config(state="disabled")
            self.entry_password.config(state="disabled")
            self.entry_port.config(state="disabled")
            self.button_login.config(state="disabled")
            self.combo_to.config(state="normal")
            self.entry_subject.config(state="normal")
            self.my_text.config(state="normal")
            self.send_button.config(state="normal")
            self.button_address_book.config(state="normal")
            self.button_attachments.config(state="normal")
            self.get_mails.config(state="normal")

            tkinter.messagebox.showinfo(title="Success!", message="Connection Successful!")

        except smtplib.SMTPAuthenticationError:
            tkinter.messagebox.showerror(title="Connection Error", message="Wrong Credentials,\n Please try again")
        except:
            tkinter.messagebox.showerror(title="Error", message="Login Failed!")

    def get_from_combo(self):
        combo = self.combo_to.get()
        combo2 = combo.split(" ")
        self.final_combo = combo2[0]
        return self.final_combo

    def send_mail(self):
        answer = tkinter.messagebox.askyesno("Question", "Do you want to send the mail?")
        if answer == True:
            try:
                self.msg['From'] = self.entry_mail.get()
                self.msg['To'] = self.get_from_combo()
                self.msg['Subject'] = self.entry_subject.get()
                self.msg.attach(MIMEText(self.my_text.get("1.0", END), 'plain'))  # <----- problems,
                # solved --> textbox not getting info without args
                text_final = self.msg.as_string()
                self.server.sendmail(self.entry_mail.get(), self.get_from_combo(), text_final)
                tkinter.messagebox.showinfo(title="Success!", message="Your Email Have Been Sent!")
            except:
                tkinter.messagebox.showerror(title="Error!", message="Sending Mail Failed!\nCheck Receiver!")

    def attach_to_mail(self):
        filenames = tkinter.filedialog.askopenfilenames(title='Open Files', filetypes=[("All Files", "*.*")])
        if filenames:
            for filename in filenames:
                attachment = open(filename, 'rb')
                filename = filename[
                           filename.rfind("/") + 1:]  # !!!Could cause promblems if the file is with interval!!!
                p = MIMEBase('application', 'octet-stream')
                p.set_payload(attachment.read())
                encoders.encode_base64(p)
                p.add_header("Content-Disposition", f"attachment; filename={filename}")
                self.msg.attach(p)
                if not self.label_attachments.cget("text").endswith(":"):
                    self.label_attachments.config(text=self.label_attachments.cget("text") + ",")
                self.label_attachments.config(text=self.label_attachments.cget("text") + " " + filename)

    def do_bold(self):
        bold_font = font.Font(self.my_text, self.my_text.cget("font"))
        bold_font.configure(weight="bold")
        self.my_text.tag_configure("bold", font=bold_font)
        current_tags = self.my_text.tag_names("sel.first")
        if "bold" in current_tags:
            self.my_text.tag_remove("bold", "sel.first", "sel.last")
        else:
            self.my_text.tag_add("bold", "sel.first", "sel.last")

    def do_italic(self):
        italic_font = font.Font(self.my_text, self.my_text.cget("font"))
        italic_font.configure(slant="italic")
        self.my_text.tag_configure("italic", font=italic_font)
        current_tags = self.my_text.tag_names("sel.first")
        if "italic" in current_tags:
            self.my_text.tag_remove("italic", "sel.first", "sel.last")
        else:
            self.my_text.tag_add("italic", "sel.first", "sel.last")

    def update_form_main(self):
        try:
            con = sl.connect('Contact_List.db')
            zaqvka = "SELECT * FROM Contacts"
            data = con.execute(zaqvka)
            self.my_list_for_contacts = [r for r in data]
            self.combo_to['values'] = tuple(self.my_list_for_contacts)
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)


class Address_book:
    def __init__(self, master=None, app=None):
        self.master = master
        self.app = app
        self.frame = tk.Frame(self.master)

        self.label_mail_to_add = Label(self.frame, text="Mail:").grid(row=0, column=0)
        self.label_fname_to_add = Label(self.frame, text="First Name:").grid(row=0, column=1)
        self.label_lname_to_add = Label(self.frame, text="Last Name:").grid(row=0, column=2)

        self.m = StringVar()
        self.entry_mail_to_add = Entry(self.frame, textvariable=self.m).grid(row=1, column=0, padx=5)

        self.fn = StringVar()
        self.entry_fname_to_add = Entry(self.frame, textvariable=self.fn).grid(row=1, column=1, padx=5)

        self.ln = StringVar()
        self.entry_lname_to_add = Entry(self.frame, textvariable=self.ln).grid(row=1, column=2, padx=5)

        self.button_add_to_contacts = Button(self.frame, text="Add to Contacts", command=self.add_to_db)
        self.button_add_to_contacts.grid(row=2, column=1, pady=10)

        self.label_change_name = Label(self.frame, text="Please choose an email to change the name")
        self.label_change_name.grid(row=3, column=1)

        try:
            con = sl.connect('Contact_List.db')
            zaqvka = "SELECT * FROM Contacts"
            data = con.execute(zaqvka)
            self.my_list_for_contacts = [r for r in data]
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)

        self.combo_for_change = ttk.Combobox(self.frame, values=self.my_list_for_contacts)
        self.combo_for_change.grid(row=4, column=0, columnspan=3, sticky=EW)

        self.label_change_fname = Label(self.frame, text="Two new names:")
        self.label_change_fname.grid(row=5, column=0)

        self.fn_ch = StringVar()
        self.entry_fname_to_change = Entry(self.frame, textvariable=self.fn_ch)
        self.entry_fname_to_change.grid(row=5, column=1, pady=5, ipadx=20)

        self.ln_ch = StringVar()
        self.entry_lname_to_change = Entry(self.frame, textvariable=self.ln_ch)
        self.entry_lname_to_change.grid(row=5, column=2, sticky=W, ipadx=20)

        self.button_update_contacts = Button(self.frame, text="Update Names", command=self.update_db)
        self.button_update_contacts.grid(row=6, column=1, pady=10, ipadx=3)

        self.label_delete_mail = Label(self.frame, text="Select Contact To Delete")
        self.label_delete_mail.grid(row=7, column=1)

        self.combo_for_delete = ttk.Combobox(self.frame, values=self.my_list_for_contacts)
        self.combo_for_delete.grid(row=8, column=0, columnspan=3, sticky=EW)

        self.button_delete_contact = Button(self.frame, text="Delete Contact", command=self.delete_db)
        self.button_delete_contact.grid(row=9, column=1, pady=10, ipadx=3)

        self.button_back = Button(self.frame, text=" Back ", command=self.go_back).grid(row=10, column=0, columnspan=3,
                                                                                        sticky=NSEW)

    def start_page(self):
        self.frame.grid()

    def go_back(self):
        self.frame.grid_forget()
        self.app.main_page()

    def add_to_db(self):
        sql = 'INSERT INTO Contacts (Email,First_name,Last_name) VALUES (?,?,?)'
        mail_to_add = self.m.get().split(" ")
        fname_to_add = self.fn.get().split(" ")
        lname_to_add = self.ln.get().split(" ")
        data = [mail_to_add[0], fname_to_add[0], lname_to_add[0]]
        con = sl.connect('Contact_List.db')
        try:
            with con:
                con.execute(sql, data)
                tkinter.messagebox.showinfo(title="Success", message="Contact added to DB!")
                self.m.set("")
                self.fn.set("")
                self.ln.set("")
                con.commit()
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)
        self.update_form()

    def update_db(self):
        combo = self.combo_for_change.get().split(" ")
        new_combo = combo[0]
        last_combo = str(new_combo)
        fn_to_ch = self.fn_ch.get()
        ln_to_ch = self.ln_ch.get()
        newsql = 'UPDATE Contacts SET "First_name" = ?, Last_name = ? WHERE Email like ?'
        data = [fn_to_ch, ln_to_ch, last_combo]
        con = sl.connect('Contact_List.db')
        cur = con.cursor()
        try:
            with con:
                cur.execute(newsql, data)
                tkinter.messagebox.showinfo(title="Success", message="Contact Updated!")
                self.fn_ch.set("")
                self.ln_ch.set("")
                self.combo_for_change.set("")
                con.commit()
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)
        self.update_form()

    def delete_db(self):
        combo = self.combo_for_delete.get().split(" ")
        new_combo = combo[0]
        newsql2 = 'DELETE from Contacts WHERE "Email" like ?'
        data = new_combo
        con = sl.connect('Contact_List.db')
        cur = con.cursor()
        try:
            with con:
                cur.execute(newsql2, [data])
                tkinter.messagebox.showinfo(title="Success", message="Contact Deleted!")
                self.combo_for_delete.set("")
                con.commit()
        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)
        self.update_form()

    def update_form(self):
        try:
            con = sl.connect('Contact_List.db')
            zaqvka = "SELECT * FROM Contacts"
            data = con.execute(zaqvka)
            self.my_list_for_contacts = [r for r in data]
            self.combo_for_change['values'] = tuple(self.my_list_for_contacts)
            self.combo_for_delete['values'] = tuple(self.my_list_for_contacts)

        except sqlite3.Error as er:
            print('SQLite error: %s' % (' '.join(er.args)))
            print("Exception class is: ", er.__class__)


class Mail_getter:

    def __init__(self, master=None, app=None):
        self.master = master
        self.app = app
        self.frame = tk.Frame(self.master)

        self.get_mails_from_gmail = Button(self.frame, text="Download Mails", command=self.get_mails_from_gmail)
        self.get_mails_from_gmail.grid(row=0, column=0, padx=10, pady=10, sticky=N, ipadx=5)

        self.show_saved_mails = Button(self.frame, text="Show Saved Mails", command=self.from_db_to_treeview)
        self.show_saved_mails.grid(row=0, column=0)

        self.back_button = Button(self.frame, text="Go Back", command=self.go_back)
        self.back_button.grid(row=0, column=0, sticky=S, ipadx=25)

        style = ttk.Style()
        style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25, fieldbackground="#D3D3D3")
        style.map('Treeview', background=[('selected', "#347083")])

        columns = ('Email_From', 'Email_To', 'Subject', 'Date', 'Email_Body')
        self.treeview = ttk.Treeview(self.frame, columns=columns, show='headings')
        self.treeview.grid(row=0, column=1)
        self.treeview.heading('Email_From', text="Email From")
        self.treeview.heading('Email_To', text="Email To")
        self.treeview.heading('Subject', text="Subject")
        self.treeview.heading('Date', text="Date")
        self.treeview.heading('Email_Body', text="Mail Body")
        self.treeview.bind('<<TreeviewSelect>>', self.item_selected)
        self.treeview.tag_configure('oddrow', background="white")
        self.treeview.tag_configure('evenrow', background="lightblue")

    def start_page(self):
        self.frame.grid()

    def go_back(self):
        self.frame.grid_forget()
        self.app.main_page()

    def get_mails_from_gmail(self):
        self.server = "imap.gmail.com"
        self.username = "voodomanbg@gmail.com"
        self.password = self.app.password
        #self.password = "insert your pass here"
        self.new_con = sl.connect('Email_Container.db')
        sql_add = 'INSERT INTO Emails ("From", "To", "Subject", "Date", "Body") VALUES (?,?,?,?,?)'
        sql_remove = 'DELETE from Emails'

        with self.new_con:
            try:
                self.new_con.execute(sql_remove)
            except sqlite3.Error as er:
                print('SQLite error: %s' % (' '.join(er.args)))
                print("Exception class is: ", er.__class__)

        with Imbox(self.server, username=self.username, password=self.pasword) as imbox:
            all_inbox_messages = imbox.messages(unread=True, folder="Inbox")
            last_1_messages = all_inbox_messages[-10:]
            for uid, message in last_1_messages:
                email_from = message.sent_from[0]['email']  # to go to DB as FROM
                email_to = message.sent_to[0]['email']  # to go to DB as TO
                email_subject = message.subject  # to go to DB as SUBJECT
                email_date = message.date
                email_body = message.body['plain'][0]  # to go to DB as BODY , to illustrate as normal ---> email_body[0]
                data = [email_from, email_to, email_subject, email_date, str(email_body)]
                with self.new_con:
                    try:
                        self.new_con.execute(sql_add, data)
                    except sqlite3.Error as er:
                        print('SQLite error: %s' % (' '.join(er.args)))
                        print("Exception class is: ", er.__class__)
        tkinter.messagebox.showinfo(title="Success!", message="Emails Downloaded")

    def from_db_to_treeview(self):
        new_con2 = sl.connect('Email_Container.db')
        cur = new_con2.cursor()
        cur.execute("SELECT * FROM Emails")
        records = cur.fetchall()

        for i in self.treeview.get_children():
            self.treeview.delete(i)

        global count
        count = 0
        for record in records:
            if count % 2 == 0:
                self.treeview.insert(parent='', index='end', text='',
                                     values=(record[0], record[1], record[2], record[3], record[4]), tags='evenrow')
            else:
                self.treeview.insert(parent='', index='end', text='',
                                     values=(record[0], record[1], record[2], record[3], record[4]), tags='oddrow')
            count += 1
        tkinter.messagebox.showinfo(title="Success!", message="Treeview Updated!")

    def item_selected(self, event):  # on click gets the email_body
        for selected_item in self.treeview.selection():
            email_text = self.treeview.item(selected_item)
            record = email_text['values'][4]
            tkinter.messagebox.showinfo(title="Email Text:", message=record)


e = Main_Page(root)
root.mainloop()
