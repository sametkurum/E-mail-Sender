# E-mail sender

import smtplib
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Sender informations
sender_email = ""
sender_name = ""
password = ""

# receivers' names and e-mails. --> ("Name Surmane", "anemail@gmail.com")
recipients = []

# Functions
def send_email():
    subject = subject_entry.get()
    body = body_entry.get("1.0", tk.END).strip()
    
    if not subject or not body:
        messagebox.showerror("Error", "Subject and body can't empty.")
        return
    
    if not recipients:
        messagebox.showerror("Error", "Recipients list is empty.")
        return

    # Connection SMTP server
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, password)
        
        progress_text.set("Sending started...")
        for i, (name, email) in enumerate(recipients):
            msg = MIMEMultipart()
            msg['From'] = f"{sender_name} <{sender_email}>"
            msg['To'] = f"{name} <{email}>"
            msg['Subject'] = Header(subject, 'utf-8').encode()
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Dosya eklerini ekle
            for file_path in attachments:
                part = MIMEBase('application', 'octet-stream')
                with open(file_path, 'rb') as file:
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
                msg.attach(part)

            server.sendmail(sender_email, email, msg.as_string())
            progress_var.set((i + 1) / len(recipients) * 100)
            progress_text.set(f"Sending: {email}")
            root.update_idletasks()
        
        progress_text.set(f"Mission completed! Receiver mails: {', '.join(email for _, email in recipients)}")
        
    except smtplib.SMTPRecipientsRefused as e:
        progress_text.set(f"Unsuccessfully: {e}")
    except Exception as e:
        progress_text.set(f"Error: {e}")
    finally:
        server.quit()

def update_sender_info():
    global sender_email, sender_name, password
    sender_email = sender_email_entry.get()
    sender_name = sender_name_entry.get()
    password = password_entry.get()

def update_recipients_info():
    global recipients
    recipients = []
    for entry in recipient_entries:
        text = entry.get()
        if ':' in text:
            name, email = text.split(':', 1)
            recipients.append((name.strip(), email.strip()))

def add_recipient():
    global recipient_entries
    entry = tk.Entry(recipients_window, width=50)
    entry.grid(row=len(recipient_entries), column=0, padx=10, pady=5, sticky='ew')
    recipient_entries.append(entry)
    recipients_window.update_idletasks()

def attach_file():
    file_paths = filedialog.askopenfilenames(title="Choose file", filetypes=[("All docs", "*.*")])
    for file_path in file_paths:
        attachments.append(file_path)
        attachment_listbox.insert(tk.END, file_path)

def open_sender_settings():
    global sender_window, sender_email_entry, sender_name_entry, password_entry
    if sender_window is None or not tk.Toplevel.winfo_exists(sender_window):
        sender_window = tk.Toplevel(root)
        sender_window.title("Sender settings")

        tk.Label(sender_window, text="Sender Email:").grid(row=0, column=0, padx=10, pady=5)
        sender_email_entry = tk.Entry(sender_window, width=50)
        sender_email_entry.insert(0, sender_email)
        sender_email_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(sender_window, text="Sender name:").grid(row=1, column=0, padx=10, pady=5)
        sender_name_entry = tk.Entry(sender_window, width=50)
        sender_name_entry.insert(0, sender_name)
        sender_name_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(sender_window, text="Password:").grid(row=2, column=0, padx=10, pady=5)
        password_entry = tk.Entry(sender_window, show="*", width=50)
        password_entry.insert(0, password)
        password_entry.grid(row=2, column=1, padx=10, pady=5)

        tk.Button(sender_window, text="Update", command=update_sender_info).grid(row=3, column=1, padx=10, pady=10, sticky="e")

def open_recipients_settings():
    global recipients_window, recipient_entries
    if recipients_window is None or not tk.Toplevel.winfo_exists(recipients_window):
        recipients_window = tk.Toplevel(root)
        recipients_window.title("Receiver settings")
        
        recipient_entries = []
        
        for i, (name, email) in enumerate(recipients):
            entry = tk.Entry(recipients_window, width=50)
            entry.insert(0, f"{name}: {email}")
            entry.grid(row=i, column=0, padx=10, pady=5, sticky='ew')
            recipient_entries.append(entry)
        
        tk.Button(recipients_window, text="Add new receiver", command=add_recipient).grid(row=len(recipient_entries), column=0, padx=10, pady=10, sticky="e")
        tk.Button(recipients_window, text="Update", command=update_recipients_info).grid(row=len(recipient_entries)+1, column=0, padx=10, pady=10, sticky="e")

# Main interface
root = tk.Tk()
root.title("Email sender program")
icon = tk.PhotoImage(file= 'my_logo.png')
root.iconphoto(True,icon)

# Creating Menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

settings_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Settings", menu=settings_menu)
settings_menu.add_command(label="Sender informations", command=open_sender_settings)
settings_menu.add_command(label="Receiver informations", command=open_recipients_settings)

# Adding receiver and files
attachments = []

tk.Label(root, text="Subject:").grid(row=0, column=0, padx=10, pady=5)
subject_entry = tk.Entry(root, width=50)
subject_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Body:").grid(row=1, column=0, padx=10, pady=5)
body_entry = tk.Text(root, width=50, height=10)
body_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Attachments:").grid(row=2, column=0, padx=10, pady=5)
attachment_listbox = tk.Listbox(root, width=50, height=5)
attachment_listbox.grid(row=2, column=1, padx=10, pady=5)

attach_button = tk.Button(root, text="Add file", command=attach_file)
attach_button.grid(row=3, column=1, padx=10, pady=5, sticky="e")

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

progress_text = tk.StringVar()
progress_text.set("Ready")
progress_label = tk.Label(root, textvariable=progress_text)
progress_label.grid(row=5, column=1, padx=10, pady=5)

send_button = tk.Button(root, text="Send it", command=send_email)
send_button.grid(row=6, column=1, padx=10, pady=10, sticky="e")

# Setting windows and global variables
sender_window = None
recipients_window = None
recipient_entries = []

root.mainloop()
