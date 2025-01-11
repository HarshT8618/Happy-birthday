import tkinter as tk
from tkinter import messagebox, Toplevel, filedialog
from datetime import datetime
from tkinter import ttk
import os
from PyPDF2 import PdfReader
from pygame import mixer
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from docx import Document
import smtplib

class BirthdayManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Birthday Management")
        self.root.geometry("1000x700")

        self.birthday_data = []
        self.music_file = "happy_birthday.mp3"  # Add your music file here

        mixer.init()

        self.setup_ui()

    def setup_ui(self):
        # Title
        self.title_label = tk.Label(self.root, text="BIRTHDAY MANAGEMENT", font=("Arial", 24, "bold"), fg="blue")
        self.title_label.pack(pady=20)
        
        # Create a main frame
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left frame for adding/removing birthdays
        self.left_frame = tk.Frame(self.main_frame, padx=20, pady=20, relief=tk.RIDGE, bd=2)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.name_label = tk.Label(self.left_frame, text="Name:", font=("Arial", 14))
        self.name_label.pack(pady=5)
        self.name_entry = tk.Entry(self.left_frame, font=("Arial", 14))
        self.name_entry.pack(pady=5)

        self.dob_label = tk.Label(self.left_frame, text="Date of Birth (YYYY-MM-DD):", font=("Arial", 14))
        self.dob_label.pack(pady=5)
        self.dob_entry = tk.Entry(self.left_frame, font=("Arial", 14))
        self.dob_entry.pack(pady=5)

        self.email_label = tk.Label(self.left_frame, text="Email:", font=("Arial", 14))
        self.email_label.pack(pady=5)
        self.email_entry = tk.Entry(self.left_frame, font=("Arial", 14))
        self.email_entry.pack(pady=5)

        self.add_button = tk.Button(self.left_frame, text="Add Birthday", font=("Arial", 14), command=self.add_birthday)
        self.add_button.pack(pady=10)
        
        self.remove_button = tk.Button(self.left_frame, text="Remove Birthday", font=("Arial", 14), command=self.remove_birthday)
        self.remove_button.pack(pady=10)

        self.upload_button = tk.Button(self.left_frame, text="Upload DOCX/PDF", font=("Arial", 14), command=self.upload_file)
        self.upload_button.pack(pady=10)

        # Right frame for displaying/showing birthdays
        self.right_frame = tk.Frame(self.main_frame, padx=20, pady=20, relief=tk.RIDGE, bd=2)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.show_button = tk.Button(self.right_frame, text="Show All Birthdays", font=("Arial", 14), command=self.show_birthdays)
        self.show_button.pack(pady=10)

        self.check_button = tk.Button(self.right_frame, text="Check Today's Birthdays", font=("Arial", 14), command=self.check_todays_birthdays)
        self.check_button.pack(pady=10)

    def add_birthday(self):
        name = self.name_entry.get()
        dob = self.dob_entry.get()
        email = self.email_entry.get()
        
        if not name or not dob or not email:
            messagebox.showwarning("Input Error", "Please provide name, date of birth, and email.")
            return
        
        try:
            datetime.strptime(dob, '%Y-%m-%d')
        except ValueError:
            messagebox.showwarning("Input Error", "Invalid date format. Please use YYYY-MM-DD.")
            return
        
        self.birthday_data.append((name, dob, email))
        self.name_entry.delete(0, tk.END)
        self.dob_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        messagebox.showinfo("Success", f"Birthday for {name} added!")

        # Check if today is the person's birthday
        if self.is_today_birthday(dob, datetime.today()):
            self.send_birthday_email(name, email)
            self.play_music()
            self.show_party_popper()

    def remove_birthday(self):
        name = self.name_entry.get()
        
        if not name:
            messagebox.showwarning("Input Error", "Please provide the name.")
            return
        
        self.birthday_data = [bd for bd in self.birthday_data if bd[0] != name]
        messagebox.showinfo("Success", f"Birthday for {name} removed!")
        
    def show_birthdays(self):
        show_window = Toplevel(self.root)
        show_window.title("All Birthdays")
        show_window.geometry("600x400")
        
        tree = ttk.Treeview(show_window, columns=("Name", "DOB", "Email", "Age"), show='headings')
        tree.heading("Name", text="Name")
        tree.heading("DOB", text="Date of Birth")
        tree.heading("Email", text="Email")
        tree.heading("Age", text="Age")
        tree.pack(pady=20, fill="both", expand=True)

        for name, dob, email in self.birthday_data:
            age = self.calculate_age(dob)
            tree.insert("", "end", values=(name, dob, email, f"{age['years']} years, {age['months']} months, {age['days']} days"))
        
    def check_todays_birthdays(self):
        today = datetime.today()
        today_birthdays = [bd for bd in self.birthday_data if self.is_today_birthday(bd[1], today)]
        
        if today_birthdays:
            birthday_names = ", ".join(bd[0] for bd in today_birthdays)
            for bd in today_birthdays:
                self.send_birthday_email(bd[0], bd[2])
            messagebox.showinfo("Today's Birthdays", f"Happy Birthday to: {birthday_names}")
            self.play_music()
            self.show_party_popper()
        else:
            messagebox.showinfo("Today's Birthdays", "No birthdays today.")
    
    def is_today_birthday(self, dob, today):
        dob_date = datetime.strptime(dob, '%Y-%m-%d')
        return dob_date.month == today.month and dob_date.day == today.day
        
    def play_music(self):
        if os.path.exists(self.music_file):
            mixer.music.load(self.music_file)
            mixer.music.play()
        else:
            messagebox.showerror("Music Error", "Music file not found.")
    
    def show_party_popper(self):
        popper_window = Toplevel(self.root)
        popper_window.geometry("200x200")
        popper_window.title("🎉")
        popper_label = tk.Label(popper_window, text="🎉", font=("Arial", 100))
        popper_label.pack(expand=True)
        popper_window.after(3000, popper_window.destroy)  # Close the popper window after 3 seconds
    
    def calculate_age(self, birth_date):
        today = datetime.today()
        birth = datetime.strptime(birth_date, '%Y-%m-%d')
        years = today.year - birth.year
        months = today.month - birth.month
        days = today.day - birth.day

        if days < 0:
            months -= 1
            days += (today.replace(month=today.month - 1, day=1) - today.replace(day=1)).days

        if months < 0:
            years -= 1
            months += 12

        return {"years": years, "months": months, "days": days}
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("DOCX files", "*.docx"), ("PDF files", "*.pdf")])
        
        if not file_path:
            return
        
        if file_path.endswith('.docx'):
            self.read_docx(file_path)
        elif file_path.endswith('.pdf'):
            self.read_pdf(file_path)
        else:
            messagebox.showerror("File Error", "Unsupported file format.")
    
    def read_docx(self, file_path):
        try:
            document = Document(file_path)
            for paragraph in document.paragraphs:
                line = paragraph.text.strip()
                if line:
                    parts = line.split()
                    if len(parts) >= 3:
                        name = " ".join(parts[:-2])
                        dob = parts[-2]
                        email = parts[-1]
                        try:
                            datetime.strptime(dob, '%Y-%m-%d')
                            self.birthday_data.append((name, dob, email))
                        except ValueError:
                            continue
            messagebox.showinfo("Success", "Birthdays added from DOCX!")
        except Exception as e:
            messagebox.showerror("DOCX Error", f"Error reading DOCX file: {e}")
    
    def read_pdf(self, file_path):
        try:
            reader = PdfReader(file_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
            
            lines = text.splitlines()
            for line in lines:
                if line:
                    parts = line.split()
                    if len(parts) >= 3:
                        name = " ".join(parts[:-2])
                        dob = parts[-2]
                        email = parts[-1]
                        try:
                            datetime.strptime(dob, '%Y-%m-%d')
                            self.birthday_data.append((name, dob, email))
                        except ValueError:
                            continue
            messagebox.showinfo("Success", "Birthdays added from PDF!")
        except Exception as e:
            messagebox.showerror("PDF Error", f"Error reading PDF file: {e}")

    def send_birthday_email(self, name, email):
        sender_email = "harshtolagatti@gmail.com"
        sender_password = "bvxm qlmb lwgi jcvh"
        
        subject = "Happy Birthday!"
        body = f"Dear {name},\n\nWishing you a fantastic birthday filled with joy and happiness!\n\nBest regards,\nBirthday Management Team"

        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = email
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))

        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(message)
            messagebox.showinfo("Email Sent", f"Birthday email sent to {name}!")
        except Exception as e:
            messagebox.showerror("Email Error", f"Failed to send email: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = BirthdayManager(root)
    root.mainloop()
