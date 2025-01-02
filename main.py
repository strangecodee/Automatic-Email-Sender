import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox


# Function to read email data from an Excel file
def read_email_data(file_path):
    try:
        df = pd.read_excel(file_path)
        print("Excel Data Columns:", df.columns)  # Debug: Print the columns
        print("Sample Data:\n", df.head())  # Debug: Print the first few rows
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel file: {e}")
        return None


# Function to send an email
def send_email(to_email, subject, message):
    from_email = 'annu.exe@gmail.com'  # Replace with your email
    from_password = 'efzb arey crxx wuzh'  # Replace with your Google App Password

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(from_email, from_password)
            server.send_message(msg)
        print(f"Email sent successfully to {to_email}")
    except Exception as e:
        print(f"Error sending email to {to_email}: {e}")


# Function to send all emails from the Excel file
def send_emails(file_path):
    email_data = read_email_data(file_path)
    required_columns = ['Name', 'Email', 'Subject', 'Message']

    if email_data is not None:
        # Check for missing columns
        missing_columns = [col for col in required_columns if col not in email_data.columns]
        if missing_columns:
            messagebox.showerror("Error", f"Missing columns in the Excel file: {', '.join(missing_columns)}")
            return

        for _, row in email_data.iterrows():
            try:
                personalized_message = row['Message'].replace("{Name}",
                                                              row['Name'])  # Replace {Name} with recipient's name
                send_email(row['Email'], row['Subject'], personalized_message)
            except KeyError as e:
                print(f"Error processing row: {e}")
        messagebox.showinfo("Success", "All emails have been sent!")
    else:
        messagebox.showwarning("Warning", "No email data found.")


# Function to browse and select an Excel file
def browse_file():
    file_path = filedialog.askopenfilename(
        title="Select an Excel File",
        filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
    )
    if file_path:
        file_label.config(text=f"Selected File: {file_path}")
        global selected_file
        selected_file = file_path


# Function to handle "Send Emails" button click
def handle_send_emails():
    if not selected_file:
        messagebox.showerror("Error", "No file selected. Please select an Excel file.")
        return
    send_emails(selected_file)


# Function to handle "Schedule Emails" button click
def handle_schedule_emails():
    global scheduler_thread
    if not selected_file:
        messagebox.showerror("Error", "No file selected. Please select an Excel file.")
        return

    send_time = time_entry.get()
    if not send_time:
        messagebox.showerror("Error", "Please enter a valid time (HH:MM).")
        return

    try:
        time.strptime(send_time, "%H:%M")  # Validate time format
        schedule.every().day.at(send_time).do(send_emails, file_path=selected_file)
        messagebox.showinfo("Scheduled", f"Emails scheduled daily at {send_time}.")
        if not scheduler_thread.is_alive():
            scheduler_thread.start()
    except ValueError:
        messagebox.showerror("Error", "Invalid time format. Use HH:MM (24-hour format).")


# Function to run the scheduler
def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)


# Initialize Tkinter
root = tk.Tk()
root.title("Automatic Email Sender")

# Global variable to store the selected file
selected_file = None

# Scheduler thread
scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)

# UI Elements
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(pady=10)

title_label = tk.Label(frame, text="Automatic Email Sender", font=("Arial", 16))
title_label.grid(row=0, column=0, columnspan=2, pady=10)

browse_button = tk.Button(frame, text="Browse File", command=browse_file, font=("Arial", 12))
browse_button.grid(row=1, column=0, padx=10, pady=5)

file_label = tk.Label(frame, text="No file selected", font=("Arial", 12))
file_label.grid(row=1, column=1, padx=10, pady=5)

time_label = tk.Label(frame, text="Schedule Time (HH:MM):", font=("Arial", 12))
time_label.grid(row=2, column=0, padx=10, pady=5)

time_entry = tk.Entry(frame, width=10, font=("Arial", 12))
time_entry.grid(row=2, column=1, padx=10, pady=5)

send_button = tk.Button(frame, text="Send Emails Now", command=handle_send_emails, bg="green", fg="white",
                        font=("Arial", 12))
send_button.grid(row=3, column=0, columnspan=2, pady=10)

schedule_button = tk.Button(frame, text="Schedule Emails", command=handle_schedule_emails, bg="blue", fg="white",
                            font=("Arial", 12))
schedule_button.grid(row=4, column=0, columnspan=2, pady=10)

# Run the Tkinter event loop
root.mainloop()
