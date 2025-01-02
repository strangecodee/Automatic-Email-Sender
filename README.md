# 🌟 **Automatic Email Sender** 🌟

Welcome to the **Automatic Email Sender**, your new best friend for sending personalized emails at the click of a button (or with a scheduled plan)! Whether you're sending birthday wishes, business updates, or newsletters, this tool automates the entire process with ease. Forget about repetitive email tasks – let the machine handle it! 🎉

## 🚀 **Features**
- 📥 **Excel File Integration**: Read email data (Name, Email, Subject, Message) directly from an Excel file.
- 💌 **Personalized Messaging**: The {Name} in the message will automatically be replaced with the recipient’s name for a personalized touch.
- ✉️ **Gmail Support**: Sends emails using Gmail’s SMTP server (requires an app password).
- ⏰ **Scheduling Made Simple**: Schedule emails to be sent daily at any time you prefer (e.g., 8:00 AM for your daily reminders).
- 🖥️ **User-Friendly GUI**: With a sleek Tkinter-based interface, browse files, set up email schedules, and hit "Send" all in one place.

## 🛠️ **Prerequisites**
To get started, make sure your machine has the following libraries installed:
- 📦 **pandas**: For reading the email data from Excel.
- 📦 **smtplib**: For sending emails.
- 📦 **schedule**: To schedule emails at specific times.
- 🖼️ **tkinter**: For the beautiful graphical interface.

You can install the required dependencies using `pip`:

```bash
pip install pandas schedule
```

## 🚦 **Getting Started** 

### 1. **Clone or Download the Repository**
Clone the repository to your local machine or download the zip file with the project code.

```bash
git clone https://github.com/strangecodee/automatic-email-sender.git
```

### 2. **Install Dependencies**
In the project folder, open your terminal or command prompt and install the necessary libraries:

```bash
pip install pandas schedule
```

### 3. **Set Up Your Gmail Credentials**
Before sending emails, make sure you’ve got your Gmail credentials ready:
- Open the Python script and replace these placeholders:
  - `from_email`: Your Gmail address (e.g., `your-email@gmail.com`).
  - `from_password`: Your **Google App password** (for added security). Here's how to generate one:
    1. Visit your [Google account settings](https://myaccount.google.com/).
    2. Enable **2-Step Verification** under the **Security** tab.
    3. Generate a new **App password** under **App passwords** and use that here.

### 4. **Run the Program**
Now you're ready to go! Launch the program with:

```bash
python main.py
```

This will open the **Automatic Email Sender GUI** for you to interact with.

## 🎨 **How to Use**

### 1. **Browse for Your Excel File** 📂
Click the **Browse File** button to open the file dialog. Select your Excel file that contains the columns: `Name`, `Email`, `Subject`, and `Message`.

**Example Excel File:**

| Name        | Email               | Subject           | Message                                      |
|-------------|---------------------|-------------------|----------------------------------------------|
| John Doe    | john@example.com     | Hello John        | Hello {Name}, here's your update!           |
| Jane Smith  | jane@example.com     | Greetings Jane    | Hi {Name}, check out this info.             |
| Michael Lee | michael@example.com  | Your Invoice      | Dear {Name}, please find your invoice here.  |
| Emily Clark | emily@example.com    | Event Invitation  | Hi {Name}, you're invited to our event!     |
| Alex Brown  | alex@example.com     | Product Update    | Hello {Name}, we have an exciting new update for you! |
| Sara White  | sara@example.com     | Newsletter        | Dear {Name}, here's your monthly newsletter. |
| Chris James | chris@example.com    | Special Offer     | Hi {Name}, check out this special offer just for you! |
| Olivia Green| olivia@example.com   | Weekly Report     | Hello {Name}, please find your weekly report attached. |
| Daniel Black| daniel@example.com   | Reminder: Meeting | Hi {Name}, this is a reminder for our meeting tomorrow. |
| Sophia Gray | sophia@example.com   | Thank You         | Dear {Name}, thank you for your continued support. |

### 2. **Send Emails Now** 📤
Click **Send Emails Now** to instantly send emails to everyone in your Excel list. The message will be personalized with each recipient's name.

### 3. **Schedule Emails** ⏳
Want to automate the process? Schedule the emails to send daily at a specific time:
- Enter the time in the **Schedule Time (HH:MM)** field.
- Click **Schedule Emails** to have the emails sent every day at that time.

### 4. **Sit Back and Relax** 🌴
Once everything is set up, let the program work its magic! The emails will be sent as per your instructions, so you can focus on other important tasks.

## 📝 **Notes**
- 🚨 **Security**: Ensure your Gmail account allows less secure apps, or use an **App password** for improved security.
- 🖥️ **GUI**: The program opens a graphical interface using Tkinter, making it easy for you to manage and send emails.
- 📅 **Scheduling**: Emails will be sent daily at the time you set, so you don’t have to worry about it.

## 💬 **Contribute**
Feel free to fork this project, make improvements, and submit pull requests! Let's make emailing as fun and easy as possible. 🎉

## 🏁 **Done!**
You’re all set to start sending personalized emails in a snap! 🚀
