import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
import pandas as pd

# Function to send birthday email
def send_birthday_email(to_email, first_name):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'yourGmail.com'
    smtp_password = 'yourPassword'  # Update with your email password

    # Create message container
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = to_email
    msg['Subject'] = 'Happy Birthday!'

    # Compose the message
    body = f'Dear {first_name},\n\nHappy Birthday! Wishing you a fantastic day filled with joy and laughter.\n\nBest regards,\nyourName'
    msg.attach(MIMEText(body, 'plain'))

    # Connect to SMTP server and send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_username, to_email, msg.as_string())
            print(f"Sent birthday wishes to {to_email}")
    except Exception as e:
        print(f"Failed to send birthday wishes to {to_email}: {e}")

def get_birthdays_from_excel(filepath):
    try:
        # Read the Excel file
        df = pd.read_excel(filepath, engine='openpyxl')

        # Convert birthday column to datetime, ignoring errors
        df['birthday'] = pd.to_datetime(df['birthday'], format='%m/%d', errors='coerce')

        # Filter out any rows with missing or invalid birthday dates
        df = df.dropna(subset=['birthday'])

        # Filter out any rows with missing or invalid email addresses
        df = df.dropna(subset=['Email'])

        # Convert today's date to the same format for comparison
        today = datetime.datetime.now().strftime('%m/%d')

        # Filter rows where birthday matches today's date
        today_birthdays = df[df['birthday'].dt.strftime('%m/%d') == today]

        return today_birthdays
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()  # Return an empty DataFrame if an error occurs

def main():
    # Path to the Excel file
    filepath = 'yourFilePath'  # Update with the path to your Excel file

    # Read birthdays from the Excel file
    today_birthdays = get_birthdays_from_excel(filepath)

    if not today_birthdays.empty:
        # Send an email to each person having a birthday today
        for _, row in today_birthdays.iterrows():
            send_birthday_email(row['Email'], row['First Name'])  # Adjust column name for first name if needed
    else:
        print("No birthdays found for today.")

if __name__ == "__main__":
    main()
