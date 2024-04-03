BirthdayBot is an automated birthday email sender programmed to send personalized birthday greetings to recipients based on data stored in an Excel file. Simply provide the Excel file with birthday dates and email addresses, and let AutoBdaySender handle the rest.


**Part 1: Importing Libraries**

`import smtplib`: Utilizes the Simple Mail Transfer Protocol (SMTP) to send email messages over the internet by connecting to an SMTP server.

`from email.mime.text import MIMEText`: Employs `MIMEText` from the `email` module to create plain text email messages. This is part of the MIME (Multipurpose Internet Mail Extensions) standard, allowing emails to include text in various character sets and multimedia attachments.

`from email.mime.multipart import MIMEMultipart`: Uses `MIMEMultipart` for constructing emails with multiple parts, such as a combination of plain text and attachments. It enables the creation of complex email structures.

`import datetime`: Imports the `datetime` module to work with dates and times. This is crucial for determining the current date and comparing it with birthdays stored in an Excel file to identify today's birthdays.

`import pandas as pd`: Integrates the `pandas` library for advanced data manipulation and analysis. In this program, `pandas` is used to read and process the birthday list from an Excel file into a DataFrame, facilitating easy data handling and filtering based on dates and other criteria.



**Part 2: send_birthday_email function**

This section of code is a function called `send_birthday_email` that sends a birthday email to a specified recipient. Here's a breakdown of what each part does:

Setting SMTP Server and Credentials:
Defines the SMTP server (`smtp_server`) and port (`smtp_port`) for sending emails.
Specifies the sender's email address (`smtp_username`) and password (`smtp_password`). Ensure this password is an App Password if using Gmail.

Creating Email Message:
Creates a message container (`msg`) using `MIMEMultipart`, allowing for attachments if needed.
`MIME multipart` enables the bundling of various content types, like text, HTML, and attachments, into one email, allowing for complex messages with multiple parts and formats.
Sets the sender (`From`), recipient (`To`), and subject of the email (`Subject`).

Composing Email Content:
Composes the body of the email (`body`) with a personalized birthday message addressed to the recipient's first name.
 Attach the plain text body to the email message using `MIMEText`.

Sending Email:
Attempts to connect to the SMTP server using `smtplib.SMTP`.
Initiates a secure connection (`starttls`) to encrypt the communication with the server.
Logs in to the SMTP server using the provided username and password.
Sends the email message to the recipient using `server.sendmail`.
Prints a success message if the email is sent successfully; otherwise, prints an error message if an exception occurs during the process.


**Part 3 - get_birthday_from_excel**


This function, `get_birthdays_from_excel`, is designed to read an Excel file containing people's birthdays and filter out those whose birthday is today. Here’s what each part of the function does:

Reading the Excel File:
Uses `pd.read_excel` from the `pandas` library to read the Excel file specified by `filepath`. It employs the `engine='openpyxl'` argument, suitable for `.xlsx` files.

Processing the Birthday Column:
Converts the `birthday` column to datetime format using `pd.to_datetime`, with the format `'%m/%d'`. It ignores conversion errors (`errors='coerce'`), replacing invalid dates with `NaT` (Not a Time).
Removes rows where the birthday date couldn't be converted to datetime (`df.dropna(subset=['birthday'])`), filtering out missing or invalid dates.

Filtering Invalid Email Addresses:
Removes rows without valid email addresses by dropping rows where the email address is missing (`df.dropna(subset=['Email'])`).

Finding Today’s Birthdays:
Compares today’s date in the `'%m/%d'` format to the `birthday` column in the DataFrame. It filters the DataFrame to include only rows where the birthday matches today’s date.

Error Handling:
The `try` block attempts to execute file reading and filtering operations. If an error occurs (e.g., file not found, incorrect file format), it catches the exception, prints an error message, and returns an empty DataFrame. This prevents program crashes due to unexpected issues with the Excel file.

Return Value:
Returns a DataFrame containing rows for individuals whose birthdays match today’s date. If there's an error reading the file or no matches are found, it returns an empty DataFrame.


**Part 4- Main Function**


This section of the code defines the `main` function, which serves as the central point of execution for the birthday email sender program. Here’s a breakdown of its components:

Setting the Excel File Path:
`filepath` stores the path to the Excel file containing birthdays.

Reading Birthdays from the Excel File:
Calls `get_birthdays_from_excel(filepath)` to find today's birthdays.

Sending Birthday Emails:
If birthdays are found, send email greetings using `send_birthday_email`.
Handles cases where no birthdays are found.

Program Entry Point:
The `if __name__ == "__main__":` line ensures the script is executable as a standalone program.
