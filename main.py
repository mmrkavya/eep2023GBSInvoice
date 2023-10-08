import schedule
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
from datetime import datetime, timedelta
from cryptography.fernet import Fernet

# Specify the path to your Excel file
excel_file_path = "C:\\Users\\KAVYA\\Documents\\eep\\notificationList.xlsx"
cipher_suite = Fernet(b'NGmSpKhwgdYIe_zRV1UPxVJP-JRRbqb396gTYpBDf3c=')

def job(recipient_email,name,accountNumber):
    print("I'm working...")

    # Email configuration
    sender_email = 'mmrkavya@outlook.com'
    #encrypted password
    sender_password = b'gAAAAABlIpjpeR7eZZD_zoUutoGBep5ilWTM7QsSmLC04i2gI9fZD7a1a_XA5RQR6AOBQiIELxYrFd2xNhgkahQLKrf0siWmrA=='
    subject = 'Notification of invoice deposited in HSBCnet'
    message = f'Hello {name}, Invoice for {accountNumber} has been deposited in HSBCNet.Thanks & Regards,HSBC'

    # Create a MIME message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    # Connect to the SMTP server and send the email
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, cipher_suite.decrypt(sender_password).decode())
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        print('Email sent successfully!')
    except Exception as e:
        print(f'Error: {str(e)}')


def sendNotfication():
# Read the Excel sheet into a DataFrame with the first row as headers
    df = pd.read_excel(excel_file_path, header=0)

    current_date =datetime.today().strftime('%Y-%m-%d %H:%M:%S')


    # Access the specific cell value at the specified row and column
    for index, row in df.iterrows():
        name = row["Name"]
        dateOfLastInvoice = row["date of last invoice depoisted"]
        lastInvoiceNumber= row["last invoice number"]
        accountNumber=row["account number"]
        email=row["email"]
        phoneNumber = row["phone no"]
        noticiationIndicator= row["notification"]
        given_date = dateOfLastInvoice

        # Get the current date and time
        current_date = datetime.now()

        # Calculate the time difference between the given date and the current date
        time_difference = current_date - given_date

        # Define the maximum allowed time difference (24 hours)
        max_time_difference = timedelta(hours=24)

        # Check if the time difference is less than or equal to 24 hours
        if time_difference <= max_time_difference:
            print("The given date is within 24 hours of now.")
            job(email,name,accountNumber)
        else:
            print("The given date is not within 24 hours of now.")
            print (dateOfLastInvoice)


schedule.every(1).minutes.do(sendNotfication)
#schedule.every().hour.do(job)
##schedule.every().day.at("10:30").do(job)

while 1:
    schedule.run_pending()
    time.sleep(1)