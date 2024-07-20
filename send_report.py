import os
from simple_salesforce import Salesforce
import pandas as pd
import requests
from io import StringIO
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_email_report():
    # Connect to Salesforce using environment variables
    sf = Salesforce(username=os.getenv('SF_USERNAME'), password=os.getenv('SF_PASSWORD'), security_token=os.getenv('SF_SECURITY_TOKEN'))

    # Salesforce report details
    sf_instance = 'https://project-lithium.my.salesforce.com/'
    report_Id = '00OC5000000KvUXMA0'
    export = '?isdtp=p1&export=1&enc=UTF-8&xf=csv'
    sfUrl = sf_instance + report_Id + export

    # Download the report
    response = requests.get(sfUrl, headers=sf.headers, cookies={'sid': sf.session_id})
    download_report = response.content.decode('utf-8')

    # Load the data into a DataFrame
    df = pd.read_csv(StringIO(download_report))

    # Convert 'Check In' and 'Check Out' to datetime format
    df['Check In'] = pd.to_datetime(df['Check In'], format='%d/%m/%Y, %I:%M %p', errors='coerce')
    df['Check Out'] = pd.to_datetime(df['Check Out'], format='%d/%m/%Y, %I:%M %p', errors='coerce')

    # Function to replace blank entries with specified strings and format time only
    def replace_blank_entries(row):
        if pd.isnull(row['Check In']):
            row['Check In'] = "Not checked in"
        else:
            row['Check In'] = row['Check In'].strftime('%I:%M %p')  # Format to display time only
        if pd.isnull(row['Check Out']):
            row['Check Out'] = "Not checked out"
        else:
            row['Check Out'] = row['Check Out'].strftime('%I:%M %p')  # Format to display time only
        return row

    # Apply the function to update the DataFrame
    df = df.apply(replace_blank_entries, axis=1)

    # Function to calculate working hours from datetime objects or time only strings
    def calculate_working_hours(check_in_dt, check_out_dt):
        if check_in_dt == "Not checked in" or check_out_dt == "Not checked out":
            return None

        # Handle cases where only time is provided without date
        if isinstance(check_in_dt, datetime) and isinstance(check_out_dt, datetime):
            # Ensure check_out_dt is greater than check_in_dt
            if check_out_dt < check_in_dt:
                check_out_dt += timedelta(days=1)  # add one day to check_out_dt if it's less than check_in_dt

            return (check_out_dt - check_in_dt).total_seconds() / 3600
        elif isinstance(check_in_dt, str) and isinstance(check_out_dt, str):
            # Parse time strings and calculate difference
            check_in_time = datetime.strptime(check_in_dt, '%I:%M %p').time()
            check_out_time = datetime.strptime(check_out_dt, '%I:%M %p').time()

            # Adjust check_out_time if it's earlier in the day than check_in_time
            if check_out_time < check_in_time:
                return ((timedelta(hours=24) - timedelta(hours=check_in_time.hour)) + timedelta(hours=check_out_time.hour)).total_seconds() / 3600
            else:
                return (timedelta(hours=check_out_time.hour) - timedelta(hours=check_in_time.hour)).total_seconds() / 3600
        else:
            return None

    # Calculate working hours
    df['Working Hours'] = df.apply(lambda row: calculate_working_hours(row['Check In'], row['Check Out']), axis=1)

    # Round 'Working Hours' to two decimal places, handle None values
    df['Working Hours'] = df['Working Hours'].apply(lambda x: round(x, 2) if pd.notnull(x) else None)

    # Select the required columns for the report
    report = df[['Lithium ID', 'Supervisor Name', 'Attendance Date', 'Primary Campus', 'City', 'Check In', 'Check Out', 'Working Hours']]

    # Save the report to a CSV file
    report_filename = 'site_manager_working_hours_report.csv'
    report.to_csv(report_filename, index=False)

    # Email settings
    email_from = 'kartik@project-lithium.com'
    email_to = 'nithya@project-lithium.com'
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_user = os.getenv('SMTP_USER')
    smtp_password = os.getenv('SMTP_PASSWORD')

    # Compose the email
    msg = MIMEMultipart()
    msg['From'] = email_from
    msg['To'] = email_to

    # Subject with specific date format
    yesterday_date = (datetime.today() - timedelta(days=1)).strftime('%d-%m-%Y')  # Get yesterday's date in dd-mm-yyyy format
    msg['Subject'] = f'Site Manager Check-In/Check-Out Report for {yesterday_date}'

    body = """
    Hello,

    Please find the attached Site Manager and Supervisor Check-In/Check-Out Report.

    Best regards,
    Kartik Pandey
    """

    msg.attach(MIMEText(body, 'plain'))

    # Attach the report file
    attachment = open(report_filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {report_filename}')
    msg.attach(part)

    # Send the email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        text = msg.as_string()
        server.sendmail(email_from, email_to, text)
        print(f"Email with the report '{report_filename}' sent successfully to {email_to}")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        if 'server' in locals():
            server.quit()

if __name__ == "__main__":
    send_email_report()
