# Import necessary libraries and modules
import pandas as pd
import os
import shutil 
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Define source, destination, and filtered files directories, Path to the Excel file containing names and email addresses
source_folder = 'D:/test/source'
destination_folder = 'D:/test/destination'
filtered_files = 'D:/test/filtered/'
search_file_path = 'D:/test/Emails.xlsx'
search_df = pd.read_excel(search_file_path)

# Sends an email with an attached file to a specified recipient
def send_email(recipient, file_path):
    '''
    Send an email with an attached file to the specified recipient.
    Parameters:
        recipient (str): Email address of the recipient.
        file_path (str): Path to the file to be attached.
    '''
    # Email configuration
    sender_email = "*******@gmail.com"     # Replace with your email address
    sender_password = "*************"       # Replace with your email password
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # Create message container
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = "Daily Report"

    # Attach the filtered Excel file
    attachment = open(file_path, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(file_path))
    msg.attach(part)

    # Connect to SMTP server and send email
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    text = msg.as_string()
    server.sendmail(sender_email, recipient, text)
    server.quit()

#Iterate through files in the source folder
for file in os.listdir(source_folder):
    today_date = datetime.now().date()
    # Process files ending with 'Report.xlsx'
    if file.endswith('Report.xlsx'):
        file_path = os.path.join(source_folder, file)    
        try:  
            # Read Excel file into a DataFrame
            df1 = pd.read_excel(file_path, engine='openpyxl', skiprows=1)
            print(f'{file} is successfully opened')
            print('------------------------------------------------------------')
            
            # Iterate through names and emails in the search DataFrame
            for name, email in zip(search_df['Names'], search_df['Emails']):
                # Filter DataFrame based on 'Management Level 2' column
                filtered_df1 = df1[df1['Management Level 2'].isin(['مدير الادارة  2', name])]
                
                # Format percentage columns
                for col in filtered_df1.columns:
                    if isinstance(col, str) and "%" in col:
                        filtered_df1[col].iloc[1::] = (filtered_df1[col].iloc[1::] * 100).map("{:.1f}%".format)
                        filtered_df1[col].iloc[1::] = filtered_df1[col].iloc[1::].replace('nan%', '0.0%')
                
                # Check if filtered DataFrame has data
                if len(filtered_df1) > 2:
                    # Save filtered DataFrame to Excel file
                    filtered_df1.to_excel(f"{filtered_files}{name} {today_date}.xlsx", index=False)
                    
                    # Send email with filtered Excel file attached
                    filtered_excel_path = f"{filtered_files}{name} {today_date}.xlsx"
                    send_email(email, filtered_excel_path)
                    
                    print(f"Filtered data with '{name}' has been saved to 'D:/test/filtered'.")
                    print('------------------------------------------------------------')
            
            # Move processed file to destination folder
            shutil.move(file_path, os.path.join(destination_folder, file))
            print(f"{file} has been transferred to {destination_folder}")
            print('------------------------------------------------------------')
        except Exception as e:
            print(f"Error processing {file}: {e}")
            
# Process files ending with 'Final.xlsx'
    if file.endswith('Final.xlsx'):
        file_path = os.path.join(source_folder, file)    
        try:  
            # Read Excel file into a DataFrame
            df2 = pd.read_excel(file_path, engine='openpyxl')
            print(f'{file} is successfully opened')
            print('------------------------------------------------------------')
            
            # Iterate through names and emails in the search DataFrame
            for name, email in zip(search_df['Names'], search_df['Emails']):
                # Filter DataFrame based on 'Management Level 2' column
                filtered_df2 = df2[df2['Management Level 3'].isin([' Level 3', name])]
                
                # Format percentage columns
                for col in filtered_df2.columns:
                    if isinstance(col, str) and "%" in col:
                        filtered_df2[col].iloc[1::] = (filtered_df2[col].iloc[1::] * 100).map("{:.1f}%".format)
                        filtered_df2[col].iloc[1::] = filtered_df2[col].iloc[1::].replace('nan%', '0.0%')
                
                # Check if filtered DataFrame has data
                if len(filtered_df2) > 2:
                    # Save filtered DataFrame to Excel file
                    filtered_df2.to_excel(f"{filtered_files}{name} {today_date}.xlsx", index=False)
                    
                    # Send email with filtered Excel file attached
                    filtered_excel_path = f"{filtered_files}{name} {today_date}.xlsx"
                    send_email(email, filtered_excel_path)
                    
                    print(f"Filtered data with '{name}' has been saved to 'D:/test/filtered'.")
                    print('------------------------------------------------------------')
            
            # Move processed file to destination folder
            shutil.move(file_path, os.path.join(destination_folder, file))
            print(f"{file} has been transferred to {destination_folder}")
            print('------------------------------------------------------------')
        except Exception as e:
            print(f"Error processing {file}: {e}")
