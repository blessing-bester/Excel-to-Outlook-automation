import pandas as pd
import os
import win32com.client as win32  # Outlook automation
from pathlib import Path


EXCEL_FILE_PATH = r"C:\Documents\Master_file\master_amortization.xlsx"
EMAIL_LIST_SHEET = r"C:\Documents\Master_file\customer_Emails.xlsx"
CC_LIST = ["accounts@xxxxx.co.zw", "sales@xxxxxx.co.zw"]  
BCC_LIST = ["archive@xxxxx.co.zw"]

def get_customer_email_map():
    """Reading the Excel sheet that maps customer names to emails."""
    df_emails = pd.read_excel(EXCEL_FILE_PATH, sheet_name=EMAIL_LIST_SHEET)
    # Converting the DataFrame to a dictionary
    email_dict = dict(zip(df_emails['Name'], df_emails['Email']))
    return email_dict

def save_customer_sheet(customer_name):
    """Saving a single customer's worksheet as a separate Excel file."""
    # Reading a specific sheet without loading the whole workbook
    df_customer = pd.read_excel(EXCEL_FILE_PATH, sheet_name=customer_name)
    
    # Creating a temporary file path
    temp_dir = Path(os.environ['TEMP'])
    output_path = temp_dir / f"{customer_name}_Statement.xlsx"
    
    # Saving the dataframe to a new Excel file
    df_customer.to_excel(output_path, index=False)
    print(f"Saved: {output_path}")
    return output_path

def send_email_via_outlook(to_email, customer_name, attachment_path):
    """Creates and sends an email using Outlook."""
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    
    mail.To = to_email
    mail.CC = ";".join(CC_LIST)   
    mail.BCC = ";".join(BCC_LIST)
    mail.Subject = f"Your Monthly Statement - {pd.Timestamp.now().strftime('%B %Y')}"
    mail.Body = f"""Dear {customer_name},

Good day,
Please find your monthly statement attached.

Best regards,
Your Accounts Team"""
    
    mail.Attachments.Add(str(attachment_path))
    mail.Display()  # Change to .Send() to send automatically
    # mail.Send()

def main():
    print("Starting customer statement automation...")
    
    # Getting the list of customers and their emails
    customer_email_map = get_customer_email_map()
    print(f"Found {len(customer_email_map)} customers in the list.")
    
    # Lopping through each customer in the dictionary
    for customer_name, email_address in customer_email_map.items():
        try:
            print(f"Processing: {customer_name}...")
            
            # Checking if a sheet for this customer actually exists
            if customer_name in pd.ExcelFile(EXCEL_FILE_PATH).sheet_names:
                # Saving their sheet as a temporary file
                attachment_path = save_customer_sheet(customer_name)
                
                # Creating and sending the email
                send_email_via_outlook(email_address, customer_name, attachment_path)
                
                print(f"Email drafted for {customer_name}")
            else:
                print(f"Warning: No sheet found for customer '{customer_name}'. Skipping.")
                
        except Exception as e:
            # Error handling: log the error and continue with next customer
            print(f"*** ERROR processing {customer_name}: {e} ***")
    
    print("Automation finished! Please review Outlook drafts.")

if __name__ == "__main__":
    main()