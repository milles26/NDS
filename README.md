import openpyxl
import os
import datetime
import pandas as pd
import win32com.client


def process_excel_files(main_file_path, output_folder):
    try:
        # Read the main Excel workbook
        df = pd.read_excel(main_file_path, sheet_name="0000")

        # Read the distribution list file
        dist_list = pd.read_excel("Distribution List .xlsx", sheet_name="SA Cover File ")

        # Get unique customer names from column I
        customers = df['Vendor Name'].unique()
        SA = df['Analyst'].unique()

        # Iterate through each customer
        for customer in customers:
            # Create a new folder with today's date as part of the filename inside the output_folder
            folder_name = f"{customer}"
            folder_path = os.path.join(output_folder, folder_name)

            # Check if the folder already exists
            if not os.path.exists(folder_path):
                # Create the directory if it doesn't exist
                os.makedirs(folder_path)
                print(f"Created folder: {folder_name}")

            ship_date = df.loc[df['Vendor Name'] == customer, 'Ship Date'].iloc[0]
            # date_obj = datetime.datetime.strptime(ship_date, "%d/%m/%Y")
            new_date_str = ship_date.strftime("%d-%m-%Y")
            # print(new_date_str)

            # Create a new workbook using xlsxwriter engine
            filename = f"{customer} - {new_date_str}.xlsx"  # {datetime.now().strftime('%d-%m-%Y')}.xlsx"
            file_path = os.path.join(folder_path, filename)

            # Create a new workbook and select the first sheet
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

            # Write the filtered dataframe to the Excel file
            df_filtered = df[df['Vendor Name'] == customer]
            columns_to_keep = ['Vendor', 'Vendor Name', 'Item', 'Description', 'Ship Date', 'CO', 'Qty']
            df_filtered = df_filtered[columns_to_keep]

            # Convert date format in column E to dd-mm-yyyy without timestamp
            df_filtered['Ship Date'] = pd.to_datetime(df_filtered['Ship Date'],
                                                      format='mixed',
                                                      errors='coerce').dt.strftime('%d-%m-%Y')

            # Write the dataframe to the Excel file
            df_filtered.to_excel(writer, sheet_name=customer[:10], index=False)

            # Get the worksheet object
            worksheet = writer.sheets[customer[:10]]

            # Adjust column widths
            for idx, col in enumerate(df_filtered.columns):
                series = df_filtered[col]
                max_len = max((
                    series.astype(str).map(len).max(),
                    len(col)
                )) + 1
                worksheet.set_column(idx, idx, max_len)

            # Close the writer to save the file
            writer.close()

            # Get the corresponding row from the distribution list
            dist_info = dist_list[dist_list['Vendor Name'] == customer]

            if not dist_info.empty:
                try:
                    # Get email address from column D
                    mailing_address = dist_info.iloc[0, 3]  # Column index 3 corresponds to column D
                    cc_address = dist_info.iloc[0, 5] + ';' + dist_info.iloc[
                        0, 7]  # Column index 5 corresponds to column E
                except IndexError:
                    print(f"No email address found for {customer}")
                    mailing_address = ""
                    cc_address = ""
            else:
                print(f"No distribution list information found for {customer}")
                mailing_address = ""
                cc_address = ""

            # Create Outlook email
            outlook = win32com.client.Dispatch('Outlook.Application')

            for account in outlook.Session.Accounts:
                if account.DisplayName == 'UKLCautomail@cat.com':  # Replace with actual account name
                    mail = outlook.CreateItem(0)  # olMailItem
            mail.To = mailing_address
            mail.CC = cc_address
            mail.Subject = f"You have a schedule due for {customer} - {df_filtered.iloc[0]['Ship Date']} - {dist_info['Analyst'].values[0]}"
            # Set the sender account
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
            # mail = outlook.CreateItem(0)  # olMailItem
            #
            # # Set email properties
            # mail.To = mailing_address  # Use mailing_address instead of customer + "@example.com"
            # mail.CC = cc_address
            # mail.Subject = f"You have a schedule due {customer} - {df_filtered.iloc[0]['Ship Date']}"  # [:21]}
            mail.Body = f"""Dear Supplier,



A shipment due to be dispatched on {df_filtered.iloc[0]['Ship Date']}  for Caterpillar Logistics (UK) Ltd OY into Desford.

This Shipment contains the attached Part/s.

Please confirm you are on track with this delivery being on-time and in full in the next 24 hours.

Please remember all schedule collaborations must take place on MRC and ASNs submitted on time to positively affect your SSP.
Please also make sure that  you update E-order on every shipment going out to CAT if you are using E order tool. 
This is mandatory – if you have any issues please contact transportation team or your current SA’s . 


This inbox is not monitored. Please respond to your Supply Analyst on Copy and please review your Part Numbers on MRC.


Thank you.

Supply Analyst Team
Product Support & Logistics Division (PSLD)
Caterpillar Logistics (UK) Ltd.
Peckleton Lane, Desford, Leics. LE9 9JU
"""

            # Attach the Excel file
            attachment_path = os.path.join(folder_path, filename)
            mail.Attachments.Add(attachment_path)

            # Send the email
            try:
                mail.Send()
                #mail.Display()
                #print(f"Email sent to {customer}")
            except Exception as e:
                print(f"Failed to send email to {customer}: {str(e)}")

    except Exception as e:
        print(f"An error occurred: {e}")


# Usage
main_file_path = r"Next Due Shipment Daily Check List 2025.xlsx"
output_folder = r"C:\Users\milles26\OneDrive - Caterpillar\Next Due Shipments 2025\Python AutoEmail Supplier Folder\Suppliers"
process_excel_files(main_file_path, output_folder)
