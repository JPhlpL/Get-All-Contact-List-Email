import win32com.client
import csv

# Connect to Outlook application
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get Address Book folder
address_book = outlook.GetDefaultFolder(10)

# Get all contacts from Address Book
contacts = address_book.Items

# Filter contacts by email address domain
filtered_contacts = [contact for contact in contacts if contact.Email1Address.endswith("@gmail.com")]

# Create CSV file and write headers
with open("gmail_contacts.csv", mode="w", newline="") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["Name", "Department", "Position", "Email"])

    # Write filtered contacts to CSV file
    for contact in filtered_contacts:
        writer.writerow([contact.FullName, contact.Department, contact.JobTitle, contact.Email1Address])