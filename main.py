import win32com.client
from datetime import datetime

if __name__ == '__main__':
    exclude_list = []

    with open('exclude.txt','r') as f:
        exclude_list = [line[:-1] for line in f.readlines()]

    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    folders = outlook.Folders

    email_list = []

    for folder in folders:
        for inner_folder in folder.Folders:
            print(f'Reading Folder {inner_folder.Name}')

            for message in inner_folder.Items:
                try:
                    recipients = message.Recipients
                    for recipient in recipients:
                        address = recipient.AddressEntry.Address.lower()
                        if not any(exclude in address for exclude in exclude_list) and \
                                address not in email_list and len(address) < 40 and \
                                '@gmail.com' in address:
                            email_list.append(address)
                            print(address)
                except Exception as e:
                    print(str(e))

    filename = f'extract_email_{datetime.timestamp(datetime.now())}.csv'
    with open(filename,'w', encoding='utf-8') as f:
        f.write('\n'.join(sorted(email_list)))
        print(f'emails saved to file - {filename}')