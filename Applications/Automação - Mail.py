import os
import imaplib
import email
from collections import defaultdict
import time
from datetime import datetime, timedelta

def download_last_attachment(username, password, imap_server, folder, target_folder, subjects):
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(username, password)
        mail.select(folder)
        email_dict = defaultdict(list)
        current_datetime = datetime.now()
        five_hours_ago = current_datetime - timedelta(hours=5)
        current_date = current_datetime.strftime('%d-%b-%Y')
        five_hours_ago_date = five_hours_ago.strftime('%d-%b-%Y')

        for subject in subjects:
            search_criteria = f'(SUBJECT "{subject}" SINCE "{five_hours_ago_date}")'
            status, messages = mail.search(None, search_criteria)
            messages = messages[0].split()
            for msg_id in messages:
                email_dict[subject].append(msg_id)
            time.sleep(2)

        for subject, msg_ids in email_dict.items():
            latest_msg_id = max(msg_ids)
            status, msg_data = mail.fetch(latest_msg_id, '(RFC822)')
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
s
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                filename = part.get_filename()
                if filename:
                    filename = os.path.basename(filename)
                    file_path = os.path.join(target_folder, filename)
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))

            mail.store(latest_msg_id, '+FLAGS', '(\Seen)')
            mail.store(latest_msg_id, '+FLAGS', '(\Deleted)')
            time.sleep(2)

        mail.expunge()
        mail.logout()

        print("Anexos baixados com sucesso!")
    except Exception as e:
        print("Ocorreu um erro:", e)

username = "@hotmail.com"
password = ""
imap_server = "outlook.office365.com"
folder = "INBOX" 
target_folder = ""
subjects = [
    "f_pedido_normal",
    "f_pedido_compra_especial",
    "f_billing_info_cnpj",
    "f_billing_info_cpf"
]

download_last_attachment(username, password, imap_server, folder, target_folder, subjects)