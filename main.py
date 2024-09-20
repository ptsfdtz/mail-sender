import smtplib
import json
from email.message import EmailMessage
from tqdm import tqdm
from utils.file_utils import read_recipient_list, read_credentials
from utils.smtp_utils import get_smtp_settings

class EmailSenderApp:
    def __init__(self, config):
        self.status = config["status"]
        self.department = config["department"]
        self.qq_group_id = config["qqGroupId"]
        self.file_path = f'./data/{self.status}名单.xlsx'
        self.username, self.password = read_credentials('./key.txt')
        recipient_data = read_recipient_list(self.file_path)

        if recipient_data:
            self.send_emails(self.username, self.password, recipient_data)

    def send_emails(self, email_address, email_password, recipient_data):
        total_recipients = len(recipient_data)
        
        with tqdm(total=total_recipients, desc="sending emails", unit="emails") as pbar:
            for i, (recipient_name, recipient_email) in enumerate(recipient_data.items(), 1):
                self.send_email(email_address, email_password, recipient_name, recipient_email)
                pbar.update(1)

    def send_email(self, email_address, email_password, recipient_name, recipient_email):
        smtp_host, smtp_port = get_smtp_settings(email_address)
        smtp = smtplib.SMTP_SSL(smtp_host, smtp_port)

        subject = "南京工业大学校大学生科学技术协会面试结果通知"
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = email_address
        msg['To'] = recipient_email

        try:
            __path__ = f'web/{self.status}通知.html'
            with open(__path__, 'r', encoding='utf-8') as f:
                html_content = f.read()

            html_content = html_content.replace('{{name}}', recipient_name)
            html_content = html_content.replace('{{qqGroupId}}', self.qq_group_id)
            html_content = html_content.replace('{{department}}', self.department)
            msg.add_alternative(html_content, subtype='html')
            smtp.login(email_address, email_password)
            smtp.send_message(msg)
            print(f"successfully sent email to {recipient_name} ({recipient_email})")
        except Exception as e:
            print(f"failed to send email to {recipient_name} ({recipient_email}): {e}")
        finally:
            smtp.quit()

def main():
    with open('config.json', 'r', encoding='utf-8') as config_file:
        config = json.load(config_file)

    print(json.dumps(config, indent=4, ensure_ascii=False))

    question = input("确认发送邮件吗？(y/n): ")
    if question.lower() == 'y':
        EmailSenderApp(config)
        input("按任意键退出...")

if __name__ == '__main__':
    main()