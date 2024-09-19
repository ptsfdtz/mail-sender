import smtplib
from email.message import EmailMessage
from tqdm import tqdm
from utils.file_utils import read_recipient_list, read_credentials
from utils.smtp_utils import get_smtp_settings

class EmailSenderApp:
    def __init__(self):
        self.file_path = './data/录取名单.xlsx'
        self.username, self.password = read_credentials('./key.txt')

        recipient_data = read_recipient_list(self.file_path)

        if recipient_data:
            self.send_emails(self.username, self.password, recipient_data)
        else:
            print("读取收件人名单失败，程序退出。")

    def send_emails(self, email_address, email_password, recipient_data):
        total_recipients = len(recipient_data)
        
        with tqdm(total=total_recipients, desc="发送进度", unit="封") as pbar:
            for i, (recipient_name, recipient_email) in enumerate(recipient_data.items(), 1):
                self.send_email(email_address, email_password, recipient_name, recipient_email)
                pbar.update(1)

        print("所有邮件发送完毕。")

    def send_email(self, email_address, email_password, recipient_name, recipient_email):
        smtp_host, smtp_port = get_smtp_settings(email_address)
        smtp = smtplib.SMTP_SSL(smtp_host, smtp_port)

        subject = "南京工业大学校大学生科学技术协会面试结果通知"
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = email_address
        msg['To'] = recipient_email

        try:
            with open('web/录取通知.html', 'r', encoding='utf-8') as f:
                html_content = f.read()

            html_content = html_content.replace('{{name}}', recipient_name)
            html_content = html_content.replace('{{qqGroupId}}', '788418335')
            msg.add_alternative(html_content, subtype='html')
            smtp.login(email_address, email_password)
            smtp.send_message(msg)
            print(f"成功发送给 {recipient_name} ({recipient_email})")
        except Exception as e:
            print(f"发送邮件失败: {e}")
        finally:
            smtp.quit()

def main():
    EmailSenderApp()

if __name__ == '__main__':
    main()
