import smtplib
import json
from email.message import EmailMessage
from tqdm import tqdm
from utils.file_utils import read_recipient_list, read_credentials
from utils.smtp_utils import get_smtp_settings

class EmailSenderApp:
    def __init__(self, config):
        self.department = config["department"]
        self.qq_group_id = config["qqGroupId"]
        self.username, self.password = read_credentials('./key.txt')

        # 录取名单
        admitted_file = './data/录取名单.xlsx'
        admitted_recipients = read_recipient_list(admitted_file)
        if admitted_recipients:
            self.send_emails(admitted_recipients, status="录取")

        # 未录取名单
        rejected_file = './data/未录取名单.xlsx'
        rejected_recipients = read_recipient_list(rejected_file)
        if rejected_recipients:
            self.send_emails(rejected_recipients, status="未录取")

    def send_emails(self, recipient_data, status):
        total_recipients = len(recipient_data)
        with tqdm(total=total_recipients, desc=f"sending {status} emails", unit="emails") as pbar:
            for recipient_name, recipient_email in recipient_data.items():
                self.send_email(recipient_name, recipient_email, status)
                pbar.update(1)

    def send_email(self, recipient_name, recipient_email, status):
        smtp_host, smtp_port = get_smtp_settings(self.username)
        smtp = smtplib.SMTP_SSL(smtp_host, smtp_port)

        subject = "南京工业大学大学生科学技术协会面试结果通知"
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = self.username
        msg['To'] = recipient_email

        try:
            template_path = f'web/{status}通知.html'
            with open(template_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # 公共替换
            html_content = html_content.replace('{{name}}', recipient_name)

            # 录取通知才替换群号和部门
            if status == "录取":
                html_content = html_content.replace('{{qqGroupId}}', self.qq_group_id)
                html_content = html_content.replace('{{department}}', self.department)

            msg.add_alternative(html_content, subtype='html')

            smtp.login(self.username, self.password)
            smtp.send_message(msg)
            print(f"successfully sent {status} email to {recipient_name} ({recipient_email})")
        except Exception as e:
            print(f"failed to send {status} email to {recipient_name} ({recipient_email}): {e}")
        finally:
            smtp.quit()

def main():
    with open('config.json', 'r', encoding='utf-8') as config_file:
        config = json.load(config_file)

    print(json.dumps(config, indent=4, ensure_ascii=False))

    question = input("确认发送录取和未录取邮件吗？(y/n): ")
    if question.lower() == 'y':
        EmailSenderApp(config)
        input("按任意键退出...")

if __name__ == '__main__':
    main()
