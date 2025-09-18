import smtplib
import json
from email.message import EmailMessage
from tqdm import tqdm
from utils.file_utils import read_recipient_list, read_credentials
from utils.smtp_utils import get_smtp_settings


class EmailSenderApp:
    def __init__(self, status):
        self.username, self.password = read_credentials('./key.txt')

        if status == "面试":
            recipient_file = './data/面试名单.xlsx'
            recipients = read_recipient_list(recipient_file)
            if recipients:
                self.send_interview_emails(recipients)

        elif status == "录取":
            recipient_file = './data/录取名单.xlsx'
            recipients = read_recipient_list(recipient_file)
            if recipients:
                self.send_admission_emails(recipients)

        else:
            print("❌ 输入错误，只能是 '面试' 或 '录取'")

    def send_interview_emails(self, recipient_data):
        total_recipients = len(recipient_data)
        with tqdm(total=total_recipients, desc=f"sending 面试 emails", unit="emails") as pbar:
            for recipient in recipient_data:
                name = recipient["name"]
                email = recipient["email"]
                self.send_email(name, email, "面试")
                pbar.update(1)

    def send_admission_emails(self, recipient_data):
        total_recipients = len(recipient_data)
        with tqdm(total=total_recipients, desc=f"sending 录取/未录取 emails", unit="emails") as pbar:
            for recipient in recipient_data:
                name = recipient["name"]
                email = recipient["email"]
                department = recipient["department"]
                qq_group_id = recipient["qqGroupId"]

                if department and department != "无":
                    status = "录取"
                else:
                    status = "未录取"

                self.send_email(name, email, status, department, qq_group_id)
                pbar.update(1)

    def send_email(self, recipient_name, recipient_email, status, department=None, qq_group_id=None):
        smtp_host, smtp_port = get_smtp_settings(self.username)
        smtp = smtplib.SMTP_SSL(smtp_host, smtp_port)

        subject = f"南京工业大学大学生科学技术协会{status}通知"
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = self.username
        msg['To'] = recipient_email

        try:
            template_path = f'web/{status}通知.html'
            with open(template_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # 公共替换
            html_content = html_content.replace('{{name}}', str(recipient_name))

            # 录取才替换群号和部门
            if status == "录取":
                html_content = html_content.replace('{{qqGroupId}}', str(qq_group_id))
                html_content = html_content.replace('{{department}}', str(department))

            msg.add_alternative(html_content, subtype='html')

            smtp.login(self.username, self.password)
            smtp.send_message(msg)
            print(f"✅ successfully sent {status} email to {recipient_name} ({recipient_email})")
        except Exception as e:
            print(f"❌ failed to send {status} email to {recipient_name} ({recipient_email}): {e}")
        finally:
            smtp.quit()


def main():
    status = input("请输入要发送的邮件类型（面试 / 录取）: ").strip()
    EmailSenderApp(status)
    input("按Enter键退出...")


if __name__ == '__main__':
    main()
