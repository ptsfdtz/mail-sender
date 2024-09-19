import smtplib
import openpyxl
from email.message import EmailMessage
from tqdm import tqdm

class EmailSenderApp:
    def __init__(self):
        self.file_path = './data/录取名单.xlsx'
        self.username, self.password = self.read_credentials('./key.txt')

        recipient_data = self.read_recipient_list(self.file_path)

        if recipient_data:
            self.send_emails(self.username, self.password, recipient_data)
        else:
            print("读取收件人名单失败，程序退出。")

    def read_credentials(self, key_file):
        try:
            with open(key_file, 'r') as file:
                lines = file.readlines()
                username = lines[0].strip()
                password = lines[1].strip()
                return username, password
        except Exception as e:
            print(f"读取密钥文件时出错：{e}")
            return None, None

    def send_emails(self, email_address, email_password, recipient_data):
        total_recipients = len(recipient_data)
        
        with tqdm(total=total_recipients, desc="发送进度", unit="封") as pbar:
            for i, (recipient_name, recipient_email) in enumerate(recipient_data.items(), 1):
                self.send_email(email_address, email_password, recipient_name, recipient_email)
                pbar.update(1)  # 更新进度条

        print("所有邮件发送完毕。")

    def send_email(self, email_address, email_password, recipient_name, recipient_email):
        smtp_host, smtp_port = self.get_smtp_settings(email_address)
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

    def get_smtp_settings(self, email_address):
        domain = email_address.split('@')[1].lower()

        if domain == '163.com':
            return 'smtp.163.com', 465
        elif domain == 'gmail.com':
            return 'smtp.gmail.com', 465
        elif domain == 'qq.com':
            return 'smtp.qq.com', 465
        else:
            raise ValueError(f"不支持的电子邮件域名: {domain}")

    def read_recipient_list(self, file_path):
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            name_row_index = 0
            email_row_index = 1

            result_dict = {}
            for row in sheet.iter_rows(min_row=2, values_only=True):
                name = row[name_row_index]
                email = row[email_row_index]
                result_dict[name] = email

            return result_dict
        except FileNotFoundError:
            print(f"文件未找到：{file_path}")
        except ValueError:
            print(f"表格中找不到 '姓名' 或 '邮件' 列")
        except Exception as e:
            print(f"读取文件时出错：{e}")
            return None


def main():
    EmailSenderApp()


if __name__ == '__main__':
    main()
