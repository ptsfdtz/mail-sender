import openpyxl

def read_credentials(key_file):
    try:
        with open(key_file, 'r') as file:
            lines = file.readlines()
            username = lines[0].strip()
            password = lines[1].strip()
            return username, password
    except Exception as e:
        print(f"读取密钥文件时出错：{e}")
        return None, None

def read_recipient_list(file_path):
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
