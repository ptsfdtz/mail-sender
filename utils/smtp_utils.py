def get_smtp_settings(email_address):
    domain = email_address.split('@')[1].lower()

    if domain == '163.com':
        return 'smtp.163.com', 465
    elif domain == 'gmail.com':
        return 'smtp.gmail.com', 465
    elif domain == 'qq.com':
        return 'smtp.qq.com', 465
    else:
        raise ValueError(f"不支持的电子邮件域名: {domain}")
