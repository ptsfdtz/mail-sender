def get_smtp_settings(email_address=None):
    """
    自建邮件服务器配置
    """
    smtp_host = 'mail.wzj.su'  
    smtp_port = 465
    return smtp_host, smtp_port
