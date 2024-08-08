import imaplib
import email
from email.header import decode_header
import ssl

# 邮箱配置信息
EMAIL = "13859915120@139.com"
PASSWORD = "50b900b59e21d5319000"
IMAP_SERVER = "imap.139.com"

# 连接到IMAP服务器
ctx = ssl.create_default_context()
ctx.set_ciphers('DEFAULT')

with imaplib.IMAP4_SSL(IMAP_SERVER, 993, ssl_context=ctx) as mail:
    mail.login(EMAIL, PASSWORD)

    # 选择邮箱文件夹，默认为"INBOX"
    mail.select("INBOX")

    # 搜索具有特定主题的邮件
    search_criteria = '(SUBJECT "周体检报告")'.encode('utf-8')
    status, messages = mail.search(None, search_criteria)

    # 解析搜索结果
    messages = messages[0].split(b' ')

    # 从最新邮件开始查找
    for msg_id in reversed(messages):
        if msg_id != b'':
            status, data = mail.fetch(msg_id, "(RFC822)")
            if data and isinstance(data[0], tuple):
                raw_email = data[0][1]
                if isinstance(raw_email, bytes):
                    email_message = email.message_from_bytes(raw_email)

                    # 查找符合要求的附件
                    for part in email_message.walk():
                        # 检查是否有附件
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        filename = part.get_filename()
                        if not filename:
                            continue

                        # 解码文件名
                        filename = decode_header(filename)[0][0]
                        if isinstance(filename, bytes):
                            # 如果是字节，则解码为字符串
                            filename = filename.decode()

                        # 检查是否包含特定文本的附件
                        if "网格调度测试" in filename:
                            with open(filename, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                                print(f"已下载附件: {filename}")
                            break  # 找到第一个符合要求的附件后停止查找
                    else:
                        continue  # 没有找到符合要求的附件，继续查找下一邮件
                    break  # 找到了符合要求的附件，停止遍历邮件

    else:
        print("No attachments meeting the criteria were found.")