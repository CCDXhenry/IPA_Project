import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
from datetime import datetime
from dateutil.relativedelta import relativedelta
# 邮件配置
smtp_server = 'smtp.139.com'  # 邮件服务器地址
smtp_port = 25  # 邮件服务器端口
sender_email = '13850099699@139.com'  # 发件人邮箱
sender_password = '187d1f1e79fc3754d100'  # 发件人邮箱密码或授权码（根据邮箱提供商的要求）
receiver_emails = '13850099699@139.com,13666042109@139.com'  # 收件人邮箱列表

# 获取当前日期
today = datetime.today()
# 计算前两天的日期
two_days_ago = today - relativedelta(days=2)
date_str = two_days_ago.strftime("%Y年%m月%d日")

# 邮件正文
subject = f'{date_str}：二卡辅助证件执行情况'
body = ""

# 构建邮件主题和正文
subject = f'{date_str}辅助证件稽核结果'
html_body = """
<html>
  <head></head>
  <body>
    <p>各位好!根据近期党委会的要求，自2024年5月17日起，针对年满16周岁及以上非厦门户籍且180天内办理第二张（不包含副卡）及以上电话卡的，要求客户提供厦门居住证或医社保卡或学生证等辅助证件，相关材料需上传系统。
管控要求：
1、经营单位辅助证件上传率低于80%，将由网络安全领导小组办公室主任及市场部主要负责人约谈经营单位市场分管领导
2、全市辅助证件上传率低于80%，将由网络安全领导小组副组长约谈市场部分管领导及不达标区县主要负责人

X月X日，全区上传率为 X%，较前一日下降/提升X%。
1日-X日，全区上传率为X%，其中X、X超80%，其余未达，请相关责任人做好各通路的宣贯
</p>
    <img src="cid:image1">
    <p>贴图
注：因支撑问题，通报数据为首次上传率，如期间有整改，数据暂时不会更新，各区可于每月8号前统一汇总收集并回复已整改数据，在整月通报及扣罚的时会更新。</p>
  </body>
</html>
"""

# 创建MIMEMultipart对象
msg = MIMEMultipart('related')

# 设置邮件头部信息
msg['Subject'] = Header(subject, 'utf-8')
msg['From'] = sender_email
msg['To'] = receiver_emails

# 创建HTML文本部分
html_part = MIMEText(html_body, 'html')
msg.attach(html_part)

# 图片文件名
image_filename = r"/data/邮件.png"  # 替换为实际图片文件的路径

# 读取图片文件并创建MIMEImage对象
with open(image_filename, 'rb') as fp:
    img = MIMEImage(fp.read())
    img.add_header('Content-ID', '<image1>')  # 设置Content-ID，与HTML中的cid相对应
    msg.attach(img)

# 发送邮件
try:
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # 开启TLS加密
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_emails.split(','), msg.as_string())
        print('邮件发送成功')
except smtplib.SMTPException as e:
    print(f'邮件发送失败：{str(e)}')
except Exception as e:
    print(f'发生未知错误：{str(e)}')