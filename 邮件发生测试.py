import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
from datetime import datetime
from dateutil.relativedelta import relativedelta
from email.mime.base import MIMEBase
from email import encoders
import os
# 邮件配置
smtp_server = 'smtp.qq.com'  # QQ邮件服务器地址
smtp_port = 587  # QQ邮件服务器端口
sender_email = '852910443@qq.com'  # 发件人邮箱，替换为你的QQ邮箱
sender_password = 'fnvvswkricrxbdca'  # 发件人邮箱授权码，不是登录密码，需要在QQ邮箱设置中生成
receiver_emails = '1965467608@qq.com'  # 收件人邮箱列表，替换为实际收件人邮箱

# 获取当前日期
today = datetime.today()
# 计算前两天的日期
two_days_ago = today - relativedelta(days=2)
date_str = two_days_ago.strftime("%Y年%m月%d日")

# 邮件正文
subject = f'{date_str}：二卡辅助证件执行情况'
body = ""
TWO_DAYS_AGO2='7月8日'
excel1_value='80%'
excel2_value='90%'
excel3_value='70%'
formatted_two_days_ago='7.08'
excel2_region=['湖里','集美','同安','营业外包']
excel1_value = '80%'
excel2_value = '90%'

# 将百分比转换为小数
excel1_decimal = float(excel1_value.strip('%')) / 100
excel3_decimal = float(excel2_value.strip('%')) / 100

# 计算差值（小数）
difference_decimal = excel1_decimal - excel3_decimal
# 如果你想要结果也是百分比形式，再乘以100并添加%
difference_percentage = abs(difference_decimal * 100)
difference_percentage_str = f'{difference_percentage:.2f}%'
print(difference_percentage_str)  # 输出：10.00%
if difference_decimal>0 :
    str3 = '提升'
else:
    str3 = '下降'

formatted_regions = [
    f"{region}营销中心" if region != "营业外包" else region
    for region in excel2_region
]
formatted_string = '、'.join(formatted_regions)
print(formatted_string)

# 构建邮件主题和正文
str1 = f'{TWO_DAYS_AGO2}，全区上传率为{excel1_value}，较前一日{str3}{difference_percentage_str}。'
str2 = f'6月1日-{formatted_two_days_ago[-2:]}日，全区上传率为{excel2_value}，其中{formatted_string}超80%，其余未达，请相关责任人做好各通路的宣贯'
subject = f'{date_str}辅助证件稽核结果'
# 邮件正文，包含颜色变化和换行
html_body = f"""
<html>
  <head>
    <style>
      .blue-text {{ color: blue; }}
      .red-text {{ color: red; }}
      .highlight {{ background-color: yellow; }}
    </style>
  </head>
  <body>
    <p>各位好!根据近期党委会的要求，
      <span class="blue-text">自2024年5月17日起</span>，
      针对年满16周岁及以上非厦门户籍且
      <span class="red-text">180天内办理第二张（不包含副卡）</span>及以上电话卡的，
      要求客户提供厦门居住证或医社保卡或学生证等辅助证件，相关材料需上传系统。
    </p>
    <p>
      <span class="red-text">管控要求：</span><br/>
      1、经营单位辅助证件上传率低于80%，将由网络安全领导小组办公室主任及市场部主要负责人约谈经营单位市场分管领导<br/>
      2、全市辅助证件上传率低于80%，将由网络安全领导小组副组长约谈市场部分管领导及不达标区县主要负责人
    </p>
    <p>
      <span class="highlight">{str1}</span><br/>
      <span class="highlight">{str2}</span>
    </p>
    <img src="cid:image1">
    <p>
      注：因支撑问题，通报数据为首次上传率，如期间有整改，数据暂时不会更新，各区可于每月8号前统一汇总收集并回复已整改数据，在整月通报及扣罚的时会更新。
    </p>
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
for file_path in ["F:\\IPA_project\\data\\6月辅助证件上传率日通报.xlsx", "F:\\IPA_project\\data\\新入网辅助证件统计表6.01-30.xlsx", "F:\\IPA_project\\data\\新入网辅助证件统计表6.30.xlsx"]:
    part = MIMEBase('application', "octet-stream")
    with open(file_path, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
    msg.attach(part)
# 图片文件名
image_filename = r"F:\IPA_project\data\邮件.png"  # 替换为实际图片文件的路径

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