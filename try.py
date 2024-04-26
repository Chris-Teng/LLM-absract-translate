import email
import mimetypes
import random
import datetime
import os

def generate_eml():
  """
  生成单段邮件的 eml 文件

  Returns:
    eml 文件内容
  """

  # 生成随机邮件内容
  sender_email = "[随机生成的电子邮件地址]"
  recipient_email = "[随机生成的电子邮件地址]"
  subject = "This is a test email"
  body = "This is the body of the email. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
  attachment_name = random.choice(["attachment1.txt", "attachment2.jpg", "attachment3.pdf"])
  attachment_content = "This is the content of the attachment." if attachment_name.endswith(".txt") else b"This is a binary attachment."

  # 生成 eml 文件内容
  eml_content = generate_eml_content(sender_email, recipient_email, subject, body, attachment_name, attachment_content)

  return eml_content

def generate_eml_content(sender_email, recipient_email, subject, body, attachment_name=None, attachment_content=None):
  """
  生成 eml 文件内容

  Args:
    sender_email: 发件人邮箱地址
    recipient_email: 收件人邮箱地址
    subject: 主题
    body: 正文内容
    attachment_name: 附件名称 (可选)
    attachment_content: 附件内容 (可选)

  Returns:
    eml 文件内容
  """

  # 生成邮件头信息
  msg = email.mime.Message()
  msg['From'] = sender_email
  msg['To'] = recipient_email
  msg['Subject'] = subject
  msg['Date'] = email.utils.formatdate(time.mktime(datetime.datetime.now().timetuple()))

  # 生成邮件正文
  msg.attach(email.mime.Text.MIMEText(body, 'plain'))

  # 添加附件 (可选)
  if attachment_name and attachment_content:
    maintype, subtype = mimetypes.guess_type(attachment_name)
    if maintype is None or subtype is None:
      raise ValueError('Unknown attachment type: %s' % attachment_name)

    attachment = email.mime.Base.MIMEBase(maintype, subtype)
    attachment.set_payload(attachment_content)
    email.encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=attachment_name)
    msg.attach(attachment)

  # 生成 eml 文件内容
  fp = io.BytesIO()
  msg.as_string(fp)
  return fp.getvalue()

# 生成 eml 文件
eml_content = generate_eml()

# 保存 eml 文件
with open('test.eml', 'wb') as f:
  f.write(eml_content)

print('Eml 文件已生成：test.eml')
