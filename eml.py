import email,re
from bs4 import BeautifulSoup

def read_eml(filename):
  """
  读取 eml 文件并获取邮件内容

  Args:
    filename: eml 文件名

  Returns:
    邮件内容
  """
  with open(filename, 'rb') as f:
    # 解析 eml 文件
    msg = email.message_from_bytes(f.read())

    # 获取邮件正文
    if msg.is_multipart():
      # 处理多段邮件
        body = ''
        for part in msg.walk():
            if part.get_content_maintype() == 'text':
            # 去除 HTML 标签
                soup = BeautifulSoup(part.get_payload(decode=True), 'html.parser')
                text = soup.get_text(separator='\n')
                body += text
                body = body.replace('\n', ' ')
    else:
      # 处理单段邮件
        body = re.sub(r'<[^>]*>', '', msg.get_payload(decode=True).decode())
    # 返回邮件内容
    return body

if __name__ == '__main__':
  # 设置 eml 文件名
  filename = 'example.eml'

  # 读取 eml 文件并打印内容
  content = read_eml(filename)
  print(content)
