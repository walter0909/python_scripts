from selenium import webdriver
import time
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import  MIMEImage
from email.utils import parseaddr, formataddr




def monitor_picture(monitor_url, url_username, url_passwd, img_dir):
    option = webdriver.ChromeOptions()
    # 全屏展示
    option.add_argument('--start-maximized')
    # chrom 提示 “正在受到自动测试软件的控制” 关闭项，新版本配置变更。后面在查资料 option.add_argument('--disable-infobars')
    # option.add_argument('disable-infobars')
    # option.add_experimental_option('excludeSwitches',['enable-automation'])
    # 无痕模式
    # option.add_argument('--incognito')
    # 无界面浏览
    # option.add_argument('--headless')
    # options.binary_location = r"D:\software\chrome\chromedriver.exe"
    # 1.创建Chrome浏览器对象，这会在电脑上在打开一个浏览器窗口
    browser = webdriver.Chrome(chrome_options=option)
    login_url = 'https://xxxx/login'
    browser.get(login_url)
    # browser.get_screenshot_as_file("D:\software\chrome\hello1.png")

    browser.find_elements_by_class_name("css-1bjepp-input-input")[0].send_keys(url_username)
    browser.find_elements_by_class_name("css-1bjepp-input-input")[1].send_keys(url_passwd)
    browser.find_element_by_class_name("css-6ntnx5-button").click()

    for i in monitor_url:
        print(i)
        time.sleep(3)
        #browser = webdriver.Chrome(chrome_options=option)
        browser.get(i)
        time.sleep(3)
        print('open url')
        #按比例缩小 https://www.cnblogs.com/wdana/p/12037567.html
        browser.execute_script("document.body.style.zoom='0.5'")

        time.sleep(3)
        picture_time = time.strftime("%Y%m%d%H%M%S",time.localtime(time.time()))
        browser.get_screenshot_as_file(""+img_dir+""+picture_time+".png")






def sendEmail(from_addr, password, smtp_server, smtp_server_ssl_port, img_dir, to_addr):
    sendmail_time = time.strftime("%Y/%m/%d/ %H:%M:%S", time.localtime(time.time()))
    subject = sendmail_time + ' 巡检信息'
    html_content = \
    '''
        <html>
        <body>
            <p>硬盘<img src="cid:1"></p>
            <p>内存<img src="cid:2"></p>
        </body>
        </html>
    '''

    msg = MIMEMultipart()
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    i = 1
    for img_file in os.listdir(img_dir):
        with open(os.path.join(img_dir + img_file), 'rb') as im:
            mime = MIMEBase('image', 'png', filename=img_file)
            mime.add_header('Content-Disposition', 'attachment', filename=img_file)
            mime.add_header('Content-ID', '<%i>' % i)
            mime.set_payload(im.read())
            encoders.encode_base64(mime)
            msg.attach(mime)
            i = i + 1

    msg['From'] = from_addr
    msg['to'] = to_addr
    # msg['Subject'] = Header('邮件主体', 'utf-8').encode()
    msg['Subject'] = subject

    server = smtplib.SMTP_SSL(smtp_server, smtp_server_ssl_port)
    server.login(from_addr, password)
    server.sendmail(from_addr, to_addr, msg.as_string())
    server.quit()




if __name__ == '__main__':
    img_dir = 'D:\\software\\chrome\\picture\\'
    from_addr = 'xxx.com'
    password = 'xxx'
    smtp_server = 'xxx'
    smtp_server_ssl_port = xxx
    to_addr = 'xxx'

    url_username = 'xxx'
    url_passwd = 'xxx'
    monitor_url = ['https://xxx',
                   'https://xxx']

    monitor_picture(monitor_url, url_username, url_passwd, img_dir)
    sendEmail(from_addr, password, smtp_server, smtp_server_ssl_port, img_dir, to_addr)






