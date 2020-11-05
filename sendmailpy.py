from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL, SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os


class SmtpTransport:
    def __init__(self, hostname, port=None, ssl=True, starttls=False):
        self.hostname = hostname

        if ssl:
            self.port = port or 465
            self.server = SMTP_SSL(self.hostname, self.port)
        else:
            self.port = port or 25
            self.server = SMTP(self.hostname, self.port)

        if starttls:
            self.server.starttls()

    def connect(self, username, password):
        self.server.login(username, password)
        return self.server


class OpenMail(object):
    def __init__(self, hostname, username=None, password=None, ssl=True,
                 port=None, policy=None, starttls=False):

        self.server = SmtpTransport(hostname, ssl=ssl, port=port, starttls=starttls)
        self.hostname = hostname
        self.username = username
        self.password = password
        self.parser_policy = policy
        self.connection = self.server.connect(username, password)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.logout()

    def logout(self):
        self.connection.quit()

    def create_header(self, message, to_mail, subject):
        message['From'] = Header(self.username, 'utf-8')
        if isinstance(to_mail, list):
            for one_mail in to_mail:
                message['To'] = Header(one_mail, 'utf-8')
        else:
            message['To'] = Header(to_mail, 'utf-8')
        message['Subject'] = Header(subject, 'utf-8')

    def attach_file(self,message,file):
        # 添加附件
        excelFile = open(file, 'rb').read()
        fileName = os.path.basename(file)
        att = MIMEApplication(excelFile)
        att.add_header('Content-Disposition', 'attachment',
                       filename=Header(fileName, 'utf-8').encode())
        message.attach(att)
        return message

    def create_message(self, content_type, to_mail, subject, content, file):
        if content_type == "plain":
            message = MIMEText(content, 'plain', 'utf-8')
            self.create_header(message, to_mail, subject)
            return message
        elif content_type == "html":
            message = MIMEText(content, 'html', 'utf-8')
            self.create_header(message, to_mail, subject)
            return message
        elif content_type == "attach":
            message = MIMEMultipart()
            self.create_header(message, to_mail, subject)
            # 邮件正文内容
            message.attach(MIMEText(content, 'plain', 'utf-8'))
            if isinstance(file, list):
                for f in file:
                    # 添加附件
                    message = self.attach_file(message, f)
            else:
                # 添加附件
                message = self.attach_file(message, file)
            return message

    def newMail(self, to_mail, subject, content, file, content_type):
        return self.create_message(content_type, to_mail, subject, content, file)

    def send(self, msg, to_mail):
        # 发送邮件：发送方，收件方，要发送的消息
        self.connection.sendmail(self.username, to_mail, msg.as_string())

    def sendmail(self, to_mail, subject="", content="", file=None, content_type="plain", send_type="same"):
        '''
        :param to_mail: 发送邮件地址，可以是list，多个邮件
        :param subject: 邮件标题
        :param content: 邮件内容
        :param content_type: 邮件内容格式：默认plain普通模式，html格式，attach附件格式
        :param send_type: 发送模式：same一对一或一对多发送相同的附件，different一对一发送不同的附件
        :param file:
        :return:
        '''
        if file and isinstance(file, str):
            if os.path.exists(file):
                content_type = "attach"
            else:
                raise ("%s:文件不存在" % file)
        if file and isinstance(file, list):
            for f in file:
                if isinstance(f,list):
                    for _f in f:
                        if os.path.exists(_f):
                            content_type = "attach"
                        else:
                            raise ("%s:文件不存在" % _f)
                else:
                    if os.path.exists(f):
                        content_type = "attach"
                    else:
                        raise ("%s:文件不存在" % f)

        if send_type=="same":
            msg = self.newMail(to_mail, subject, content, file, content_type)
            self.send(msg, to_mail)
        elif send_type=="different":
            if len(to_mail)!=len(file):
                raise ("发送地址数量与附件数量必须一致")
            for i in range(len(to_mail)):
                msg=self.newMail(to_mail[i], subject, content, file[i], content_type)
                self.send(msg, to_mail[i])


if __name__ == '__main__':
    emailaddress = "ffm1110@qq.com"
    password = "gglgeyeycrfcbjjj"
    hostname = "imap.qq.com"
    om = OpenMail(hostname, emailaddress, password, port=465)
    om.sendmail(["ffm1110@qq.com", "ffm1110@163.com"], "测试", "内容测试",
                file=[[r"C:\Users\rpadev\Desktop\驾照统计.xlsx"],
                      [r"C:\Users\rpadev\Desktop\批发报表们的自动化统计.docx",
                      r"C:\Users\rpadev\Desktop\附件2-7人民银行征信系统标准  企业信用报告产品说明（二代试行）.pdf"]],send_type="different"
                )
    om.logout()
