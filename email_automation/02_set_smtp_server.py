import os
import smtplib
from email.mime.text import MIMEText

from openpyxl.reader.excel import load_workbook


class EmailSender:
    email_addr = None
    password = None
    smtp_server_map = {
        'gmail.com':'smtp.gmail.com',
        'naver.com':'smtp.naver.com'
    }
    smtp_server = None

    def __init__(self, email_addr, password):
        print('생성자')
        self.email_addr = email_addr
        self.password = password
        self.smtp_server = self.smtp_server_map[email_addr.split('@')[1]]

    def send_email(self, msg, from_addr, to_addr, subject):
        '''
        :param msg: 보낼 메세지
        :param from_addr: 보내는 사람
        :param to_addr: 받는 사람
        :return:
        '''
        with smtplib.SMTP(self.smtp_server, 587) as smtp:
            msg = MIMEText(msg)
            msg['From'] = from_addr
            msg['To'] = to_addr
            msg['Subject'] = subject
            print(msg.as_string())

            smtp.starttls()
            smtp.login(self.email_addr, self.password)
            smtp.sendmail(from_addr=from_addr, to_addrs=to_addr, msg=msg.encode('utf-8'))
            smtp.quit()
        print('이메일 전송이 완료 되었습니다.')

    def send_all_emails(self, filename):
        wb = load_workbook(filename)
        ws = wb.active




if __name__ == '__main__':
    es = EmailSender('fc.krkim@gmail.com', os.getenv('MY_GMAIL_PASSWORD'))
    # es = EmailSender('fc.krkim@naver.com', os.getenv('MY_NAVER_PASSWORD'))
    # es.send_email('테스트 입니다.',
    #               from_addr='fc.krkim@naver.com',
    #               to_addr='fc.krkim@gmail.com', subject='이메일 전송 테스트')
    es.send_all_emails('이메일리스트.xlsx')