import yagmail
import pandas as pd
from time import sleep
import os
import random
import configparser


class EmailHandle:
    def __init__(self, user, passwd, host):
        self.yag = yagmail.SMTP(user=user, password=passwd, host=host)
        if os.path.isfile('email.xlsx') is False:
            col_name = ['邮箱', '名字', '主题', '内容', '文件名']
            df = pd.DataFrame(columns=col_name)
            df.to_excel('email.xlsx', index=False)

    def send_mail(self, to_mail, subject='', content='', file=None):
        try:
            if file is not None:
                self.yag.send(to_mail, subject, content, file)
            else:
                self.yag.send(to_mail, subject, content)

            return True
        except Exception as e:
            print(e)
            return False

    def batch_send(self, excel_file):
        sleep_t = random.randint(1, 3)
        df = pd.read_excel(excel_file)
        to_mail_list = list(df['邮箱'].values)
        name_list = list(df['名字'].values)
        file_list = list(df['文件名'].values)
        subject_list = list(df['主题'].values)
        content_list = list(df['内容'].values)
        for to_mail, subject, content, file, name in zip(to_mail_list, subject_list, content_list, file_list,
                                                         name_list):
            if os.path.isfile(file):
                if self.send_mail(to_mail, subject=subject, content=content, file=file):
                    print(f'{name} 发送成功')
                else:
                    print(f'{name} 发送失败')
            else:
                print(f'{file} 找不到文件，请确认文件是否存在')
            sleep(sleep_t)


if __name__ == '__main__':
    sd = EmailHandle('zhouyijiazc@163.com', 'xfah206', 'smtp.163.com')
    sd.batch_send('email.xlsx')
