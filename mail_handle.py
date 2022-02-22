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
        sleep_t = random.random()
        df = pd.read_excel(excel_file)
        for index, row in df.iterrows():
            if os.path.isfile(row['文件名']):
                if self.send_mail(row['邮箱'], subject=row['主题'], content=row['内容'], file=row['文件名']):
                    print(f'{row["名字"]} 发送成功')
                else:
                    print(f'{row["名字"]} 发送失败')
            else:
                print(f'{row["文件名"]} 找不到文件，请确认文件是否存在')
            sleep(sleep_t)
            print(index, row['名字'])


if __name__ == '__main__':
    sd = EmailHandle('zhouyijiazc@163.com', 'xfah206', 'smtp.163.com')
    sd.batch_send('email.xlsx')
