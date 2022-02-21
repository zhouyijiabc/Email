import pandas as pd


columns = ['邮箱', '名字', '主题', '内容', '文件']
df = pd.DataFrame(columns=columns)
df.to_excel('email.xlsx', index=False)
