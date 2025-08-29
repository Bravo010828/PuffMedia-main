import pandas as pd

# 读取你的文件
df = pd.read_excel('TotalDatabaseTemplate (27).xlsx')

# 去重：只要 Handle Name 或 Email 重复就保留第一条
df_nodup = df.drop_duplicates(subset=['Handle Name'], keep='first')
df_nodup = df_nodup.drop_duplicates(subset=['Email'], keep='first')

# 保存结果
df_nodup.to_excel('TotalDatabase去重后.xlsx', index=False)
