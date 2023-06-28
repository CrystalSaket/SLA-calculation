import numpy as np
import pandas as pd

df = pd.read_excel('SLA_LS_2023_06_23.xlsx')

df['day'] = df['MI Date'].dt.day
df['month'] = df['MI Date'].dt.month
df['year'] = df['MI Date'].dt.year

t = pd.Timestamp.now()
t.to_datetime64()
curDay = t.day
curMonth = t.month
curYear = t.year


actual = []
# df['actual'] = df.sum(axis=1,numeric_only=True)
cnt = 0
for index,row in df.iterrows():
    cnt = cnt + 1
    if cnt % 1000 == 0:
        print(f'Done : {cnt}')
    x = 0
    if row['month'] == curMonth and row['year'] == curYear:
        x = row.iloc[row['day']*3 + 2:-3].sum()
    else:
        x = row.iloc[2:-3].sum()
    actual.append(x)
df['actual'] = actual

expected = []   
cnt = 0
for index,row in df.iterrows():
    cnt = cnt + 1
    if cnt % 1000 == 0:
        print(f'Done : {cnt}')
    x = 0
    if row['month'] == curMonth and row['year'] == curYear:
        x = (curDay - row['day'] - 1)*3*48
    else:
        
        x = (curDay - 1)*3*48
    expected.append(x)
df['expected'] = expected

df['SLA'] = (df['actual'] / df['expected']) * 100
df.to_excel('SLA_LS_Calculated_2023_06_23.xlsx',sheet_name='Sheet1',index=False)

    