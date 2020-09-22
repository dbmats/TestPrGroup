import pandas as pd
# собираем несколько файлов в один csv.

# df = pd.read_csv('PrTest.csv')
# dp = pd.read_csv('PrTest5.csv')
# df = df.append(dp)
# dp = pd.read_csv('PrTest3.csv')
# df = df.append(dp)
# dp = pd.read_csv('PrTest4.csv')
# df = df.append(dp)
# df.to_csv('PrTest.csv', index=False)

df = pd.read_excel('1.xlsx', sheet_name='1')
dp = pd.read_excel('2.xlsx', sheet_name='1')
df = df.append(dp)
dp = pd.read_excel('3.xlsx', sheet_name='1')
df = df.append(dp)


df = df.fillna(0)
df.to_csv('PrTest.csv', index=False)

print(df.shape)
print(df.dtypes)
