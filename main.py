#%%
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

pd.options.mode.chained_assignment = None

#%% 
#Подготовка данных

df = pd.read_excel("data.xlsx")

cols = df.columns

df.drop(columns=cols[5], inplace=True) # Удаление Unnamed 5, пустого столбца 

indices = df[df[cols[0]].isnull()].index.tolist()
indices.append(len(df))

df['month'] = np.nan
for idx in range(len(indices)-1):
    df.loc[indices[idx]:indices[idx+1], 'month'] = idx + 5    

df.drop(index=indices[:-1], inplace=True)
df['receiving_date'] = pd.to_datetime(df['receiving_date'], errors='coerce', format='%Y-%m-%d')

#%%
#Задание 1

june_total = df.loc[(df['month'] == 7) & ~(df['status'] == 'ПРОСРОЧЕНО')] \
    ['sum'].sum()

print('Задание 1')
print(f'Общая выручка за июль по сделкам, приход средств которых не просрочено: {round(june_total, 2)} \n')  

#%%
#Задание 2

df.month = df.month.astype('int')

revenue = df.loc[~(df['status'] == 'ПРОСРОЧЕНО')].groupby('month')['sum'].sum() / 1000
revenue.plot(kind='line', title='Задание 2, общая выручка по месяцам', xlabel='Месяц', ylabel='Сумма (в тыс.)')
plt.show()

#%%
#Задание 3

revenue_by_man = df.loc[~(df['status'] == 'ПРОСРОЧЕНО') & (df['month'] == 9)].groupby('sale')['sum'].sum().sort_values(ascending=False)

print('Задание 3')
print(f'В сентябре больше всего привлёк средств менеджер {revenue_by_man.index[0]}, сумма: {float(revenue_by_man.iloc[0])} \n')

#%%
#Задание 4

october = df.loc[~(df['status'] == 'ПРОСРОЧЕНО') & (df['month'] == 10)]['new/current'].value_counts().sort_values(ascending=False)

print('Задание 4')
print('Кол-во текущих и новых сделок в октябре:')
print(october)
print(f'В октябре больше всего следующего типа сделок: {october.index[0]}, кол-во: {int(october.iloc[0])} \n')
#%%
#Задание 5

may = df.loc[(df['document'] == 'оригинал') & (df['month'] == 5) & ~(df['receiving_date'].isnull())]
june_doc_num = len(may.loc[may['receiving_date'].dt.month == 6])

print('Задание 5')
print(f'Кол-во майских оригиналов, полученных в июне: {june_doc_num} \n')
#%%
#Бонусное задание

remainders = df.loc[(df['document'] == 'оригинал') & 
    (df['status'] != 'ПРОСРОЧЕНО') &
    (df['receiving_date'].dt.month == 7) & 
    (df['month'] < 7) ]

remainders['remainder'] = remainders['sum'].apply(lambda x: x*.03 if x < 10000 else x*.05)
result = remainders.groupby('sale')['remainder'].sum().sort_index(ascending=True)

print('Бонусное задание')
print('Остаток бонусов для менеджеров на июль:')
print(result)