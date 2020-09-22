import pandas as pd
import pickle
import xlwings as xw

pd.options.display.float_format = '{:,.1f}'.format
pd.set_option('min_rows', 20)
pd.set_option('max_rows', 100)
pd.set_option('max_column', 10)
pd.set_option('max_colwidth', 20)
pd.set_option('display.width', 1000)

# dp = pd.read_csv('PrTest.csv')
# проставить МУКи по компаниям
# dp.loc[dp['org'] == 'ТСК Мосэнерго', 'MUUKK'] = 10000
# dp.loc[dp['org'] == 'ТЭР', 'MUUKK'] = 15000
# dp.loc[dp['org'] == 'ТГК-Сервис', 'MUUKK'] = 3500
# dp.loc[dp['org'] == 'Метрология', 'MUUKK'] = 6000
# dp.loc[dp['org'] == 'ГЭХ Инжиниринг', 'MUUKK'] = 7000
# dp.loc[dp['org'] == 'МТЭР', 'MUUKK'] = 2000
# dp.loc[dp['org'] == 'МП-Проектстрой', 'MUUKK'] = 641
# dp.loc[dp['org'] == 'ТЭР-Сервис', 'MUUKK'] = 2000
# dp.loc[dp['org'] == 'ГЭР', 'MUUKK'] = 600
# dp.loc[dp['org'] == 'МРЭС', 'MUUKK'] = 400
# dp.loc[dp['org'] == 'ИТЦ', 'MUUKK'] = 238
# dp.loc[dp['org'] == 'ЕрмакНГ', 'MUUKK'] = 400
# dp.loc[dp['org'] == 'ЦРМЗ', 'MUUKK'] = 83
# dp.loc[dp['org'] == 'РИК', 'MUUKK'] = 2

# dp.loc[dp['MUUKK'] < dp['sum'].abs(), 'muk'] = 1
# dp = dp.fillna(0)
# dp.to_csv('PrTest.csv', index=False)

# формирование ГК
dp = pd.read_csv('PrTest.csv')
# tb = pd.pivot_table(dp, values=['sum', 'data'],
#                     index=['org', 'schD2', 'schK2'],
#                     columns=['muk'],
#                     aggfunc={'sum': sum, 'data': 'count'})
# tb.reset_index(inplace=True)
# tb = tb.fillna(0)
# tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')
#
# with open('prVn.pickle', 'rb') as f: prVn = pickle.load(f)
# with open('pr.pickle', 'rb') as f: pr = pickle.load(f)
# tb["vn"] = tb["DK"].map(prVn)
# tb["opis"] = tb["DK"].map(pr)
# tb = tb.fillna(0)
#
# xlapp = xw.apps.active
# rng = xlapp.selection
# rng.options(index=False).value = tb

# делаем помесячную ГК
tb = pd.pivot_table(dp, values=['sum', 'data'],
                    index=['org', 'schD2', 'schK2'],
                    columns=['god', 'mes'],
                    aggfunc={'sum': sum, 'data': 'count'}
                    , margins=True
                    )
tb.reset_index(inplace=True)
tb = tb.fillna(0)
tb['DK'] = tb.schD2.astype(str).str.cat(tb.schK2.astype(str), sep=';')

xlapp = xw.apps.active
rng = xlapp.selection
rng.options(index=False).value = tb

# tb.to_excel("output.xlsx")
# print(tb)
# print(dp.head(5))
# print(dp.columns)
# print(dp.shape)
# print(tb.dtypes)
