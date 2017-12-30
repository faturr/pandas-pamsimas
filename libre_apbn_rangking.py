import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
# import time

pd.set_option('display.width', 1500) 
# pd.set_option('display.max_columns', 30) 
pd.set_option('display.max_columns', None) 
pd.set_option('display.max_rows', None)


qs = pd.ExcelFile('Quick_Status_30-12-2017_APBN_REKAP.xlsx')
# df = qs.parse('Sheet1', header=None, usecols=[1,2,14])

# dfk = qs.parse('Sheet1', Index_col=0, header=None)

dfk = qs.parse('Sheet1', header=None)
dfk.columns = ['0','Propinsi','Desa APBN','Baru','Perl','Penkt','SKBLM','PKS','SPM1','SP2D1','SPM2','SP2D2','SPM3','SP2D3','Progres Fisik','Amandemen PKS','BASTDANABLMAPBN','BASTKeg.','BASTPeng.SPAM','Dana','Keg.','Belum','Selesai','%','Alo.DanaLain','Real.Keg.','Y','Z','W','X','Y','Z','W','X','Y','Z','W','X','Y','Z','W','Z',]

dfk3 = dfk.loc[:, ['Propinsi', 'Desa APBN', 'Progres Fisik']]
sp2d = dfk.loc[:, ['Propinsi', 'Desa APBN', 'SKBLM', 'PKS', 'SPM1', 'SP2D1', 'SPM2', 'SP2D2', 'SPM3', 'SP2D3', 'Progres Fisik']]

# df = qs.parse('Sheet1', header=None, usecols=[1,2,14])

# df1 = qs.parse('Sheet1', header=None, usecols=[1,2,14,8,9,10,11,12,13])

prov = dfk3.loc[[5, 24, 48, 61, 71, 81, 95, 105, 118, 125, 131, 147, 176, 181, 209, 214, 220, 229, 249, 262, 272, 284, 289, 294, 304, 317, 337, 349, 354, 360, 369, 376, 381], :]
prov_sp2d = sp2d.loc[[5, 24, 48, 61, 71, 81, 95, 105, 118, 125, 131, 147, 176, 181, 209, 214, 220, 229, 249, 262, 272, 284, 289, 294, 304, 317, 337, 349, 354, 360, 369, 376, 381], :]

urut = prov.sort_values('Progres Fisik')
urut_sp2d = prov_sp2d.sort_values('Progres Fisik')

urut.to_html ('QS_APBN_Nasional.html', index_names=False)
urut_sp2d.to_html ('QS_APBN_Nasional_with_SP2D.html', index_names=False)        
# urut.to_excel ('rangking.xlsx')
urut.to_excel ('QS_APBN_Nasional.xlsx', sheet_name='rank_prov', index=False, header=False, startrow=1, startcol=1,)
urut_sp2d.to_excel ('QS_APBN_Nasional_with_SP2D.xlsx', sheet_name='rank_prov', index=False, header=False, startrow=1, startcol=1,)