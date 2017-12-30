
import pandas as pd
import numpy as np
# import matplotlib.pyplot as plt
# import time

pd.set_option('display.width', 1500) 
# pd.set_option('display.max_columns', 30) 
pd.set_option('display.max_columns', None) 
pd.set_option('display.max_rows', None)


qs = pd.ExcelFile('Quick_Status_30-12-2017_BLM_REKAP.xlsx')

dfk = qs.parse('Sheet1', header=None)
dfk.columns = ['0','Propinsi','APBN & APBD','Baru','Perl','Penkt','SKBLM','PKS','SPM1','SP2D1','SPM2','SP2D2','SPM3','SP2D3','Progres Fisik','Amandemen PKS','BASTDANABLMAPBN','BASTKeg.','BASTPeng.SPAM','Dana','Keg.','Belum','Selesai','%','Alo.DanaLain','Real.Keg.','Y','Z','W','X','Y','Z','W','X','Y','Z','W','X','Y','Z','W','Z',]

dfk3 = dfk.loc[:, ['Propinsi', 'APBN & APBD', 'Progres Fisik']]
sp2d = dfk.loc[:, ['Propinsi', 'APBN & APBD', 'SKBLM', 'PKS', 'SPM1', 'SP2D1', 'SPM2', 'SP2D2', 'SPM3', 'SP2D3', 'Progres Fisik']]


prov = dfk3.loc[[5, 24, 48, 61, 71, 81, 95, 105, 118, 125, 131, 147, 176, 181, 209, 214, 220, 229, 249, 262, 272, 284, 289, 294, 304, 317, 338, 351, 356, 362, 372, 379, 385], :]
prov_sp2d = sp2d.loc[[5, 24, 48, 61, 71, 81, 95, 105, 118, 125, 131, 147, 176, 181, 209, 214, 220, 229, 249, 262, 272, 284, 289, 294, 304, 317, 338, 351, 356, 362, 372, 379, 385], :]


urut = prov.sort_values('Progres Fisik')
urut_sp2d = prov_sp2d.sort_values('Progres Fisik')

urut.to_html ('QS_BLM_Nasional.html', index_names=False)
urut_sp2d.to_html ('QS_BLM_Nasional_with_SP2D.html', index_names=False)        
# urut.to_excel ('rangking.xlsx')
urut.to_excel ('QS_BLM_Nasional.xlsx', sheet_name='rank_prov', index=False, header=False, startrow=1, startcol=1,)
urut_sp2d.to_excel ('QS_BLM_Nasional_with_SP2D.xlsx', sheet_name='rank_prov', index=False, header=False, startrow=1, startcol=1,)