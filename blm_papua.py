import pandas as pd
import numpy as np
import time

pd.set_option('display.width', 1500) 
# pd.set_option('display.max_columns', 30) 
pd.set_option('display.max_columns', None) 
pd.set_option('display.max_rows', None)


qs = pd.ExcelFile('QS_BLM_PAPUA_DETAIL.xlsx')

dfk = qs.parse('Sheet1', header=None, skiprows=7)
dfk.columns = ['0','Propinsi','APBN & APBD','Baru','Perl','Penkt','SKBLM','PKS','SPM1','SP2D1','SPM2','SP2D2','SPM3','SP2D3','Progres Fisik','Amandemen PKS','BASTDANABLMAPBN','BASTKeg.','BASTPeng.SPAM','Dana','Keg.','Belum','Selesai','%','Alo.DanaLain','Real.Keg.','Y','Z','W','X','Y','Z','W','X','Y','Z','W','X','Y','Z','W','Z',]

dfk3 = dfk.loc[:, ['Propinsi', 'APBN & APBD', 'Progres Fisik']]
sp2d = dfk.loc[:, ['Propinsi', 'APBN & APBD', 'SPM1', 'SP2D1', 'SPM2', 'SP2D2', 'SPM3', 'SP2D3', 'Progres Fisik']]

def color_negative_red(val):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for negative
    strings, black otherwise.
    """
    color = 'red' if val < 80 else 'black'
    return 'color: %s' % color

red = dfk3.style.applymap(color_negative_red, subset=['Progres Fisik'])
red_sp2d = sp2d.style.applymap(color_negative_red, subset=['Progres Fisik'])

red.to_html ('QS_BLM_Papua.html', index_names=False)
sp2d.to_html ('QS_BLM_Papua_with_SP2D.html', index_names=False)        
# urut.to_excel ('rangking.xlsx')
red.to_excel ('QS_BLM_Papua.xlsx', sheet_name='Papua', index=False, header=False, startrow=1, startcol=1,)
red_sp2d.to_excel ('QS_BLM_Papua_with_SP2D.xlsx', sheet_name='rank_prov', index=False, header=False, startrow=1, startcol=1,)