from openpyxl import load_workbook
import pandas as pd 
import numpy as np
from itertools import islice

wb = load_workbook('sample.xlsx')
ws = wb['Sheet1']

samp = ws.values
columns = next(samp)[1:]
samp = list(samp)
idx = [r[0] for r in samp]
samp = (islice(r, 1, None) for r in samp)
df = pd.DataFrame(samp, index=idx, columns=columns)

ws2 = wb['Sheet2']

data = ws2.values
columns = next(data)[1:]
data = list(data)
idx = [r[0] for r in data]
data = (islice(r, 1, None) for r in data)
df2 = pd.DataFrame(data, index=idx, columns=columns)

print(df, '\n', df2)


def findPosInDf(dfIn,findme):
    positions = []
    irow =0
    while ( irow < len(dfIn.index)):
        list_colPositions=dfIn.columns[dfIn.iloc[irow,:]==findme].tolist()   
        if list_colPositions != []:
            colu_iloc = dfIn.columns.get_loc(list_colPositions[0])
            positions.append([irow  , colu_iloc ])
            irow +=1

    return positions


print(findPosInDf(df, 'I'))

