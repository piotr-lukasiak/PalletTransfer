from pandas import read_excel, DataFrame, ExcelWriter
from sys import argv
from pathlib import Path
completeList = []

file = Path(argv[1])

LX02DF = read_excel(file, sheet_name="LX02")
requirementDF = read_excel(file, sheet_name="Orders")

dealzStockDF = LX02DF[LX02DF["Storage Type"] == "202"]
if dealzStockDF.empty:
    dealzStockDF = LX02DF[LX02DF["Storage Type"] == 202]
u12StockDF = LX02DF[LX02DF["Storage Type"].isin(["101","102"])]
if u12StockDF.empty:
    u12StockDF = LX02DF[LX02DF["Storage Type"].isin([101,102])]
stockFiguresInDealz = dealzStockDF.groupby("Material")["Available stock"].sum()
stockFiguresInU12 = u12StockDF.groupby("Material")["Available stock"].sum()

requirementDF = DataFrame(requirementDF[requirementDF["Material"].isin(dealzStockDF["Material"])]).groupby("Material")["Total Cases"].sum()
dealzStockDF = DataFrame(dealzStockDF[dealzStockDF["Material"].isin(requirementDF.index)])

print(dealzStockDF)

dealzStockDF = dealzStockDF.sort_values(by="SLED/BBD")

requirementDF = DataFrame((requirementDF - stockFiguresInU12).dropna())
requirementDF = requirementDF[requirementDF[0]>0]

print(requirementDF)

for material in requirementDF.index:
    for line in dealzStockDF.index:
        if material == dealzStockDF["Material"].loc[line] and requirementDF[0].loc[material]>0:
            completeList.append(dealzStockDF.loc[line])
            requirementDF[0].loc[material] -= dealzStockDF["Available stock"].loc[line]


completeList = DataFrame(completeList)
writer = ExcelWriter("Unit12F_To_Unit12B_Transfer.xlsx", engine="xlsxwriter")
completeList.to_excel(writer, sheet_name="U12F")
writer.close()
print(completeList)
