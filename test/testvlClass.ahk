#Requires AutoHotkey v2.0

#Include ../include.ahk

excelfil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
excel := excelIndlæsVlData(excelfil, 1)
vlFraExcel := excel.getVlArray()

vlConstruct := VognløbConstructor(vlFraExcel)
vlContainer := vlConstruct.getBehandletVognløbsArray()



return
