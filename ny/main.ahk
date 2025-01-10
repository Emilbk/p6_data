#Include data\modules\includeModules.ahk


; excelData := _excelHentData(excelMock.excelMockfil, 1)
data := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker


data[1]["Ugedage"].data["forventetIndholdArray"] := ["11/24"]

return