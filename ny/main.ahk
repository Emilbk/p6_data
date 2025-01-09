#Include data\modules\includeModules.ahk


; excelData := _excelHentData(excelMock.excelMockfil, 1)
data := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker


data[1]["Vognløbskategori"].data["forventetIndhold"] := "testset"

return