#Include data\modules\includeModules.ahk



excelData := _excelHentData(excelMock.excelMockfil, 1).getDataArray
behandletParameterData := excelDataBehandler(excelData, parameterFactory).behandledeRækker
VlArray := vlFactory.udrulVognløb(behandletParameterData)

MsgBox VlArray[1][3]["Vognløbsdato"].forventet
MsgBox VlArray[2][3]["Vognløbsnummer"].forventet
VlArray[1][3]["Vognløbsdato"].faktisk := "indlæstFraP6"
MsgBox VlArray[1][3]["Vognløbsdato"].faktisk

return