#Include ../modules/includeModules.ahk
; #Include ../modules/excelHentData.ahk
;#Include ../test/exelarrayMock.ahk
;#Include ../modules/gyldigeKolonner.ahk
;#Include ../modules/parameter.ahk

dataRække := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
vlRække := vlFactory.udrulVognløb(dataRække)

actual := vlRække[2][3].Vognløbsdatoforventet
expected := "TI"

return