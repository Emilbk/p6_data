#Include ../modules/includeModules.ahk
#Include ../lib/cJSON.ahk
#include ../modules/gyldigeKolonner/gyldigeKolonnerJson.ahk
; #Include ../modules/excelHentData.ahk
;#Include ../test/exelarrayMock.ahk
;#Include ../modules/gyldigeKolonner.ahk
;#Include ../modules/parameter.ahk

; dataRække := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
; vlRække := vlFactory.udrulVognløb(dataRække)

; actual := vlRække[2][3].Vognløbsdatoforventet
; expected := "TI"
ugyldigKolonne := "test"
actual := gyldigKolonneJson.erGyldigKolonne(ugyldigKolonne)

return