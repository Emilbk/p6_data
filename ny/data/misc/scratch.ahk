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
tIn := FileRead("../modules/gyldigeKolonner/gyldigeKolonner.json")
t := jsongo.Parse(tin)

t2 := json.Load(tIn)


t3 := json.load(gyldigKolonneJson.data)
return