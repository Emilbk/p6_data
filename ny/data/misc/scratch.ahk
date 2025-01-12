#Include ../modules/includeModules.ahk
; #Include ../modules/excelHentData.ahk
;#Include ../test/exelarrayMock.ahk
;#Include ../modules/gyldigeKolonner.ahk
;#Include ../modules/parameter.ahk

; dataRække := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
; vlRække := vlFactory.udrulVognløb(dataRække)

; actual := vlRække[2][3].Vognløbsdatoforventet
; expected := "TI"

test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
for testFastDag in ["NO", "ONSDAG", "ONS"] {
    test[1]["Ugedage"].data["forventetIndholdArray"][1] := testFastDag
    test[1]["Ugedage"].tjekGyldighed()

    expected := Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", testFastDag)
    actual := test[1]["Ugedage"].data["fejl"].fejlbesked
}
return