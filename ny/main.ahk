#Include data\modules\includeModules.ahk



; excelData := _excelHentData(excelMock.excelMockfil, 1).getDataArray
behandletParameterData := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
test := parameterFactory.forExcelParameter(excelParameter({kolonneNavn: "Ugedage"})) 
; VlArray := vlFactory.udrulVognløb(behandletParameterData)
; test := VlArray[2][3].vognløbsdato.forventet */
; MsgBox VlArray[1][3].parameter["Vognløbsdato"].forventet
; MsgBox VlArray[2][3].parameter["Vognløbsnummer"].forventet
; VlArray[1][3].parameter["Vognløbsdato"].faktisk := "TI"
; ; VlArray[1][3].parameter["Vognløbsdato"].faktisk := "fejlIParameter"
; ; MsgBox VlArray[1][3].parameter["Vognløbsdato"].faktisk

; planet := P6Mock()

; planet.nytVognløb(VlArray[1][3])

; ; msgbox planet.vognløbsnummerForventet
; ; ; planet.vognløbsnummerFaktisk := "nyvl"
; ; msgbox planet.vognløbsnummerFaktisk

; ; MsgBox planet.vognløb.vognløbsnummerForventet
; ; MsgBox planet.vognløb.vognløbsdatoforventet
; MsgBox planet.vognløb.styresystem.forventet
; MsgBox planet.vognløb.chfKontaktnummer.forventet
; planet.nytVognløb(VlArray[2][3])
; MsgBox planet.vognløb.styresystem.forventet
; MsgBox planet.vognløb.chfKontaktnummer.forventet

return