#Include data\modules\includeModules.ahk

; parameterData := {}
; parameterData.parameterIndhold := ["24-11-2024", "24.11.2024", "24112024", "24/11/24", "11/24/2024", "24/11"]
; parameterData.kolonneNavn := "Ugedage"
; parameterData.rækkeIndex := 1
; testParameter := parameterFactory.forExcelParameter(excelParameter(parameterData))
; testParameter.tjekGyldighed()P

; t := _excelStrukturerData(excelMock.excelDataGyldig, parameterFactory)
; t.danRækkeArray()

; t := parameterFactory.forExcelParameter(excelParameter({ kolonneNavn: "Ugedage" }))

; test := excelDataB dataRække := exc ehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

    test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
    test[1]["Vognløbsnummer"].data["forventetIndhold"] := "ForMangeTegn"
    test[1]["Vognløbsnummer"].tjekGyldighed()

    expected := "For mange tegn i parameter `"ForMangeTegn`". Nuværende 12, maks 5."
    actual := test[1]["Vognløbsnummer"].fejlObj["fejlBesked"]
return