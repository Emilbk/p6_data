#Include ../modules/includeModules.ahk
#Include tests.ahk
FileEncoding "UTF-8"
; #Include ../modules/excelClass.ahk
; #Include excel.mock.ahk


class testExcelHentData extends AutoHotUnitSuite {

    _excelArrayFraFil() {


        testFil := A_WorkingDir "\assets\VLMock.xlsx"

        actual := _excelHentData(testFil, 1).excelDataArray
        this.assert.equal(actual[1][1], "Budnummer")
        this.assert.equal(actual[2][2], "31400")
        jstring := jsongo.Stringify(actual)
        jobj := jsongo.Parse(jstring)

        return
    }

    _excelSpeedTest() {

        A_WorkingDir := "../"
        testFil := A_WorkingDir "\assets\150vl.xlsx"
        app := _excelHentData(testFil, 1)
        loop 30 {
            Timer.add("exceltest")
            actual := app.excelDataArray
        }
        app._quit()
        Timer.show()
    }

}

class testExcelDataStruktur extends AutoHotUnitSuite {

    arrayTest() {

        jsonMock := FileRead("json/excelDataMockArray.txt", "UTF-8")

        actual := jsongo.Stringify(_excelStrukturerData(mock).danRækkeArray())

        this.assert.equal(jsonMock, actual)

    }

}

class testExcelVerificerData extends AutoHotUnitSuite {

    testVerificerUgyldigeKolonner() {

        ugyldigeKolonner := _excelVerificerData.ugyldigeKolonner
        expectedLength := 2
        actualLength := ugyldigeKolonner.Count

        this.assert.equal(actualLength, expectedLength)
    }
}


class excelParameter extends AutoHotUnitSuite {

    testUgedageFejlKalenderdagFormat() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Ugedage"].data["forventetIndholdArray"][1] := "42/11/2024"
        test[1]["Ugedage"].tjekGyldighed()

        expected := "Fejl i kalenderdato: 42/11/2024. Skal være gyldig dato i formatet mm/dd/åååå."
        actual := test[1]["Ugedage"].data["fejlBesked"]

        this.assert.equal(actual, expected)

    }
    testUgedageFejlKalenderdagDato() {
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        for testUgedag in ["24-11-2024", "24.11.2024", "24112024", "24/11/24", "11/24/2024", "24/11"]
        {
            test[1]["Ugedage"].data["forventetIndholdArray"][1] := testUgedag
            test[1]["Ugedage"].tjekGyldighed()

            expected := Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", testUgedag)
            actual := test[1]["Ugedage"].data["fejlBesked"]

            this.assert.equal(actual, expected)
        }
    }
    testUgedageFejlFastdag() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        for testFastDag in ["NO", "ONSDAG", "ONS"]
        {
            test[1]["Ugedage"].data["forventetIndholdArray"][1] := testFastDag
            test[1]["Ugedage"].tjekGyldighed()

            expected := Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", testFastDag)
            actual := test[1]["Ugedage"].data["fejlBesked"]

            this.assert.equal(actual, expected)
        }
    }
    testParameterFejlTegnLængde() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "ForMangeTegn"
        test[1]["Vognløbsnummer"].tjekGyldighed()

        expected := "For mange tegn i parameter `"ForMangeTegn`". Nuværende 12, maks 5."
        actual := test[1]["Vognløbsnummer"].data["fejlBesked"]


        this.assert.equal(actual, expected)

    }
    testParameterFejlUlovligtTegn() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "3!400"
        test[1]["Vognløbsnummer"].tjekGyldighed()

        expected := "Ulovligt tegn (`"!`") i parameter."
        actual := test[1]["Vognløbsnummer"].data["fejlBesked"]


        this.assert.equal(actual, expected)

    }
    testParameterFejlArrayStørrelse() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["KørerIkkeTransporttyper"].data["forventetIndholdArray"] := [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
        test[1]["KørerIkkeTransporttyper"].tjekGyldighed()

        expected := "For mange mange kolonner i kategori. Maks 10, nuværende 11"
        actual := test[1]["KørerIkkeTransporttyper"].data["fejlBesked"]


        this.assert.equal(actual, expected)

    }

    testParameterKlokkeslætFormat() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        for testKlokkeslæt in ["2359", "23.59", "23:83", "1:30"]
        {
        test[1]["Starttid"].data["forventetIndhold"] := testKlokkeslæt
        test[1]["Starttid"].tjekGyldighed()

        expected := Format("Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat", testKlokkeslæt)
        actual := test[1]["Starttid"].data["fejlBesked"]


        this.assert.equal(actual, expected)
        }
        
        test[1]["Starttid"].data["fejlBesked"] := 0
        test[1]["Starttid"].data["forventetIndhold"] := "23:59"
        test[1]["Starttid"].tjekGyldighed()

        expected := 0
        actual := test[1]["Starttid"].data["fejlBesked"]

        this.assert.equal(actual, expected)
        
    }
    testParameterKlokkeslætAsterisk() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Sluttid"].data["forventetIndhold"] := "22:22*"
        test[1]["Sluttid"].tjekGyldighed()

        expected := "22:22"
        actual := test[1]["Sluttid"].data["forventetIndhold"]

        this.assert.equal(actual, expected)

        expected := true
        actual := test[1]["Sluttid"].data["sluttidspunktErNæsteDag"]

        this.assert.equal(actual, expected)


    }
}