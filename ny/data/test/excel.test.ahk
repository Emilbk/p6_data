#Include ../modules/includeModules.ahk
#Include tests.ahk
FileEncoding "UTF-8"
; #Include ../modules/excelClass.ahk
; #Include excel.mock.ahk

class testExcelHentData extends AutoHotUnitSuite {

    _excelArrayFraFil() {

        testFil := A_WorkingDir "\assets\VLMock.xlsx"

        actual := _excelHentData(testFil, 1).getDataArray
        this.assert.equal(actual[1][1], "Budnummer")
        this.assert.equal(actual[2][2], "31400")
        jstring := json.dump(actual)
        jobj := json.load(jstring)

        return
    }

    _excelSpeedTest() {

        testFil := A_WorkingDir "\assets\150vl.xlsx"
        app := _excelHentData(testFil, 1)
        loop 30 {
            Timer.add("exceltest")
            actual := app.getDataArray
        }
        app._quit()
        Timer.show()
    }

}

class testExcelDataStruktur extends AutoHotUnitSuite {

    _arrayTest() {

        jsonMock := FileRead("json/excelDataMockArray.json", "UTF-8")

        actual := json.dump(_excelStrukturerData(excelMock.excelDataGyldig, parameterFactory).danRækkeArray())

        this.assert.equal(jsonMock, actual)

    }

    antalUgedageTest() {

        data := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

        actual := data[1]["Ugedage"].forventet.length
        expected := 8
        this.assert.equal(actual, expected)
        actual := data[2]["Ugedage"].forventet[1]
        expected := "18/11/2024"
        this.assert.equal(actual, expected)

    }

}

class testExcelVerificerData extends AutoHotUnitSuite {

    testVerificerUgyldigeKolonner() {

        ugyldigKolonne := "ikkeGyldig"
        expected := false
        actual := gyldigKolonneJson.erGyldigKolonne(ugyldigKolonne)

        this.assert.equal(actual, expected)
    }

    testGyldigKolonne() {

        gyldigKolonne := gyldigKolonneJson.data

        return

    }
}
; TODO omskriv til direkte brug af paramaterfactory
class parameterFactoryTest extends AutoHotUnitSuite {

    testAlm() {

        test := parameterFactory.forExcelParameter(excelParameter({ kolonneNavn: "Vognløbsnummer" }))

        actual := Type(test)
        expected := "parameterAlm"

        this.assert.equal(actual, expected)
    }
    testUgedag() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

        actual := Type(test[1]["Ugedage"])
        expected := "parameterUgedage"

        this.assert.equal(actual, expected)
    }
    testUndtagneTransportTyper() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

        actual := Type(test[1]["KørerIkkeTransportTyper"])
        expected := "parameterTransportType"

        this.assert.equal(actual, expected)
    }
    testKørerIkkeTransportTyper() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

        actual := Type(test[1]["UndtagneTransportTyper"])
        expected := "parameterTransportType"

        this.assert.equal(actual, expected)
    }
    testKlokkeslæt() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker

        actual := Type(test[1]["Starttid"])
        expected := "parameterKlokkeslæt"

        this.assert.equal(actual, expected)
    }
}
; TODO omskriv til direkte bruge af parameter.
class parameterTest extends AutoHotUnitSuite {

    testUgedageFejlKalenderdagFormat() {

        parametre := {
            kolonneNavn: "Ugedage",
            parameterNavn: "Ugedage",
            parameterIndhold: ["42/11/2024"]
        }
        test := parameterFactory.forExcelParameter(excelParameter(parametre))

        expected := "Fejl i kalenderdato: 42/11/2024. Skal være gyldig dato i formatet mm/dd/åååå."
        actual := test.fejlObj["fejlBesked"]

        this.assert.equal(actual, expected)

    }
    ; TODO lav testmodul, der samler i array, ikke blot sender seneste.
    testUgedageFejlKalenderdagDato() {
        for dato in [["24-11-2024"], ["24.11.2024"], ["24112024"], ["24/11/24"], ["11/24/2024"], ["24/11"]] {
            parametre := {
                kolonneNavn: "Ugedage",
                parameterNavn: "Ugedage",
                parameterIndhold: dato
            }
            test := parameterFactory.forExcelParameter(excelParameter(parametre))

            expected := Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", dato[1])
            actual := test.fejlObj["fejlBesked"]
        }

        this.assert.equal(actual, expected)

    }
    testUgedageFejlFastdag() {
        for testFastDag in [["NO"], ["ONSDAG"], ["ONS"]] {
            parametre := {
                kolonneNavn: "Ugedage",
                parameterNavn: "Ugedage",
                parameterIndhold: testFastDag
            }
            test := parameterFactory.forExcelParameter(excelParameter(parametre))

            expected := Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", testFastDag[1])
            actual := test.fejlObj["fejlBesked"]

            this.assert.equal(actual, expected)
        }
    }
    testParameterFejlTegnLængde() {

        parametre := {
            kolonneNavn: "Vognløbsnummer",
            parameterNavn: "Vognløbsnummer",
            parameterIndhold: "ForMangeTegn"
        }
        test := parameterFactory.forExcelParameter(excelParameter(parametre))

        expected := "For mange tegn i parameter `"ForMangeTegn`". Nuværende 12, maks 5."
        actual := test.fejlObj["fejlBesked"]

        this.assert.equal(actual, expected)

    }
    testParameterFejlUlovligtTegn() {
        parametre := {
            kolonneNavn: "Vognløbsnummer",
            parameterNavn: "Vognløbsnummer",
            parameterIndhold: "3!400"
        }
        test := parameterFactory.forExcelParameter(excelParameter(parametre))

        expected := "Ulovligt tegn (`"!`") i parameter."
        actual := test.fejlObj["fejlBesked"]

        this.assert.equal(actual, expected)

    }
    testParameterFejlArrayStørrelse() {

        parametre := {
            kolonneNavn: "KørerIkkeTransportTyper",
            parameterNavn: "KørerIkkeTransportTyper",
            parameterIndhold: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
        }
        test := parameterFactory.forExcelParameter(excelParameter(parametre))

        expected := "For mange mange kolonner i kategori. Maks 10, nuværende 11"
        actual := test.fejlObj["fejlBesked"]

        this.assert.equal(actual, expected)

    }

    testParameterKlokkeslætFormat() {

        for testKlokkeslæt in ["2359", "23.59", "23:83", "1:30"] {

            parametre := {
                kolonneNavn: "Starttid",
                parameterNavn: "Starttid",
                parameterIndhold: testKlokkeslæt
            }
            test := parameterFactory.forExcelParameter(excelParameter(parametre))

            expected := Format(
                "Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat",
                testKlokkeslæt)
            actual := test.fejlObj["fejlBesked"]

            this.assert.equal(actual, expected)
        }

    }
    testParameterKlokkeslætAsterisk() {

        parametre := {
            kolonneNavn: "Starttid",
            parameterNavn: "Starttid",
            parameterIndhold: "04:00*"
        }

        test := parameterFactory.forExcelParameter(excelParameter(parametre))

        actual := test.data["forventetIndhold"]
        expected := "04:00"

        this.assert.equal(actual, expected)

        actual := test.data["sluttidspunktErNæsteDag"]
        expected := true

        this.assert.equal(actual, expected)

    }
}
