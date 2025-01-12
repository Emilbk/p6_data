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
        jstring := jsongo.Stringify(actual)
        jobj := jsongo.Parse(jstring)

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

        
        actual := jsongo.Stringify(_excelStrukturerData(excelMock.excelDataGyldig, parameterFactory).danRækkeArray())

        this.assert.equal(jsonMock, actual)

    }
    
    antalUgedageTest(){

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

        ugyldigeKolonner := _excelVerificerData.ugyldigeKolonner
        expectedLength := 2
        actualLength := ugyldigeKolonner.Count

        this.assert.equal(actualLength, expectedLength)
    }
}
class parameterFactoryTest extends AutoHotUnitSuite{

    testAlm(){
        
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        
        actual := Type(test[1]["Vognløbsnummer"])
        expected := "parameterAlm"
        
        this.assert.equal(actual, expected)
    }
    testUgedag(){
        
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        
        actual := Type(test[1]["Ugedage"])
        expected := "parameterUgedage"
        
        this.assert.equal(actual, expected)
    }
    testUndtagneTransportTyper(){
        
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        
        actual := Type(test[1]["KørerIkkeTransportTyper"])
        expected := "parameterTransportType"
        
        this.assert.equal(actual, expected)
    }
    testKørerIkkeTransportTyper(){
        
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        
        actual := Type(test[1]["UndtagneTransportTyper"])
        expected := "parameterTransportType"
        
        this.assert.equal(actual, expected)
    }
    testKlokkeslæt(){
        
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        
        actual := Type(test[1]["Starttid"])
        expected := "parameterKlokkeslæt"
        
        this.assert.equal(actual, expected)
    }
}
class parameterTest extends AutoHotUnitSuite {

    _testUgedageFejlKalenderdagFormat() {

        parameterData := {}
        parameterData.kolonneNavn := "Ugedage"
        parameterData.parameterIndhold := ["42/11/2024"]
        testParameter := parameterFactory.forExcelParameter(excelParameter(parameterData))
        testParameter.tjekGyldighed()
        
        expected := "Fejl i kalenderdato: 42/11/2024. Skal være gyldig dato i formatet mm/dd/åååå."
        actual := testParameter.data["fejl"].fejlbesked

        this.assert.equal(actual, expected)
        

    }
    testUgedageFejlKalenderdagDato() {
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        for testUgedag in ["24-11-2024", "24.11.2024", "24112024", "24/11/24", "11/24/2024", "24/11"]
        {
            test[1]["Ugedage"].data["forventetIndholdArray"][1] := testUgedag
            test[1]["Ugedage"].tjekGyldighed()

            expected := Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", testUgedag)
            actual := test[1]["Ugedage"].data["fejl"].fejlbesked

            this.assert.equal(actual, expected)
        }
    }
    testUgedageFejlFastdag() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        for testFastDag in ["NO", "ONSDAG", "ONS"]
        {
            test[1]["Ugedage"].data["forventetIndholdArray"][1] := testFastDag
            test[1]["Ugedage"].tjekGyldighed()

            expected := Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", testFastDag)
            actual := test[1]["Ugedage"].data["fejl"].fejlbesked

            this.assert.equal(actual, expected)
        }
    }
    testParameterFejlTegnLængde() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "ForMangeTegn"
        test[1]["Vognløbsnummer"].tjekGyldighed()


        expected := "For mange tegn i parameter `"ForMangeTegn`". Nuværende 12, maks 5."
        actual := test[1]["Vognløbsnummer"].data["fejl"].fejlbesked


        this.assert.equal(actual, expected)

    }
    testParameterFejlUlovligtTegn() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "3!400"
        test[1]["Vognløbsnummer"].tjekGyldighed()

        expected := "Ulovligt tegn (`"!`") i parameter."
        actual := test[1]["Vognløbsnummer"].data["fejl"].fejlbesked


        this.assert.equal(actual, expected)

    }
    testParameterFejlArrayStørrelse() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        test[1]["KørerIkkeTransportTyper"].data["forventetIndholdArray"] := [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
        test[1]["KørerIkkeTransportTyper"].tjekGyldighed()

        expected := "For mange mange kolonner i kategori. Maks 10, nuværende 11"
        actual := test[1]["KørerIkkeTransportTyper"].data["fejl"].fejlbesked


        this.assert.equal(actual, expected)

    }

    testParameterKlokkeslætFormat() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        for testKlokkeslæt in ["2359", "23.59", "23:83", "1:30"]
        {
        test[1]["Starttid"].data["forventetIndhold"] := testKlokkeslæt
        test[1]["Starttid"].tjekGyldighed()

        expected := Format("Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat", testKlokkeslæt)
        actual := test[1]["Starttid"].data["fejl"].fejlbesked


        this.assert.equal(actual, expected)
        }
        
        test[1]["Starttid"].data["fejl"].fejlbesked := 0
        test[1]["Starttid"].data["forventetIndhold"] := "23:59"
        test[1]["Starttid"].tjekGyldighed()

        expected := 0
        actual := test[1]["Starttid"].data["fejl"].fejlbesked

        this.assert.equal(actual, expected)
        
    }
    testParameterKlokkeslætAsterisk() {

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        test[1]["Sluttid"].data["forventetIndhold"] := "22:22*"
        test[1]["Sluttid"].tjekGyldighed()

        actual := test[1]["Sluttid"].data["forventetIndhold"]
        expected := "22:22"

        this.assert.equal(actual, expected)

        actual := test[1]["Sluttid"].data["sluttidspunktErNæsteDag"]
        expected := true

        this.assert.equal(actual, expected)

    }
}
