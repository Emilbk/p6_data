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
    
    arrayTest(){

        jsonMock := FileRead("json/excelDataMockArray.txt", "UTF-8")
        
        actual := jsongo.Stringify(_excelStrukturerData(mock).danRækkeArray())
        
        this.assert.equal(jsonMock, actual)

    }

}

class testExcelVerificerData extends AutoHotUnitSuite {
   
    testVerificerUgyldigeKolonner(){

        ugyldigeKolonner := _excelVerificerData.ugyldigeKolonner
        expectedLength := 2
        actualLength := ugyldigeKolonner.Count
        
        this.assert.equal(actualLength, expectedLength)
    }
}


class excel extends AutoHotUnitSuite {

    testUgedageFejlKalenderdagFormat(){

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Ugedage"].data["forventetIndholdArray"][1] := "42/11/2024"
        test[1]["Ugedage"].tjekGyldighed()

        expected := "Fejl i kalenderdato: 42/11/2024. Skal være gyldig dato i formatet mm/dd/åååå."
        actual := test[1]["Ugedage"].data["fejlBesked"]
        
        this.assert.equal(actual,expected)

    }
    testUgedageFejlKalenderdagDato(){
        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Ugedage"].data["forventetIndholdArray"][1] := "24-11-2024"
        test[1]["Ugedage"].tjekGyldighed()

        expected := "Fejl i kalenderdato: 24-11-2024. Skal være gyldig dato i formatet mm/dd/åååå."
        actual := test[1]["Ugedage"].data["fejlBesked"]
        
        this.assert.equal(actual,expected)
    }
    testUgedageFejlfastdag(){

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Ugedage"].data["forventetIndholdArray"][1] := "NO"
        test[1]["Ugedage"].tjekGyldighed()

        expected := "fejl i fast dag: NO. Skal være i formatet XX, f. eks MA"
        actual := test[1]["Ugedage"].data["fejlBesked"]
        
        this.assert.equal(actual,expected)

    }
    testParameterFejlTegnLængde(){

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "ForMangeTegn"
        test[1]["Vognløbsnummer"].tjekGyldighed()

        expected := "For mange tegn i parameter."
        actual := test[1]["Vognløbsnummer"].data["fejlBesked"]
        
        
        this.assert.equal(actual,expected)

    }
    testParameterFejlUlovligtTegn(){

        test := excelDataBehandler(excelMock.excelDataGyldig, parameterAlm).behandledeRækker
        test[1]["Vognløbsnummer"].data["forventetIndhold"] := "3!400"
        test[1]["Vognløbsnummer"].tjekGyldighed()

        expected := "Ulovligt tegn (`"!`") i parameter."
        actual := test[1]["Vognløbsnummer"].data["fejlBesked"]
        
        
        this.assert.equal(actual,expected)

    }
}