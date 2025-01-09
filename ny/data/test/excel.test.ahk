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

        ugyldigeKolonner := _excelVerificerData(excelDataUgyldigMock).ugyldigeKolonner
        expectedLength := 2
        actualLength := ugyldigeKolonner.Count
        
        this.assert.equal(actualLength, expectedLength)
    }
}