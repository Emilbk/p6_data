#Include ../modules/excelHentData.ahk
#Include ../modules/excelVerificerData.ahk
#Include tests.ahk
; #Include ../modules/excelClass.ahk
; #Include excel.mock.ahk


class testExcelHentData extends AutoHotUnitSuite {

    _excelArrayFraFil() {


        testFil := A_WorkingDir "\assets\VLMock.xlsx"

        actual := excelHentData(testFil, 1).excelDataArray
        this.assert.equal(actual[1][1], "Budnummer")
        this.assert.equal(actual[2][2], "31400")
        jstring := jsongo.Stringify(actual)
        jobj := jsongo.Parse(jstring)

        return
    }

    _excelSpeedTest() {

        A_WorkingDir := "../"
        testFil := A_WorkingDir "\assets\150vl.xlsx"
        app := excelHentData(testFil, 1)
        loop 30 {
            Timer.add("exceltest")
            actual := app.excelDataArray
        }
        app._quit()
        Timer.show()
    }

}

class testExcelVerificerData extends AutoHotUnitSuite {
   
    testVerificerUgyldigeKolonner(){

        ugyldigeKolonner := excelVerificerData(excelDataUgyldigMock).ugyldigeKolonner
        expectedLength := 2
        actualLength := ugyldigeKolonner.Count
        
        this.assert.equal(actualLength, expectedLength)
    }
}