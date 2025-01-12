#Include ../modules/includeModules.ahk


class vltest extends AutoHotUnitSuite{

    testVognløbsrække(){

        dataRække := excelDataBehandler(excelMock.excelDataGyldig, parameterFactory).behandledeRækker
        vlRække := vlFactory.udrulVognløb(dataRække)
        
        actual := vlRække[2][3].VognløbsdatoForventet 
        expected := "TI"
        
        
       this.assert.equal(actual, expected) 
        return
    }

} 