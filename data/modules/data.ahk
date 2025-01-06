#Include includeModules.ahk

class data {

    hentExcelData(pExceldata){

        this.dataArray := pExceldata
    }

}

test := data()
test.hentExcelData(excelDataMock)