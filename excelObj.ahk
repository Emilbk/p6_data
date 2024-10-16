#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent
class excelObj extends Class
{

    excelFil := Object()
    
    vælgExcelFil(pExcelFilPathLong)
    {
        this.excelFil.NavnLong := pExcelFilPathLong
        SplitPath(this.excelFil.NavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFil.Navn := varFilNavn
        this.excelFil.Dir := varFilDir
        this.excelFil.NavnUdenExtension := varFilNavnUdenExtension
        
        
        return
    }

    vælgExcelFilMenu()
    {
        this.excelFil.NavnLong := FileSelect()
        SplitPath(this.excelFil.NavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFil.Navn := varFilNavn
        this.excelFil.Dir := varFilDir
        this.excelFil.NavnUdenExtension := varFilNavnUdenExtension
        
        
        return
    }

    aktiverExcelWorkbookReadonly()
    {

        this.excelObj := ComObject("Excel.Application")
        this.aktivWorkbook := this.excelObj.Workbooks.open(this.excelFil.NavnLong, , "ReadOnly" = true)

        return
    }

    vælgAktivWorkbookSheet(pSheetNummerEllerNavn)
    {
        this.aktivWorkbookSheet:= this.aktivWorkbook.Sheets(pSheetNummerEllerNavn)
    
        return
    }

    findBrugtExcelRangeIAktivWorkbookSheet()
    {

        this.aktivWorkbookSheetRækkerEnd := this.aktivWorkbookSheet.usedrange.rows.count
        this.aktivWorkbookSheetKolonnerEnd := this.aktivWorkbookSheet.usedrange.columns.count

        return
    }
    
    excelAktivRangetilArray()
    {
        this.aktivWorksheetArray := Array()
        this.aktivWorksheetArray := this.aktivWorkbookSheet.usedrange.value
        
        return
    }

    excelQuit()
    {
        this.excelObj.quit()
        return
    }

}

test := excelObj()

test.vælgExcelFilMenu()
test.aktiverExcelWorkbookReadonly()
test.vælgAktivWorkbookSheet(1)
test.findBrugtExcelRangeIAktivWorkbookSheet()
test.excelAktivRangetilArray()

msgbox test.excelFil.navn
MsgBox test.aktivWorksheetArray[1, 3]

test.excelQuit()

return
