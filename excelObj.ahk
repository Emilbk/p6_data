#Requires AutoHotkey v2.0
#SingleInstance Force

/**
 * @param pExcelFil valgt excel-fil-path
 *   @property excelObj ComObject
 *   @property aktivWorkbook 
 *   @property aktivWorkbookSheet 
 *   @property aktivWorkbookSheetRækkerEnd 
 *   @property aktivWorkbookSheetKolonnerEnd 
 *   @property aktivWorksheetArray ComArray, [kolonne, række] - int.
 */
class excelObj extends Class {
    __New(pExcelFil) {
        this.excelObj := ComObject("Excel.Application")

        this.excelFilNavnLong := pExcelFil

        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension
    }

    ; Properties
    excelObj := ""
    aktivWorkbook := ""
    aktivWorkbookSheet := ""
    aktivWorkbookSheetRækkerEnd := ""
    aktivWorkbookSheetKolonnerEnd := ""
    aktivWorksheetArray := Array()

    /**
     * Aktiver excel-dokument
     * @property this.aktivWorkbook resultat
     */
    aktiverExcelWorkbookReadonly() {

        this.aktivWorkbook := this.excelObj.Workbooks.open(this.excelFilNavnLong, , "ReadOnly" = true)

        return
    }

    /**
     * Aktiver sheet i aktiv workbook
     * @param pSheetNummerEllerNavn det valgte ark, string eller int
     * @property aktivWorkbookSheet resultat
     */
    vælgAktivWorkbookSheet(pSheetNummerEllerNavn) {
        this.aktivWorkbookSheet := this.aktivWorkbook.Sheets(pSheetNummerEllerNavn)

        return
    }

    /**
     * Definer de fyldte celler i aktivt ark
     * @property this.aktivWorkbookSheetRækkerEnd Sidste række, int
     * @property this.aktivWorkbookSheetKolonnerEnd Sidste Kolonne, int
     */
    findBrugtExcelRangeIAktivWorkbookSheet() {

        this.aktivWorkbookSheetRækkerEnd := this.aktivWorkbookSheet.usedrange.rows.count
        this.aktivWorkbookSheetKolonnerEnd := this.aktivWorkbookSheet.usedrange.columns.count

        return
    }

    /**
     * Hent range af udfyldte celler til comObj-array
     * @property this.aktivWorksheetArray array, [kolonne-nr, række-nr]
     */
    excelAktivRangetilArray() {
        this.aktivWorksheetArray := this.aktivWorkbookSheet.usedrange.value

        return
    }

    quit() {
        this.excelObj.quit()
        return
    }

}
