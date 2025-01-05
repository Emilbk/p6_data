class excelHentData {

    __New(pExcelFil, pArkNavnEllerNummer, excelApp := "") {
        if excelApp
            this.app := excelApp
        else this.app := ComObject("Excel.Application")
        this.fil := {}
        this.excel := {}
        this.fil.path := pExcelFil
        this.excel.arkNavn := pArkNavnEllerNummer

        this._setFilVariabler()
    }

    _setFilVariabler() {

        SplitPath(this.fil.path, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.fil.navn := varFilNavn
        this.fil.navnIngenExt := varFilNavnUdenExtension
        this.fil.dir := varFilDir
    }
    _åbenWorkbookReadonly() {
        this.excel.aktivWorkbook := this.app.Workbooks.open(this.fil.path, "ReadOnly" = true)
        this.excel.aktivWorksheet := this.excel.aktivWorkbook.Sheets(this.excel.arkNavn)
    }

    _indlæsAktivRangeTilArray() {
        this.excel.SafeArray := this.excel.aktivWorksheet.usedrange.value
    }

    _quit() {
        this.app.quit()
    }

    excelDataArray {
        get {
            this._åbenWorkbookReadonly()
            this._indlæsAktivRangeTilArray()

            outputArray := []
            safeArray := this.excel.SafeArray
            maxRækker := safeArray.MaxIndex(1)
            maxKolonner := safeArray.MaxIndex(2)

            loop maxRækker {
                rækkeIndex := A_Index
                outputArray.Push(Array())
                loop maxKolonner {
                    kolonneIndex := A_Index
                    aktivCelle := safeArray[rækkeIndex, kolonneIndex]
                    if  IsFloat(aktivCelle)
                        aktivCelle := String(Floor(aktivCelle))
                    outputArray[rækkeIndex].Push(aktivCelle)
                }
            }

            return outputArray
        }
    }

}