#Requires AutoHotkey v2.0
#SingleInstance Force

/**
 * @param pExcelFil valgt excel-fil-path
 * @property excelObj ComObject
 * @property aktivWorkbook 
 * @property aktivWorksheet 
 * @property aktivWorksheetRækkerEnd 
 * @property aktivWorksheetKolonnerEnd 
 * @property aktivWorksheetArrayAlt ComArray, [kolonne, række] - int.
 * @property aktivWorksheetKolonneNavnOgNummer Map, Kolonnenavn: kolonnerække
 * @property aktivWorksheetArrayRække array med map for hver række, kolonnenavn: kolonneindhold
 */
class excelObj extends Class {
    __New(pExcelFil := "") {
        this.excelObj := ComObject("Excel.Application")

        this.excelFilNavnLong := pExcelFil

        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        return
    }

    ; Properties
    excelObj := ""
    aktivWorkbook := ""
    aktivWorksheet := ""
    aktivWorksheetRækkerEnd := ""
    aktivWorksheetKolonnerEnd := ""
    aktivWorksheetArrayAlt := Array()
    aktivWorksheetKolonneNavnOgNummer := Map()

    /**
     * Vælg excel-fil hvis ikke indlæst gennem constructor
     * @param pExcelFil 
     */
    filVælgExcelFil(pExcelFil) {

        this.excelFilNavnLong := pExcelFil

        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        return
    }

    /**
     * Aktiver excel-dokument
     * @property this.aktivWorkbook resultat
     */
    filAktiverExcelWorkbookReadonly() {

        this.aktivWorkbook := this.excelObj.Workbooks.open(this.excelFilNavnLong, , "ReadOnly" = true)

        return
    }

    /**
     * Aktiver sheet i aktiv workbook
     * @param pSheetNummerEllerNavn det valgte ark, string eller int
     * @property aktivWorksheet resultat
     */
    filVælgAktivWorksheet(pSheetNummerEllerNavn) {
        this.aktivWorksheet := this.aktivWorkbook.Sheets(pSheetNummerEllerNavn)

        return
    }

    /**
     * Definer de fyldte celler i aktivt ark
     * @property this.aktivWorksheetRækkerEnd Sidste række, int
     * @property this.aktivWorksheetKolonnerEnd Sidste Kolonne, int
     */
    dataFindBrugtExcelRangeIAktivWorksheet() {

        this.aktivWorksheetRækkerEnd := this.aktivWorksheet.usedrange.rows.count
        this.aktivWorksheetKolonnerEnd := this.aktivWorksheet.usedrange.columns.count

        return
    }

    /**
     * Hent range af udfyldte celler til comObj-array
     * @property this.aktivWorksheetArrayAlt array, [kolonne-nr, række-nr]
     */
    dataIndlæsAktivRangetilArray() {
        this.aktivWorksheetArrayAlt := this.aktivWorksheet.usedrange.value
        ; this.aktivWorksheetArrayAlt := this.aktivWorksheet.usedrange.value

        return
    }

    /**
     * Opretter map med kolonnenavn og nummer, kolonnenanv[kolonnenummer]
     * Hvis flere kolonner med samme navn oprettes der et array med kolonnenummer[i...i]
     * @property this.aktivWorksheetKolonneNavnOgNummer
     */
    dataIndlæsKolonneNavnogNummerTilMap() {
        if not this.aktivWorksheetArrayAlt is ComObjArray
            throw Error("aktivWorksheetArrayAlt er ikke indlæst")

        loop this.aktivWorksheetRækkerEnd {
            rækkeNummer := A_Index
            rækkeKolonneNavn := 1
            if rækkeNummer != rækkeKolonneNavn
                break
            loop this.aktivWorksheetKolonnerEnd {
                kolonneNummer := A_Index
                nuværendeKolonneNavn := this.aktivWorksheetArrayAlt[rækkeNummer, kolonneNummer]
                if Type(nuværendeKolonneNavn) = "Float"
                    nuværendeKolonneNavn := String(Floor(nuværendeKolonneNavn))
                if (this.aktivWorksheetKolonneNavnOgNummer.Has(nuværendeKolonneNavn)) {
                    if (type(this.aktivWorksheetKolonneNavnOgNummer[nuværendeKolonneNavn]) != "Array")
                        this.aktivWorksheetKolonneNavnOgNummer[nuværendeKolonneNavn] := Array(this.aktivWorksheetKolonneNavnOgNummer[
                            nuværendeKolonneNavn])
                    this.aktivWorksheetKolonneNavnOgNummer[nuværendeKolonneNavn].Push(kolonneNummer)
                }
                else
                    this.aktivWorksheetKolonneNavnOgNummer[nuværendeKolonneNavn] := kolonneNummer
            }
        }

        return
    }

    /**
     * Indlæser rækker til maps i array, en række pr. map, organiseret i samlet array
     * Hvis flere kolonner med samme navn oprettes der array med cellerne, i det underordnede map
     * @property this.WorksheetArrayRække indlæste rækker, første række fjernes (kolonne-overskrifter)
     */
    dataIndlæsRækkeArrayMinusKolonneNavne() {

        this.aktivWorksheetArrayRække := Array()
        loop this.aktivWorksheetRækkerEnd {
            rækkenummer := A_Index
            kolonneNavnRække := 1
            this.aktivWorksheetArrayRække.Push(Map())
            loop this.aktivWorksheetKolonnerEnd {
                kolonneNummer := A_Index
                kolonneNavn := this.aktivWorksheetArrayAlt[kolonneNavnRække, kolonneNummer]
                celleIndhold := this.aktivWorksheetArrayAlt[rækkenummer, kolonneNummer]
                if Type(celleIndhold) = "Float"
                    celleIndhold := String(Floor(celleIndhold))
                if (this.aktivWorksheetArrayRække[rækkenummer].Has(kolonneNavn)) {
                    if (type(this.aktivWorksheetArrayRække[rækkenummer][kolonneNavn]) != "Array")
                        this.aktivWorksheetArrayRække[rækkenummer][kolonneNavn] := Array(this.aktivWorksheetArrayRække[
                            rækkenummer][kolonneNavn])
                    this.aktivWorksheetArrayRække[rækkenummer][kolonneNavn].push(celleIndhold)
                }
                else
                    this.aktivWorksheetArrayRække[rækkenummer][kolonneNavn] := celleIndhold
            }
        }
        this.aktivWorksheetArrayRække.RemoveAt(1)
        return
    }
    /**
     * 
     * @param pArkNavnEllerNummer 
     */
    helperIndlæsAlt(pArkNavnEllerNummer) {
        this.filAktiverExcelWorkbookReadonly()
        this.filVælgAktivWorksheet(pArkNavnEllerNummer)
        this.dataFindBrugtExcelRangeIAktivWorksheet()
        this.dataIndlæsAktivRangetilArray()
        this.dataIndlæsKolonneNavnOgNummerTilMap()
        this.dataIndlæsRækkeArrayMinusKolonneNavne()
        return
    }

    quit() {
        this.excelObj.quit()
        return
    }

}

excelpath := "C:\Users\ebk\makro\p6_data\VL.xlsx"
test := excelObj()
test.filVælgExcelFil(excelpath)

test.helperIndlæsAlt(2)
test.quit()
return