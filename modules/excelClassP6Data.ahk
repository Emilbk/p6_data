/************************************************************************
 * @description Excel-class til brug ved P6-data-makro
 * @author 
 * @date 2024/10/18
 * @version 0.0.1
 * @extends excelclass.ahk
 ***********************************************************************/


#Include excelClass.ahk

/**
 * @parameter gyldigeKolonner Map,
 */
class excelObjP6Data extends excelIndlæsVlData {

    /** @type {Map} */
    gyldigeKolonner := Map(
        "Budnummer", 0,
        "Vognløbsnummer", 0,
        "Kørselsaftale", 0,
        "Styresystem", 0,
        "Startzone", 0,
        "Slutzone", 0,
        "Hjemzone", 0,
        "MobilnrChf", 0,
        "Vognløbskategori", 0,
        "Planskema", 0,
        "Økonomiskema", 0,
        "Statistikgruppe", 0,
        "Vognløbsnotering", 0,
        "Starttid", 0,
        "Sluttid", 0,
        "Sluttid", 0,
        "Undtagne transporttyper", 0,
        "Ugedage", 0
    )

    ugyldigeKolonner := Map()

    testKolonneNavnOgNummer := Map(
        "Budnummer", 1,
        "Vognløbsnummer", 2,
        "Kørselsaftale", 3,
        "Styresystem", 4,
        "Startzone", 5,
        "Slutzone", 6,
        "Hjemzone", 7,
        "MobilnrChf", 8,
        "Starttid", 9,
        "Sluttid", 10,
        "Ugedage", Array(11, 12, 13, 14, 15, 16, 17, 18),
        "Vognløbskategori", 19,
        "Planskema", 20,
        "Økonomiskema", 21,
        "Statistikgruppe", 22,
        "Vognløbsnotering", 23,
        "UndtagneTransporttyper", Array(24, 25, 26, 27)
    )

    vlResultatKolonneNavnOgNummer := Map(
        "Budnummer", 1,
        "Vognløbsnummer", 2,
        "Kørselsaftale", 3,
        "Vognløbsdato", 4,
        "Styresystem", 5,
        "Startzone", 6,
        "Slutzone", 7,
        "Hjemzone", 8,
        "MobilnrChf", 9,
        "Starttid", 10,
        "Sluttid", 11,
        "Vognløbskategori", 12,
        "Planskema", 13,
        "Økonomiskema", 14,
        "Statistikgruppe", 15,
        "Vognløbsnotering", 16,
        "UndtagneTransporttyper", 17
    )
    kolonneNavnogNummerArray := Array()

    ; testKolonneNavnOgNummerArray := this.kolonneNavnogNummerTilArray(this.testKolonneNavnOgNummer)


    ; ; Kolonnenavne opgivet i excel-ark, men ikke defineret i script
    ; p6DataTjekForGyldigeKolonner() {
    ;     for kolonneNavn in this.aktivWorksheetKolonneNavnOgNummer
    ;         if this.gyldigeKolonner.Has(kolonneNavn)
    ;             this.gyldigeKolonner[kolonneNavn] := 1
    ;         else
    ;             this.ugyldigeKolonner[kolonneNavn] := 0

    ;     for kolonnenavn, indhold in this.ugyldigeKolonner
    ;         if indhold = 0
    ;             MsgBox kolonneNavn " er ikke gyldig"
    ;     return
    ; }

    skrivExcelVognløbsResultat(pTjekketVognløbsdata) {


    }

    setWorkbookSavePath(pSavePath) {

        this.savePath := pSavePath
    }

    lavNyWorkbook() {

        if FileExist(this.savePath)
            throw Error("Workbook eksisterer allerede")

        xl := this.app

        xl.Workbooks.Add()
        workbook := xl.activeWorkbook
        activeSheet := workbook.activeSheet

        savePath := this.savePath
        workbook.Saveas(savepath)

        xl.quit()
    }

    setAktivWorkbookDir(pWorkbookDir) {

        this.aktivWorkbook := pWorkbookDir

    }

    setAktivWorksheetName(pAktivWorksheetName) {
        this.aktivWorksheetName := pAktivWorksheetName
    }

    setAktivRække(pRækkeNummer) {

        this.AktivRække := pRækkeNummer

    }

    skrivExcelKolonneNavn(pKolonneNavnOgNummerArray) {

        xl := this.app


        workbook := this.aktivWorkbook
        activeSheet := workbook.activeSheet
        ; activeSheet.name := this.aktivWorksheetName

        kolonneRække := 1
        for kolonne in pKolonneNavnOgNummerArray
        {

            aktivCelle := activeSheet.cells(kolonneRække, A_Index)
            aktivCelle.Value := StrTitle(kolonne)
            if kolonne = "Ugedage"
            {
                aktivCelle.addcomment()
                aktivCelle.comment.text("Udfyldte ugedage kan undlades")

            }

        }

        activeSheet.Columns().AutoFit


        ; workbook.Save()
        ; xl.quit()

    }

    skrivExcelVognløb(pVognløb, pVlResultatKolonneNavnOgNummer) {

        vl := pVognløb
        rækkeNummer := this.AktivRække

        workbook := this.aktivWorkbook
        activeSheet := workbook.activeSheet

        for vognløbsKolonne in pVlResultatKolonneNavnOgNummer
        {

            aktivCelle := activeSheet.cells(rækkeNummer, A_Index)
            if vognløbsKolonne = "UndtagneTransporttyper"
            {
                kolonneNummer := A_Index
                for kolonne in pVognløb.tilIndlæsning.UndtagneTransporttyper
                {
                    kolonneNavn := activeSheet.cells(1, kolonneNummer)
                    aktivCelle := activeSheet.Cells(rækkeNummer, kolonneNummer)
                    kolonneNavn.value := "UndtagneTransporttyper"
                    aktivCelle.Value := kolonne
                    kolonneNummer += 1
                }
            }
            else
                aktivCelle.Value := StrTitle(pVognløb.tilIndlæsning.%vognløbsKolonne%)

        }

    }
}