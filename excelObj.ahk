#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; var

; MsgBox test[1, 2].value
class excelObj extends Class
{

    excel_fil_long := ""
    excel_fil_tekst := ""
    excel_data := []

    ; mulige kolonneindlæsninger
    kolonner := Map(
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
        "Vognløbskategori", 0,
        "Undtagne transporttyper", [],
        "Ugedage", [0, 0, 0, 0, 0, 0, 0]
    )

    vælgfil()
    {
        valgtExcelFilLong := FileSelect()
        if !valgtExcelFilLong
            return
        SplitPath(valgtExcelFilLong, &valgtExcelFil)
        this.excel_fil_long := valgtExcelFilLong

        this.excel_fil_tekst := "Indlæst excel-fil: " . valgtExcelFil

        return
    }

    ; lav optional parameter til excel-fil
    indlæsfil()
    {

        ; undtagneTrKolStart := 13
        ; undtagneTrKolSlut := 22
        if !this.excel_fil_long
            throw Error("Ingen fil indlæst!")

        excel := ComObject("Excel.Application")
        excel.Visible := 0
        excel_fil := this.excel_fil_long
        workbook := excel.Workbooks.open(excel_fil, , "ReadOnly" = true)
        workbook_sheet := workbook.Sheets(1)


        EndRow := workbook_sheet.usedrange.rows.count
        EndCol := workbook_sheet.usedrange.columns.count
        usedrangeArr := workbook_sheet.usedrange.value


        loop EndRow
        {
            row_index := A_Index
            this.excel_data.Push(Map())
            this.excel_data[row_index] := this.kolonner.Clone()

            loop EndCol
            {
                col_index := A_Index
                nuvKolonne := usedrangeArr[1, col_index]
                nuvCelle := usedrangeArr[row_index, col_index]
                if Type(nuvCelle) = "Float"
                    nuvCelle := String(Floor(nuvCelle))
                if nuvKolonne = "Ugedage" or nuvKolonne = "Undtagne transporttyper"
                {
                    if nuvKolonne = "Ugedage"
                    {
                        if nuvCelle = "ma"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "ti"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "on"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "to"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "fr"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "lø"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                        if nuvCelle = "sø"
                            this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                    }
                    if nuvKolonne = "Undtagne transporttyper"
                        this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                }
                else
                    this.excel_data[row_index][nuvKolonne] := nuvCelle


                ; if nuvKolonne := "Ugedage"
                ; {
                ;     if nuvCelle = "ma"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "ti"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "on"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "to"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "fr"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "lø"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ;     if nuvCelle = "sø"
                ;         this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                ; }
                ; if nuvKolonne := "Undtagne transporttyper"
                ;     this.excel_data[row_index][nuvKolonne].push(nuvCelle)

                ; this.excel_data[row_index][nuvKolonne] := nuvCelle
            }

        }


        excel.quit()
        return
    }

    ; hentKolonneNummer()
    ; {
    ;     for kolonneNummerExcel, kolonneNavnExcel in this.excel_data[1]
    ;         for kolonneNavnIntern, kolonneNummerIntern in this.kolonne_nummer
    ;         {
    ;             if kolonneNavnExcel = "Ugedage"
    ;             {
    ;                 this.kolonne_nummer["Ugedage"].push(kolonneNummerExcel)
    ;                 break
    ;             }
    ;             if kolonneNavnExcel = "Undtagne transporttyper"
    ;             {
    ;                 this.kolonne_nummer["Undtagne transporttyper"].push(kolonneNummerExcel)
    ;                 break
    ;             }
    ;             if kolonneNavnExcel = kolonneNavnIntern
    ;             {
    ;                 this.kolonne_nummer[kolonneNavnIntern] := kolonneNummerExcel
    ;                 break
    ;             }
    ;         }
    ;     return
    ; }

    indlæsfilFunk()
    {
        this.vælgfil()
        this.indlæsfil()
        ; this.hentKolonneNummer()
    }
}

; test := excelObj()
; test.indlæsfilFunk()

; MsgBox test.excel_fil_long
; MsgBox test.kolonne_nummer["Vognløbsnummer"]
; MsgBox test.kolonne_nummer["Undtagne transporttyper"][3]
;     MsgBox "Data indlæst!"
;     ; DataGUI.opt("-Disabled")
;     ; WinActivate(DataguiNavn)
; }
