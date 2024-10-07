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
            for key, value in this.kolonner
                this.excel_data[row_index][key] := value
            this.excel_data[row_index]["Ugedage"] := [0, 0, 0, 0, 0, 0, 0]
            this.excel_data[row_index]["Undtagne transporttyper"] := []
            loop EndCol
            {
                col_index := A_Index
                nuvKolonne := usedrangeArr[1, col_index]
                nuvCelle := usedrangeArr[row_index, col_index]
                if Type(nuvCelle) = "Float"
                    nuvCelle := String(Floor(nuvCelle))
                if nuvKolonne = "Undtagne transporttyper"
                    this.excel_data[row_index][nuvKolonne].push(nuvCelle)
                else if nuvKolonne = "Ugedage"
                {
                    for index, ugedag in ["ma", "ti", "on", "to", "fr", "lø", "sø"]
                        if nuvCelle = ugedag
                            this.excel_data[row_index][nuvKolonne][index] := ugedag
                }
                else
                    this.excel_data[row_index][nuvKolonne] := nuvCelle
            }
        }
        excel.quit()
        return
    }

    indlæsfilFunk()
    {
        this.vælgfil()
        this.indlæsfil()
    }
}