#Include excelObj.ahk

class excelObjP6Data extends excelObj

{

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

    ; lav optional parameter til excel-fil
    indlæsfil()
    {
        ; håndter hvor?
        if !this.excel_fil_long
            throw Error("Ingen fil indlæst!")



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

}