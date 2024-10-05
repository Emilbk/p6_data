#Requires AutoHotkey v2.0

class vlObj extends Class
{
    vl_data := Map(
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

    ; fjern kobling til datagui
    IndhentData(p_excel_array)
    {
        this.vl_data := p_excel_array
        return
    }
}

    ; DataGUI.totalExcelRække := p_data_array.Length - 1
    ; DataGUI.nuværendeExcelRække := p_række_nummer - 1
    ; DataGUI.excelRækkeTekst := "Excelrække " DataGUI.nuværendeExcelRække "/" DataGUI.totalExcelRække
    ; overskriftExcelRækker.Text := DataGUI.excelRækkeTekst

    ; PlanskemaEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonnePlanSkema]
    ; økonomiskemaEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonneØkonomiSkema]
    ; vognløbskategoriEditboxForventet.text := p_data_array[p_række_nummer][DataGUI.kolonneVognløbsKategori]

    ; overskriftVognløb.text := "Vognløb " this.vlVognløbsNummer ", " this.vlKørselsAftale "_" this.vlStyreSystem}
; ???
; p6_indhent_data()
; {

; }
; p6IndlæsData()
; {
;     MsgBox this.vlBudnummer " - " this.vlVognløbsNummer
;     MsgBox this.vlVognløbsKategori

; }
