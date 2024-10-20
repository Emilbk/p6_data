#Include vlClass.ahk
#Include p6Navigering.ahk


class mockExcelP6Data extends Class {

    __New() {

        gyldigeKolonner := Map(
            "Budnummer", 1,
            "Vognløbsnummer", 1,
            "Kørselsaftale", 1,
            "Styresystem", 1,
            "Startzone", 1,
            "Slutzone", 1,
            "Hjemzone", 1,
            "MobilnrChf", 1,
            "Vognløbskategori", 1,
            "Planskema", 1,
            "Økonomiskema", 1,
            "Statistikgruppe", 1,
            "Vognløbsnotering", 1,
            "Starttid", 1,
            "Sluttid", 1,
            "Sluttid", 1,
            "Undtagne transporttyper", 1,
            "Ugedage", 1
        )

        this.aktivWorksheetArrayRække := Array()

        this.aktivWorksheetArrayRække.Push(Map(
            "Budnummer", "24-2267",
            "Vognløbsnummer", "31400",
            "Kørselsaftale", "3400",
            "Styresystem", "1",
            "Startzone", "ÅRH142",
            "Slutzone", "ÅRH142",
            "Hjemzone", "ÅRH142",
            "MobilnrChf", "701122010",
            "Vognløbskategori", "FG9",
            "Planskema", "31400",
            "Økonomiskema", "31400",
            "Statistikgruppe", "2GVEL",
            "Vognløbsnotering", "Notering1",
            "Starttid", "09",
            "Sluttid", "23",
            "Sluttid", "23",
            "Undtagne transporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Ugedage", Array("ma", "mA", "Ma")
        ))
        this.aktivWorksheetArrayRække.Push(Map(
            "Budnummer", "24-2266",
            "Vognløbsnummer", "31400",
            "Kørselsaftale", "3400",
            "Styresystem", "1",
            "Startzone", "ÅRH143",
            "Slutzone", "ÅRH143",
            "Hjemzone", "ÅRH143",
            "MobilnrChf", "701122011",
            "Vognløbskategori", "FG9",
            "Planskema", "31401",
            "Økonomiskema", "31401",
            "Statistikgruppe", "2GVEL",
            "Vognløbsnotering", "Notering2",
            "Starttid", "10",
            "Sluttid", "22",
            "Sluttid", "22",
            "Undtagne transporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Ugedage", Array("ma")
        ))

    }

    getData()
    {
        return this.aktivWorksheetArrayRække
    }

}

udrulÆndringerMock()
{

    mockExcel := mockExcelP6Data()

    vldata := mockExcel.getData()

    testvl := VognløbObj()
    testvlArray := Array()
    for vognløb in vldata
    {
        testvl.indhentVognløbsdata(vognløb)
        testvl.opretVognløbForHverDato()
        testp6 := P6()

        for vognløbsdag, vognløb in testvl.Vognløb
        {
            testp6.dataIndhentVlObj(vognløb)
            testp6.funkÆndrVognløb()
        }

    }
    MsgBox "Done!"


}

; for vl in vldata
; {
;     testp6.dataIndhentVlObj(vl)
;     MsgBox testp6.dataVognløb["Vognløbsnummer"]
; }
;     for key, value in testvl.vlData
;         MsgBox key ": " value
; ; testvl := Array()

; for vl in mockExcel.aktivWorksheetArrayRække
; {
;     testvl.push(konkretVognløb())
;     testvl[A_Index].indhentVognløbsdataTilOprettelse(vl)
;     testvl[A_Index].opretVognløbForHverDato()

; }


; for vl in testvl
; {
;     vl.P6 := P6()
;     vl.p6.dataIndhentVlObj(vl)
; }
; testvl := konkretVognløb()

; testvl.indhentVognløbsdataTilOprettelse(mockExcel.aktivWorksheetArrayRække[1])
; testvl.opretVognløbForHverDato()

; testvl.eksempelDatastruktur()


return