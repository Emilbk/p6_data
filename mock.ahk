ListLines 1

#Include modules
#Include vlClass.ahk
#Include p6.ahk
#Include excelClass.ahk
#Include excelClassP6Data.ahk
#Include config.ahk

F12:: Pause
!F12:: konfigurering.setBreakLoop()
+F12:: ExitApp()
::tudrul::
{
    udrulÆndringerMock()
    return
}

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
        this.kolonneNavnOgNummer := Map("Budnummer", 1, "Vognløbsnummer", 2)

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
            "UndtagneTransporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Vognløbsdato", "",
            "Ugedage", Array("ma", "ma", "ma")
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
            "UndtagneTransporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Vognløbsdato", "",
            "Ugedage", Array("ma")
        ))

        this.færdigbehandletData := { kolonneNavnOgNummer: this.kolonneNavnOgNummer, rækkerSomMapIArray: this.aktivWorksheetArrayRække }

    }

    getKolonneNavnOgNummer()
    {
        return this.færdigbehandletData.kolonneNavnOgNummer
    }

    getRækkeData()
    {
        return this.færdigbehandletData.rækkerSomMapIArray
    }
}
udrulÆndringerExcel(){
    excelobj := mockExcelP6Data()
    excelfil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
    excelobj := excelObjP6Data()
    excelobj.setExcelFil(excelfil)
    excelobj.helperIndlæsAlt(1)
    excelobj.quit()
    udrulÆndringer(excelobj)

    return
}

udrulÆndringermock(){
    excelobj := mockExcelP6Data()
    udrulÆndringer(excelobj)

    return
}

udrulÆndringer(pExcelobj)
{
    excelobj := pExcelobj
    rækkeArray := excelobj.getRækkeData()

    vlObj := VognløbConstructor()
    vlObj.setVognløbsdata(rækkeArray)
    vlArray := vlObj.getVognløbsdata()

    p6nav := p6()
    p6nav.navAktiverP6Vindue()
    p6nav.navLukAlleVinduer()
    tlf := 7011000
    for vognløbssamling in vlArray
    {
        samlingsNummer := A_Index
        for vognløb in vognløbssamling
        {
            tlf += 1
            p6Obj := p6()
            p6Obj.setVognløb(vognløb)
            if konfigurering.getBreakLoopStatus()
            {
                MsgBox "Break=loop"
                ; TODO save-state json
                konfigurering.removeBreakLoop()
                break 2
            }
            p6Obj.vognløb.tilIndlæsning.MobilnrChf := tlf
            try {
                p6Obj.funkÆndrVognløb()
            } catch Error as fejl {
                SendInput("{enter}")
                ; MsgBox fejl.Message
            }
        }
    }


    MsgBox "Done!"


}

konfigurering := config()


; excelfil := "C:\Users\ebk\makro\p6_data\VL.xlsx"
; excelobj := excelObjP6Data()
; excelobj.setExcelFil(excelfil)
; excelobj.helperIndlæsAlt(1)
; excelobj.quit()

; test := P6()
; loop 100
; {
;     test.kopierVærdi("ctrl", 1)
; }


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

udrulÆndringerExcel()

return