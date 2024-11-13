#Include include.ahk
p6vindID := 336108

F12:: Pause
!F12:: konfigurering.setBreakLoop()
+F12:: ExitApp()
^F12::
{
    udrulÆndringerExcel()
    return
}
+F11:: {
    p6vindue := p6nav.setP6Vindue()
    MsgBox(p6vindue)
    return
}

p6nav := p6()
p6vindue := p6nav.setP6Vindue(p6vindID)
; MsgBox p6vindue


udrulÆndringerExcel() {
    ; excelobj := mockExcelP6Data()
    excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
    excelArk := 1
    excelobj := excelObjP6Data(excelFil, excelArk)
    excelArray := excelobj.getVlArray()
    ; excelobj.setExcelFil(excelfil)
    ; excelobj.helperIndlæsAlt(1)
    ; excelobj.quit()
    udrulÆndringer(excelArray)

    return
}

udrulÆndringermock() {
    excelobj := mockExcelP6Data()
    udrulÆndringer(excelobj)

    return
}

udrulÆndringer(pExcelobj)
{

    vlConstruct := VognløbConstructor(pExcelobj)
    vlArray := vlConstruct.getBehandletVognløbsArray()

    p6nav.navAktiverP6Vindue()
    p6nav.navLukAlleVinduer()
    tlf := 7011000
    samlingIgen := ""
    for vognløbssamling in vlArray
    {
        samlingsNummer := A_Index
        ; if samlingIgen
        ; samlingsNummer := samlingIgen, samlingIgen := ""
        for vognløb in vognløbssamling
        {
            tlf += 1
            vognløbsnummerISamling := A_Index
            p6Obj := p6()
            p6Obj.setP6Vindue(p6vindID)
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
            } catch P6MsgboxError as msgboxFejl {

                vognløb.fejlLog.importP6MsgboxFejl(msgboxFejl)
                SendInput("{enter}")
                ; MsgBox fejl.Message
            } catch P6Indtastningsfejl as indtastningsfejl {
                ; MsgBox("Datafejl!")
                ; indtastningsfejl.test()
                ; indtastningsfejl.construct(vognløb)
                p6Obj.vognløbsbilledeAfbryd()
            } catch p6ForkertDataError as datafejl {
                ; MsgBox("Datafejl!")
                ; indtastningsfejl.test()
                p6Obj.vognløbsbilledeAfbryd()
            }
            else {
            }
            ; hvornår i loopet?
        }


    }
    MsgBox "Done!"

}
konfigurering := config()

; excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
; excelArk := 1
; excelobj := excelObjP6Data()
; excelobj.get(excelFil, excelArk)

; ; excelMock := mockExcelP6Data()
; excelRækkeArray := excelobj.getRækkeData()
; vlConstruct := VognløbConstructor()
; vlContainer := vlConstruct.behandlVognløsbsdata(excelRækkeArray)
; vognlob := vlContainer[1][1]

; excelTest := excelObjP6Data()
; excelPath := A_ScriptDir "\exceltest\test-" FormatTime(, "dd-HH-mm-ss") ".xlsx" ; Saves in the same folder as the script
; excelTest.setWorkbookSavePath(excelPath)
; excelTest.setAktivWorkbookDir(excelPath)
; excelTest.lavNyWorkbook()
; ; excelTest.setAktivWorksheetName("Vognløbstest")

; kolonneArray := exceltest.kolonneNavnogNummerTilArray(excelObj.aktivWorksheetKolonneNavnOgNummer)
; excelTest.skrivExcelKolonneNavn(kolonneArray)


; test := p6Mock()
; test.setVognløb(vognlob)

; test.tjekkedeParametre.skabOgTestParameter("Vognløbsnummer", "31400", "31401")
; test.tjekkedeParametre.skabOgTestParameter("Styresystem", "47", "47")


; excelfil := "C:\Users\ebk\makro\p6_data\VL.xlsx"
; excelobj := excelObjP6Data()
; excelobj.setExcelFil(excelfil)
; excelobj.helperIndlæsAlt(1)
; excelobj.quit()
; test := P6()

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

; udrulÆndringerExcel()

; return