#Include ../include.ahk


excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
excelArk := 1
excelobj := excelObjP6Data(excelFil, excelArk)
excelArray := excelobj.getVlArray()

vlConst := VognløbConstructor(excelArray, excelobj.getGyldigeKolonner())
vlContainer := vlConst.getBehandletVognløbsArray()

vl := vlContainer[1][1]
p6obj := P6()

; MsgBox vl.parametre.vognløbsnummer.forventetIndhold

p6Vindue := 331448 

p6obj.setP6Vindue(p6Vindue)


testExcelPath := A_ScriptDir "\exceltest\test2.xlsx" 

if !FileExist(testExcelPath)
    excelNyWorkbook := excelLavNyWorkbook(testExcelPath)

testexcel := excelBehandlWorkbook()
testexcel.app.Visible := 1
testexcel.åbenWorkbookReadWrite(testExcelPath)
testexcel.setAktivSheet(1)
testexcel.setGyldigeKolonner(excelobj.getGyldigeKolonner())
testexcel.gyldigKolonneNavnOgNummer.Vognløbsdato :={ kolonneNavn: "Vognløbsdato", kolonneNummer: 1, iBrug: 0 },
rækkeNummer := 1
for vlSamling in vlContainer
{
    vlFørste := vlSamling[1]
    p6Samling := p6()
    p6Samling.setP6Vindue(p6Vindue)
    p6Samling.setVognløb(vlFørste)
    ; p6Samling.funkKørselsaftaleÆndrHjemzone()

    p6Samling.navAktiverP6Vindue()
    p6Samling.navLukAlleVinduer()
    p6Samling.navVindueVognløb()
    for vl in vlSamling
    {
        rækkeNummer += 1
        p6obj := p6()
        p6obj.setP6Vindue(p6Vindue)
        p6obj.setVognløb(vl)
        try {
            ; p6obj.funkIndhentData()
            p6obj.funkIndhentData()


        } catch P6MsgboxError as msg {
            
            vl.setFejlLog(msg)
            testexcel.aktivSheet.Cells(rækkeNummer, 1).value := msg.Message
            testexcel.aktivSheet.Cells(rækkeNummer, 2).value := vl.parametre.Vognløbsnummer.forventetIndhold
        } 
        else
            {
                testexcel.udfyldVognløbRækker(vl, rækkeNummer)
                
            }
    }
}
; SendInput("{Enter}")
; p6obj.navAktiverP6Vindue()
; p6obj.navLukAlleVinduer()
; p6obj.navVindueKørselsaftale()
; p6obj.kørselsaftaleIndtastKørselsaftale()
; p6obj.kørselsaftaleTjekKørselsaftaleOgStyresystem()
; p6obj.kørselsaftaleÆndr()
; p6obj.kørselsaftaleIndhentPlanskema()
; p6obj.kørselsaftaleIndhentØkonomiskema()
; p6obj.kørselsaftaleIndhentStatistikgruppe()
; p6obj.kørselsaftaleIndhentNormalHjemzone()
; p6obj.kørselsaftaleIndhentKørerIkkeTransportTyper()
; p6obj.kørselsaftaleIndhentObligatoriskVognmand()
; p6obj.kørselsaftaleIndhentPauseRegel()
; p6obj.kørselsaftaleIndhentPauseDynamisk()
; p6obj.kørselsaftaleIndhentPauseStart()
; p6obj.kørselsaftaleIndhentPauseSlut()
; p6obj.kørselsaftaleIndhentVognmandNavn()
; p6obj.kørselsaftaleIndhentVognmandCO()
; p6obj.kørselsaftaleIndhentVognmandAdresse()
; p6obj.kørselsaftaleIndhentVognmandPostNr()
; p6obj.kørselsaftaleIndhentVognmandTelefon()
msgbox vlContainer[1][1].fejlLog.Extra
