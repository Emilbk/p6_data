#Include ../include.ahk


excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
excelArk := 1
excelobj := excelObjP6Data(excelFil, excelArk)
excelArray := excelobj.getVlArray()

vlConst := VognløbConstructor(excelArray)
vlContainer := vlConst.getBehandletVognløbsArray()

vl := vlContainer[1][1]
p6obj := P6()

; MsgBox vl.parametre.vognløbsnummer.forventetIndhold

p6Vindue := 1706740

p6obj.setP6Vindue(p6Vindue)
; p6obj.setVognløb(vl)

; p6obj.funkIndhentVognløbsbillede()
; p6obj.funkIndhentKørselsaftale()
for vlSamling in vlContainer
{
    vlFørste := vlSamling[1]
    p6Samling := p6()
    p6Samling.setP6Vindue(p6Vindue)
    p6Samling.setVognløb(vlFørste)
    p6Samling.funkKørselsaftaleÆndrHjemzone()

    p6Samling.navLukAlleVinduer()
    p6Samling.navVindueVognløb()
    for vl in vlSamling
    {
        p6obj := p6()
        p6obj.setP6Vindue(p6Vindue)
        p6obj.setVognløb(vl)
        try {
            p6obj.funkVognløbsbilledeÆndrHjemzone()

        } catch P6MsgboxError as msg {

            vl.setFejlLog(msg)
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
