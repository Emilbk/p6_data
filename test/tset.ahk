#Include ../include.ahk


excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
excelArk := 1
excelobj := excelObjP6Data(excelFil, excelArk)
excelArray := excelobj.getVlArray()

vlConst := VognløbConstructor(excelArray)
vlContainer := vlConst.getBehandletVognløbsArray()

vl := vlContainer[1][1]
p6obj := P6()

p6Vindue := 333930 

; p6obj.setP6Vindue(p6Vindue)
p6obj.setVognløb(vl)

; p6obj.navAktiverP6Vindue()
; p6obj.navLukAlleVinduer()
; p6obj.navVindueVognløb()
; p6obj.vognløbsbilledeIndtastVognløbOgDato()
; p6obj.vognløbsbilledeÆndrVognløb()
; p6obj.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
; p6obj.vognløbsbilledeÆndrVognløb()
; ; SendInput("{Enter}")
; p6obj.vognløbsbilledeIndhentÅbningstiderogZone()
; p6obj.vognløbsbilledeIndhentØvrige()
; p6obj.vognløbsbilledeIndhentTransporttyper()
; SendInput("{Enter}")
; p6obj.navAktiverP6Vindue()
; p6obj.navLukAlleVinduer()
; p6obj.navVindueKørselsaftale()
; p6obj.kørselsaftaleIndtastKørselsaftale()
; p6obj.kørselsaftaleTjekKørselsaftaleOgStyresystem()
; p6obj.kørselsaftaleÆndr()
p6obj.kørselsaftaleIndhentPlanskema()
p6obj.kørselsaftaleIndhentØkonomiskema()
p6obj.kørselsaftaleIndhentStatistikgruppe()
p6obj.kørselsaftaleIndhentNormalHjemzone()
p6obj.kørselsaftaleIndhentTransportyper()
p6obj.kørselsaftaleIndhentVognmand()
p6obj.kørselsaftaleIndhentObligatoriskVognmand()
p6obj.kørselsaftaleIndhentPauseRegel()
p6obj.kørselsaftaleIndhentPauseDynamisk()
p6obj.kørselsaftaleIndhentPauseStart()
p6obj.kørselsaftaleIndhentPauseSlut()
p6obj.kørselsaftaleIndhentVognmandNavn()
p6obj.kørselsaftaleIndhentVognmandCO()
p6obj.kørselsaftaleIndhentVognmandAdresse()
p6obj.kørselsaftaleIndhentVognmandPostNr()
p6obj.kørselsaftaleIndhentVognmandTelefon()
msgbox vlContainer[1][1].tilIndlæsning.vognløbsnummer