#Include ../include.ahk

#hotif WinActive("P6-makro")
Escape:: guiLuk()
#Hotif
+Escape:: ExitApp()
F12:: setPauseStatus()


goo := Gui()
goo.Title := "P6-makro"
goo.tekst := {}
goo.knap := {}

goo.logDir := A_ScriptDir "\log"
goo.indskrivningFilpath := goo.logDir "\logIndskrivning" FormatTime(, "ddMM-HHmmss") ".txt"
goo.tjekFilPath := goo.logDir "\logTjek" FormatTime(, "ddMM-HHmmss") ".txt"
goo.fejlLogPath := goo.logDir "\logFejl" FormatTime(, "ddMM-HHmmss") ".txt"

goo.pauseStatus := 0

; Objs
goo.excel := {}
goo.excel.indlæst := 0
goo.excel.gemtdata := 0
goo.excel.obj := Object()
goo.vognløb := {}

goo.vlTider := Array()
; goo.vognløb.constr := VognløbConstructor()


goo.p6 := {}
goo.p6.obj := p6()
goo.p6.vinduehandle := ""


gooMenuBar := MenuBar()

filMenu := Menu()
omMenu := Menu()
dataMenu := Menu()
excelMenu := Menu()
p6DatabehandlingMenu := Menu()
dataFilMenu := Menu()
ændrDataMenu := Menu()
tjekDataMenu := Menu()
p6Menu := Menu()

omMenu.Add("Tjek for nyeste version", (*) => tjekVersion())
omMenu.Add()
omMenu.Add("Hjælp", (*) => visHjælp())
; filMenu.Add "&Open`tCtrl+O", (*) => FileSelect()  ; See remarks below about Ctrl+O.
filMenu.Add
filMenu.Add "E&xit", (*) => ExitApp()

dataMenu.Add("&Excel", excelMenu)
dataMenu.Add "&P6-handilinger", p6DatabehandlingMenu
; dataMenu.add()
; dataMenu.Add("&Datafil", dataFilMenu)

excelMenu.Add("&Indlæs Excel-fil`tCtrl+E", (*) => indlæsExcel())
excelMenu.Add("&Dan Excel-skabelon", (*) => danExcelSkabelon())

p6DatabehandlingMenu.Add("Ændr data", ændrDataMenu)
p6DatabehandlingMenu.Add("Indhent data", tjekDataMenu)
p6DatabehandlingMenu.Add("Fejlcheck data", (*) =>)


ændrDataMenu.Add("Ændr hjemzone", (*) => ændrVognløbHjemzone())
ændrDataMenu.Add("Ændr hjemzone (kun vognløb)", (*) => ændrVognløbHjemzoneKunVognløbsbillede())
ændrDataMenu.Add("Ændr vognløb", (*) => ændrVognløbAlt())
ændrDataMenu.Add("Indlæg vognløb", (*) => indlægVognløb())


tjekDataMenu.Add("Tjek hjemzone", (*) => indhentOgTjekVognløbHjemzone())
tjekDataMenu.Add("bladr vognløb", (*) => bladrVognløb())

dataFilMenu.Add("Indlæs data fra fil", (*) => hentVognløbsdata())
dataFilMenu.Add("Gem data til fil", (*) => gemVognløbsdata())


p6Menu.Add("Vælg P6-Vindue`tctrl+P", (*) => vælgP6Vindue())

gooMenuBar.Add("&Fil", filMenu)
gooMenuBar.Add("&Data", dataMenu)
gooMenuBar.Add("&P6", p6Menu)
gooMenuBar.Add("&Om", omMenu, "Right")

goo.MenuBar := gooMenuBar


excelfilTekst := "Ingen fil                               "
goo.tekst.valgtP6Vindue := goo.Add("Text", , "Aktivt P6-vindue: ")
goo.knap.valgtP6VindueKnap := goo.Add("Button", "XP+80 YP-5", "Aktiver valgt")
goo.tekst.indlæstExcelFil := goo.Add("Text", "XM", "Indlæst data-fil: " excelfilTekst)
goo.tekst.PauseTekst := goo.Add("Text", "XM W100", "")
; goo.tekst.indlæstExcelRækkerTekst := goo.Add("Text", , "Antal Rækker: ")
goo.knap.valgtP6VindueKnap.OnEvent("Click", (*) => goo.p6.obj.navAktiverP6Vindue())

goo.tekst.indlæstExcelFil.text := "Indlæst data-fil: " excelfilTekst


goo.OnEvent("Close", guiLuk)
goo.Show("W300 H200")

guiLuk(*) {

    msgsvar := MsgBox("Vil du lukke vinduet?", "Exit?", "0x21")
    if msgsvar = "OK"
        ExitApp()
}

vælgP6Vindue() {

    goo.p6.vinduehandle := ""

    loop {
        goo.p6.vindueHandle := WinActive("PLANET version 6")
    } until goo.p6.vinduehandle

    goo.knap.valgtP6VindueKnap.text := "Aktiver Valgt"
    sleep 100
    MsgBox("P6-vindue valgt:`n" WinGetTitle(), "P6-vindue", "iconi")

    goo.p6.obj.setP6Vindue(goo.p6.vindueHandle)
    ; goo.p6.obj.setP6Vindue()
}


indlæsExcel() {

    goo.excel.valgtExcelFil := FileSelect()
    if !goo.excel.valgtExcelFil
        return

    SplitPath(goo.excel.valgtExcelFil, &excelFil, &excelDir, &excelExt, &excelIngenExt)

    goo.excel.valgtExcelFilKort := excelFil
    goo.tekst.indlæstExcelFil.text := "Indlæser Excel-fil. Vent venligst."
    ; goo.Hide()

    goo.excel.obj := excelIndlæsVlData(goo.excel.valgtExcelFil, 1)
    goo.excel.vlArray := goo.excel.obj.getVlArray()
    goo.excel.gyldigeKolonner := goo.excel.obj.getGyldigeKolonner()

    goo.vognløb.constr := VognløbConstructor(goo.excel.vlArray, goo.excel.gyldigeKolonner)
    goo.vognløb.vlArray := goo.vognløb.constr.getBehandletVognløbsArray()

    goo.tekst.indlæstExcelFil.text := "Indlæst data-fil: " excelFil
    goo.excel.indlæst := 1

    MsgBox "Indlæst!", "Excel"
    return
}

danExcelSkabelon() {

    excelPath := A_ScriptDir "\excelSkabelon.xlsx"

    testExcel := udfyldTestExcelArk()
    testExcel.lavExcelTemplate(excelPath)


    MsgBox("Excelskabelon gemt som " excelpath, "Excel", "iconi")
}

indlægVognløb() {
    svar := MsgBox("Ændrer data på vognløb", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    tjekEksisterendeVindueHandle()
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    try {
        ; p6obj.funkKørselsaftaleÆndrHjemzone()

    } catch Error as e {
        ; kørselsaftalefejl
    }

    FileAppend(format("Makro startet {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        loopStartTid := A_Now
        vlSamlindIndex := A_Index
        FileAppend("-----------`n", goo.indskrivningFilpath)
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                tjekPauseStatus()
                vlStartTid := A_Now
                vl.tjekForbudtVognløbsDato()
                p6obj.funkIndlægVognløb()


            } catch P6MsgboxError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch p6ForkertDataError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch P6Indtastningsfejl as msg {

                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl i indtastning: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            else {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} OK`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold), goo.indskrivningFilpath)
                vlSlutTid := A_Now

            }
            ; FileAppend(format("Vognløb fuldført {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
        }

    }
    ; fix korrekt tidmåling
    loopSlutTid := A_Now
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidTime := Floor(slutTidDifferenceSec / 60 / 60)
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    FileAppend("-----------`n", goo.indskrivningFilpath)
    FileAppend(format("Makro færdig {1}`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    ; FileAppend(Format("Færdig efter {} {1} min, {2} sek", slutTidMin, slutTidSec), goo.indskrivningFilpath)
    MsgBox("Excelark færdigbehandlet.", "Vognløb ændret!", "Iconi")


}
ændrVognløbAlt() {

    svar := MsgBox("Ændrer data på vognløb", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    tjekEksisterendeVindueHandle()
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    try {
        ; p6obj.funkKørselsaftaleÆndrHjemzone()

    } catch Error as e {
        ; kørselsaftalefejl
    }

    FileAppend(format("Makro startet {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        loopStartTid := A_Now
        vlSamlindIndex := A_Index
        FileAppend("-----------`n", goo.indskrivningFilpath)
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                tjekPauseStatus()
                vlStartTid := A_Now
                vl.tjekForbudtVognløbsDato()
                p6obj.funkÆndrVognløb()


            } catch P6MsgboxError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch p6ForkertDataError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch P6Indtastningsfejl as msg {

                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl i indtastning: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            else {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} OK`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold), goo.indskrivningFilpath)
                vlSlutTid := A_Now

            }
            ; FileAppend(format("Vognløb fuldført {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
        }

    }
    ; fix korrekt tidmåling
    loopSlutTid := A_Now
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidTime := Floor(slutTidDifferenceSec / 60 / 60)
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    FileAppend("-----------`n", goo.indskrivningFilpath)
    FileAppend(format("Makro færdig {1}`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    ; FileAppend(Format("Færdig efter {} {1} min, {2} sek", slutTidMin, slutTidSec), goo.indskrivningFilpath)
    MsgBox("Excelark færdigbehandlet.", "Vognløb ændret!", "Iconi")


}

ændrVognløbHjemzoneKunVognløbsbillede() {

    svar := MsgBox("Ændrer hjemzone på vognløb og kørselsaftale", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    tjekEksisterendeVindueHandle()
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    try {
        ; p6obj.funkKørselsaftaleÆndrHjemzone()

    } catch Error as e {
        ; kørselsaftalefejl
    }

    FileAppend(format("Makro startet {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        loopStartTid := A_Now
        vlSamlindIndex := A_Index
        FileAppend("-----------`n", goo.indskrivningFilpath)
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                tjekPauseStatus()
                vlStartTid := A_Now
                vl.tjekForbudtVognløbsDato()
                p6obj.funkVognløbsbilledeÆndrHjemzone()


            } catch P6MsgboxError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch p6ForkertDataError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            catch P6Indtastningsfejl as msg {

                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl i indtastning: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.fejlLogPath)
                vlSlutTid := A_Now
                continue
            }
            else {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} OK`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold), goo.indskrivningFilpath)
                vlSlutTid := A_Now

            }
            ; FileAppend(format("Vognløb fuldført {1}.`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
        }

    }
    ; fix korrekt tidmåling
    loopSlutTid := A_Now
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidTime := Floor(slutTidDifferenceSec / 60 / 60)
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    FileAppend("-----------`n", goo.indskrivningFilpath)
    FileAppend(format("Makro færdig {1}`n", FormatTime(, "HH:mm:ss")), goo.indskrivningFilpath)
    ; FileAppend(Format("Færdig efter {} {1} min, {2} sek", slutTidMin, slutTidSec), goo.indskrivningFilpath)
    MsgBox("Excelark færdigbehandlet.", "Vognløb ændret!", "Iconi")

    return
}
ændrVognløbHjemzone() {

    svar := MsgBox("Ændrer hjemzone på vognløb og kørselsaftale", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    if p6Obj.vindueHandle = ""
    {
        MsgBox("P6-vindue er ikke valgt endnu!")
        return
    }
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    try {
        ; p6obj.funkKørselsaftaleÆndrHjemzone()

    } catch Error as e {
        ; kørselsaftalefejl
    }

    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        loopStartTid := A_Now
        FileAppend("-----------`n", goo.indskrivningFilpath)
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                vlStartTid := A_Now
                vl.tjekForbudtVognløbsDato()
                p6obj.funkVognløbsbilledeÆndrHjemzone()


            } catch P6MsgboxError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                vlSlutTid := A_Now
                continue
            }
            catch p6ForkertDataError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} Fejl: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.indskrivningFilpath)
                vlSlutTid := A_Now
                continue
            }
            else {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2} OK`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold), goo.indskrivningFilpath)
                vlSlutTid := A_Now

            }
        }

    }
    loopSlutTid := A_Now
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    FileAppend("-----------`n", goo.indskrivningFilpath)
    FileAppend(Format("Færdig efter {1} min, {2} sek", slutTidMin, slutTidSec), goo.indskrivningFilpath)
    MsgBox("Excelark færdigbehandlet.", "Vognløb ændret!", "Iconi")

    return
}
bladrVognløb() {
    svar := MsgBox("Indhenter hjemzone på vognløb og kørselsaftale", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb
    tjekEksisterendeVindueHandle()

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    p6Obj.navVindueVognløb()

    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        for Vl in vlSamling
        {
            try {
            tjekPauseStatus()
            p6Obj.setVognløb(vl)
            p6Obj.vognløbsbilledeIndtastVognløbOgDato()
            KeyWait("Esc", "D")
            } catch Error as e {
                
            }
        }
    }
}
indhentOgTjekVognløbHjemzone() {

    svar := MsgBox("Indhenter hjemzone på vognløb og kørselsaftale", , "1")
    if svar != "OK"
        return


    p6Obj := goo.p6.obj
    vlKørselaftale := goo.vognløb.vlArray.masterVognløb
    tjekEksisterendeVindueHandle()

    p6obj.setVognløb(vlKørselaftale)
    p6obj.navAktiverP6Vindue()
    p6obj.navLukAlleVinduer()
    try {
        ; p6obj.funkKørselsaftaleÆndrHjemzone()

    } catch Error as e {
        ; kørselsaftalefejl
    }

    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        loopStartTid := A_Now
        FileAppend("-----------`n", goo.tjekFilPath)
        for Vl in vlSamling
        {
            tjekPauseStatus()
            p6Obj.setVognløb(vl)
            try {
                vl.tjekForbudtVognløbsDato()
                p6obj.funkVognløbsbilledeIndhentHjemzone()

                vl.parametre.tjekParameterForFejl("Statistikgruppe")
                vl.parametre.tjekParameterForFejl("Startzone")
                vl.parametre.tjekParameterForFejl("Slutzone")
                vl.parametre.tjekParameterForFejl("Hjemzone")


            } catch P6MsgboxError as msg {
                ; if FileExist(filPath)
                vlSlutTid := A_Now
                FileAppend(Format("Vognløb {1} - {2}. Fejl i indhentning af data: {3}`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.tjekFilPath)
                FileAppend(msg.Message "`n", goo.tjekFilPath)
                continue
            }
            catch p6ForkertDataError as msg {
                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2}. Fejl i indhentning af data: {}3`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.tjekFilPath)
                vlSlutTid := A_Now
                continue
            }
            catch P6Indtastningsfejl as msg {

                ; if FileExist(filPath)
                FileAppend(Format("Vognløb {1} - {2}. Fejl i indhentning af data: {}3`n", vl.parametre.vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, msg.Message), goo.tjekFilPath)
                vlSlutTid := A_Now
                continue
            }
            else {
                ; if FileExist(filPath)
                vlSlutTid := A_Now
                FileAppend(vl.parametre.Vognløbsnummer.forventetIndhold " - " vl.parametre.vognløbsdato.forventetIndhold ": Indhentet OK`n", goo.tjekFilPath)

                if (vl.parametre.statistikGruppe.fejl = 1)
                    FileAppend(Format("Vognløb {1} - {2} Fejl i parameter. Forventet: Startzone: {3}, Slutzone: {4}, Hjemzone: {5}, Statistikgruppe: {6}. Fundet: Startzone: {7}, Slutzone {8}, Hjemzone: {9}, Statistikgruppe: {10}`n", vl.parametre.Vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, vl.parametre.startzone.forventetIndhold, vl.parametre.slutzone.forventetIndhold, vl.parametre.slutzone.forventetIndhold, vl.parametre.statistikGruppe.forventetIndhold, vl.parametre.startzone.eksisterendeIndhold, vl.parametre.slutzone.eksisterendeIndhold, vl.parametre.slutzone.eksisterendeIndhold, vl.parametre.statistikGruppe.eksisterendeIndhold,), goo.tjekFilPath)
                else
                    FileAppend(Format("Vognløb {1} - {2} OK. Forventet: Startzone: {3}, Slutzone: {4}, Hjemzone: {5}, Statistikgruppe: {6}. Fundet: Startzone: {7}, Slutzone {8}, Hjemzone: {9}, Statistikgruppe: {10}`n", vl.parametre.Vognløbsnummer.forventetIndhold, vl.parametre.vognløbsdato.forventetIndhold, vl.parametre.startzone.forventetIndhold, vl.parametre.slutzone.forventetIndhold, vl.parametre.slutzone.forventetIndhold, vl.parametre.statistikGruppe.forventetIndhold, vl.parametre.startzone.eksisterendeIndhold, vl.parametre.slutzone.eksisterendeIndhold, vl.parametre.slutzone.eksisterendeIndhold, vl.parametre.statistikGruppe.eksisterendeIndhold,), goo.tjekFilPath)
            }
        }

    }
    loopSlutTid := A_Now
    slutTidDifferenceSec := DateDiff(loopSlutTid, loopStartTid, "Seconds")
    slutTidMin := Floor(slutTidDifferenceSec / 60)
    slutTidSec := Mod(slutTidDifferenceSec, 60)
    FileAppend(Format("Færdig efter {1} min, {2} sek", slutTidMin, slutTidSec), goo.tjekFilPath)
    MsgBox("Excelark færdigbehandlet.", "Vognløb ændret!", "Iconi")

    return

}

tjekEksisterendeVindueHandle() {

    if goo.p6.obj.vindueHandle = ""
    {
        MsgBox("P6-vindue er ikke valgt endnu!")
        return
    }
}

tjekPauseStatus() {
    if goo.PauseStatus
        Pause
}

setPauseStatus() {
    if goo.PauseStatus
    {
        goo.pauseStatus := 0
        Pause 0
        goo.tekst.PauseTekst.text := ""
        if goo.p6.obj.vindueHandle
        {
            goo.P6.obj.navAktiverP6Vindue()
            sleep 200
        }
        ; FileAppend(Format("Sat på pause {1}", FormatTime(,"HH:mm:ss"), goo.indskrivningFilpath))
    }
    else
    {
        goo.pauseStatus := 1
        goo.tekst.PauseTekst.text := "På pause!"
        ; FileAppend(Format("Genoptaget {1}", FormatTime(,"HH:mm:ss"), goo.indskrivningFilpath))
    }
}

tjekOmOpdatering() {

    hentJaNej := 0

    HTTP := ComObject("WinHttp.WinHttpRequest.5.1")
    endPoint := "https://api.github.com/repos/emilbk/p6_data/releases/latest"

    http.open("GET", endPoint)
    http.Send()

    result := JSON.Load(http.ResponseText)

    gitV := result["tag_name"]
    localV := programVersion

    testVer := VerCompare(localV, gitV)

    if testVer < 0
        hentJaNej := MsgBox(Format("Ny version: {1}`n`n{2}`n`nHent nyeste version?", gitv, result["body"]), "Ny version tilgængelig", 0x1)

    if hentJaNej = "OK"
        Run result["zipball_url"]

}
tjekVersion() {

    hentJaNej := 0

    HTTP := ComObject("WinHttp.WinHttpRequest.5.1")
    endPoint := "https://api.github.com/repos/emilbk/p6_data/releases/latest"

    http.open("GET", endPoint)
    http.Send()

    result := JSON.Load(http.ResponseText)

    gitV := result["tag_name"]
    localV := programVersion

    testVer := VerCompare(localV, gitV)

    if testVer = 0
        MsgBox("Nyeste version er installeret!", "Versionstjek", "iconi")
    if testVer < 0
        hentJaNej := MsgBox(Format("Ny version: {1}`n`n{2}`n`nHent nyeste version?", gitv, result["body"]), "Ny version tilgængelig", 0x1)
    if testVer > 0
        MsgBox "?"

    if hentJaNej = "OK"
        Run result["zipball_url"]

}

visHjælp() {

    hjælpStr := "
    (
    Genveje:
    Ctrl+Escape: Stop og luk makro
    F12: Pause/Sæt igang
    )"

    MsgBox(hjælpStr, "Hjælp", "iconi")

}

gemVognløbsdata() {

    ; if !goo.excel.indlæst
    ; {
    ;     MsgBox("Er ikke indlæst")
    ;     return
    ; }

    valgtDatafil := ""

    valgtDatafil := FileSelect("S 24", "vognløbsdata", "titel", "json")
    if !valgtDatafil
        return
    jsonData := {}
    jsonData.excelGyldigeKolonner := goo.excel.gyldigeKolonner
    jsonData.vognløbsArray := goo.vognløb.vlArray
    jsonOutput := JSON.Dump(jsonData)
    if FileExist(valgtDatafil)
        FileDelete(valgtDatafil)
    FileAppend(jsonOutput, valgtDatafil)
    ; MsgBox valgtDatafil
}

hentVognløbsdata() {
    valgtDatafil := ""

    valgtDatafil := FileSelect()

    if !valgtDatafil
        return

    jsonInput := FileRead(valgtDatafil)
    try {
        jsonObj := JSON.Load(jsonInput)
    } catch Error as e {
        MsgBox e.Message
    }

    goo.excel.gyldigeKolonner := jsonObj["excelGyldigeKolonner"]
    goo.vognløb.vlArray := jsonObj["vognløbsArray"]

}