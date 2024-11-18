#Include ../include.ahk

#hotif WinActive("P6-makro")
Escape:: guiLuk()
#Hotif
+Escape:: ExitApp()

goo := Gui()
goo.Title := "P6-makro"
goo.tekst := {}
goo.knap := {}

; Objs
goo.excel := {}
goo.excel.obj := Object()
goo.vognløb := {}
; goo.vognløb.constr := VognløbConstructor()


goo.p6 := {}
goo.p6.obj := p6()
goo.p6.vinduehandle := ""


gooMenuBar := MenuBar()

filMenu := Menu()
dataMenu := Menu()
excelMenu := Menu()
databehandlingMenu := Menu()
ændrDataMenu := Menu()
p6Menu := Menu()


; filMenu.Add "&Open`tCtrl+O", (*) => FileSelect()  ; See remarks below about Ctrl+O.
filMenu.Add
filMenu.Add "E&xit", (*) => ExitApp()

dataMenu.Add("&Excel", excelMenu)
dataMenu.Add "&P6-data", databehandlingMenu

excelMenu.Add("&Indlæs Excel-fil`tCtrl+E", (*) => indlæsExcel())
excelMenu.Add("&Dan Excel-skabelon", (*) => danExcelSkabelon())

databehandlingMenu.Add("Ændr data", ændrDataMenu)
databehandlingMenu.Add("Indhent data", (*) =>)
databehandlingMenu.Add("Fejlcheck data", (*) =>)


ændrDataMenu.Add("Ændr hjemzone", (*) => ændrVognløbHjemzone())
ændrDataMenu.Add("Ændr vognløb", (*) => ændrVognløbAlt())

p6Menu.Add("Vælg P6-Vindue`tctrl+P", (*) => vælgP6Vindue())

gooMenuBar.Add("&Fil", filMenu)
gooMenuBar.Add("&Data", dataMenu)
gooMenuBar.Add("&P6", p6Menu)

goo.MenuBar := gooMenuBar


excelfilTekst := "Ingen fil"
goo.tekst.valgtP6Vindue := goo.Add("Text", , "Aktivt P6-vindue: ")
goo.knap.valgtP6VindueKnap := goo.Add("Button", "XP+80 YP-5", "Aktiver valgt")
goo.tekst.indlæstExcelFil := goo.Add("Text", "XM", "Indlæst Excel-fil: " excelfilTekst)
; goo.tekst.indlæstExcelRækkerTekst := goo.Add("Text", , "Antal Rækker: ")
goo.knap.valgtP6VindueKnap.OnEvent("Click", (*) => goo.p6.obj.navAktiverP6Vindue())

goo.tekst.indlæstExcelFil.text := "Indlæst Excel-fil: " excelfilTekst


goo.OnEvent("Close", guiLuk)
goo.Show("W300 H200")

guiLuk(*){

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
    MsgBox goo.p6.vindueHandle

    goo.p6.obj.setP6Vindue(goo.p6.vindueHandle)
    ; goo.p6.obj.setP6Vindue()
}


indlæsExcel() {

    goo.excel.valgtExcelFil := FileSelect()
    if !goo.excel.valgtExcelFil
        return

    SplitPath(goo.excel.valgtExcelFil, &excelFil, &excelDir, &excelExt, &excelIngenExt)

    goo.excel.valgtExcelFilKort := excelIngenExt
    goo.Hide()

    goo.excel.obj := excelIndlæsVlData(goo.excel.valgtExcelFil, 1)
    goo.excel.vlArray := goo.excel.obj.getVlArray()
    goo.excel.gyldigeKolonner := goo.excel.obj.getGyldigeKolonner()
    MsgBox "Indlæst!", "Excel"
    goo.Show()
    goo.tekst.indlæstExcelFil.text := "Indlæst Excel-fil: " excelFil

    goo.vognløb.constr := VognløbConstructor(goo.excel.vlArray, goo.excel.gyldigeKolonner)
    goo.vognløb.vlArray := goo.vognløb.constr.getBehandletVognløbsArray()


    return
}

danExcelSkabelon() {

    excelPath := A_ScriptDir "\excelSkabelon.xlsx"

    testExcel := udfyldTestExcelArk()
    testExcel.lavExcelTemplate(excelPath)


    MsgBox("Excelskabelon gemt som " excelpath, "Excel", "iconi")
}

ændrVognløbAlt() {
    p6Obj := goo.p6.obj
    if p6Obj.vindueHandle = ""
    {
        MsgBox("P6-vindue er ikke valgt endnu!")
        return
    }

    p6obj.navAktiverP6Vindue()
    p6Obj.navLukAlleVinduer()
    for vlSamling in goo.vognløb.vlArray
    {
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                vl.tjekForbudtVognløbsDato()
                p6obj.funkÆndrVognløb()

            } catch P6MsgboxError as msg {
                continue

            }
        }

    }

    MsgBox("Excelark færdigindlæst.", "Vognløb ændret!", "Iconi")
    return
}
ændrVognløbHjemzone() {
    p6Obj := goo.p6.obj
    if p6Obj.vindueHandle = ""
    {
        MsgBox("P6-vindue er ikke valgt endnu!")
        return
    }
        goo.VognløbConstructor.vlArray.masterVognløb := vlKørselaftale
        p6obj.setVognløb(vlKørselaftale)
        p6obj.navAktiverP6Vindue()
        p6obj.navLukAlleVinduer()
        p6obj.setVognløb(vlKørselsaftale)
        try {
            p6obj.funkKørselsaftaleÆndrHjemzone()
            
        } catch Error as e {
           ; kørselsaftalefejl 
        }

    for vlSamling in goo.vognløb.vlArray.vognløbsListe
    {
        for Vl in vlSamling
        {
            p6Obj.setVognløb(vl)
            try {
                vl.tjekForbudtVognløbsDato()
                p6obj.funkVognløbsbilledeÆndrHjemzone()

            } catch P6MsgboxError as msg {
                continue
            }
        }

    }
    MsgBox("Excelark færdigindlæst.", "Vognløb ændret!", "Iconi")

    return
}