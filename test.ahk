#Requires AutoHotkey v2.0
#Include p6Navigering.ahk
#Include testvl.ahk
#Include excelObj.ahk

+escape:: ExitApp()

testudrul(p_vl_obj)
{
    P6_aktiver()
    P6_luk_vinduer()
    P6_nav_vognløbsbillede()
    for ugedage in p_vl_obj.vl_data["Ugedage"]
    {
        if !ugedage
            continue
        p_vl_obj.vl_data["Dato"] := ugedage
        testfunkvl(p_vl_obj)
    }
    ; testfunkvl(test_vl2)
    MsgBox "færdig"
    return
}
; loop
; {
;     test_vl.vl_data["MobilnrChf"] += 1
;     testfunkvl(test_vl)
;     ; MsgBox a_index

; }

; p6_åben_vognløb(test_vl)
; p6_åben_vognløb_kørselsaftale(test_vl)
; p6_åben_vognløb_åbningstider(test_vl)
; p6_åben_vognløb_resten(test_vl)
; MsgBox(p6_afslut_indlæsning_vognløb(test_vl))

testfunkvl(p_vl_obj)
{
    ; P6_aktiver()
    ; P6_nav_vognløbsbillede()
    p6_åben_vognløb(p_vl_obj)
    p6_åben_vognløb_kørselsaftale(p_vl_obj)
    p6_åben_vognløb_åbningstider(p_vl_obj)
    p6_åben_vognløb_resten(p_vl_obj)
    p6_afslut_indlæsning_vognløb(p_vl_obj)
    return
}
; tjekUgedagArray := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
; UgedagArray := ["ma", "ti", "on", 0 , "fr", "lø", "sø"]

; ugedagTjek(p_var, p_list)
; {
;     for ugedag in tjekUgedagArray
;     {
;         if p_var = ugedag
;             return 1
;         else
;             return 0
;     }
; }


; xl.excel_fil_long := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
; xl.indlæsfil()

; vl := vlObj()
; vl.IndhentData(xl.excel_data[1])

; MsgBox vl.vl_data["Startzone"]
; MsgBox vl.vl_data["Ugedage"][2]

;     kolonne_nummer := Map(

;         "Budnummer", 0,
;         "Vognløbsnummer", 0,
;         "Kørselsaftale", 0,
;         "Styresystem", 0,
;         "Startzone", 0,
;         "Slutzone", 0,
;         "Hjemzone", 0,
;         "MobilnrChf", 0,
;         "Vognløbskategori", 0,
;         "Planskema", 0,
;         "Statistikgruppe", 0,
;         "undtagneTransportTyper", []

;     )


; for index, navn in kolonne_nummer
;     MsgBox index
