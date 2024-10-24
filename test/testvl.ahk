#Requires AutoHotkey v2.0

#Include vlObj.ahk
#Include dataGUI.ahk
#Include excelObj.ahk

p_vl_obj := vlObj()

p_vl_obj.vl_data["Vognløbsnummer"] := "31400"
p_vl_obj.vl_data["Kørselsaftale"] := "3400"
p_vl_obj.vl_data["Styresystem"] := "1"
p_vl_obj.vl_data["Planskema"] := "31300" 
p_vl_obj.vl_data["Økonomiskema"] := "31200"
p_vl_obj.vl_data["Startzone"] := "Årh804"
p_vl_obj.vl_data["Slutzone"] := "Årh804"
p_vl_obj.vl_data["Hjemzone"] := "Årh804"
; test_vl.vl_data["Vognløbsnotering"] := 0
p_vl_obj.vl_data["Vognløbsnotering"] := "Ny notering til VL"
p_vl_obj.vl_data["MobilnrChf"] := "70112210"
p_vl_obj.vl_data["Statistikgruppe"] := "2GVEL"
; test_vl.vl_data["Undtagne transporttyper"] := 0
p_vl_obj.vl_data["Undtagne transporttyper"] := ["LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER" ]
p_vl_obj.vl_data["Dato"] := ["ma"]
p_vl_obj.vl_data["Starttid"] := "08:00"
p_vl_obj.vl_data["Sluttid"] := "17:00"
p_vl_obj.vl_data["Vognløbskategori"] := "FV8"
p_vl_obj.vl_data["Ugedage"] := ["ma", "ti", 0, "to", "fr", "lø", "sø"]

test_vl2 := vlObj()

test_vl2.vl_data["Vognløbsnummer"] := "31400"
test_vl2.vl_data["Kørselsaftale"] := "3400"
test_vl2.vl_data["Styresystem"] := "1"
test_vl2.vl_data["Planskema"] := "31300" 
test_vl2.vl_data["Økonomiskema"] := "31200"
test_vl2.vl_data["Startzone"] := "Årh804"
test_vl2.vl_data["Slutzone"] := "Årh804"
test_vl2.vl_data["Hjemzone"] := "Årh804"
; test_2vl.vl_data["Vognløbsnotering"] := 0
test_vl2.vl_data["Vognløbsnotering"] := "Ny notering til VL"
test_vl2.vl_data["MobilnrChf"] := "70112210"
test_vl2.vl_data["Statistikgruppe"] := "2GVEL"
; test_2vl.vl_data["Undtagne transporttyper"] := 0
test_vl2.vl_data["Undtagne transporttyper"] := ["LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER" ]
test_vl2.vl_data["Dato"] := ["TI"]
test_vl2.vl_data["Starttid"] := "08:15"
test_vl2.vl_data["Sluttid"] := "17:35"
test_vl2.vl_data["Vognløbskategori"] := "FV8"
test_vl2.vl_data["Ugedage"] := ["ma", "ti", "on", "to", "fr", "lø", "sø"]

