#Requires AutoHotkey v2.0
#Include excelObj.ahk
#Include vlObj.ahk
#Include p6Navigering.ahk

+Escape::Pause


xl := excelObj()
xl.excel_fil_long := "C:\Users\ebk\makro\p6_data\VL.xlsx"
xl.indlæsfil()

vl := vlObj()
vl.IndhentDataArray(xl.excel_data)


P6_aktiver()
P6_luk_vinduer()
P6_nav_vognløbsbillede()
for index, vl in vl.vl_array
{
    for index, vl_dag in vl["Ugedage"]
        if vl_dag
        {
            vl["Dato"] := vl_dag
            p6_åben_vognløb(vl)
            p6_åben_vognløb_kørselsaftale(vl)
            p6_åben_vognløb_åbningstider(vl)
            p6_åben_vognløb_resten(vl)
            p6_afslut_indlæsning_vognløb(vl)
        }
}

; testfunkvl(test_vl2)
MsgBox "færdig"
return