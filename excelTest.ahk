#Requires AutoHotkey v2.0
#Include excelClassP6Data.ahk


kolonneNavne := ["asd", "asd1", "asd2", "asd3", "asd4", "asd5"]

xl := excelObjP6Data()
; xl2 := excelObjP6Data()
xl.app.visible := true

xl.app.Workbooks.Add()
; xl2.app.Workbooks.Add()

; MsgBox xl.app.activeWorkbook.Name
workbook := xl.app.activeWorkbook
savePath := A_ScriptDir "\exceltest\test-" FormatTime(,"dd-HH-mm-ss") ".xlsx" ; Saves in the same folder as the script
workbook.Saveas(savepath)
; MsgBox xl.app.activeWorkbook.Name
; MsgBox xl2.app.activeWorkbook.Name

activeSheet := workbook.activeSheet


activeSheet.Name := "VognløbLog"


kolonneRække := 1
colorindex := 36
for kolonne in kolonneNavne
{
    
    aktivCelle := activeSheet.cells(kolonneRække, A_Index) 
    aktivCelle.Value := StrTitle(kolonne)
    aktivCelle.Interior.Colorindex := colorindex
    ; aktivCelle.addcomment()
    ; aktivCelle.comment.text("Komment nr" colorindex)
    aktivCelle.addcommentthreaded("Kommentar nr. " colorindex)
    colorindex += 1

}


; workbook.

workbook.save()
xl.quit()
return 