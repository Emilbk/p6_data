#Requires AutoHotkey v2.0
#Include excelClassP6Data.ahk


kolonneNavne := ["asd", "asd1", "asd2", "asd3", "asd4", "asd5"]
rakker := [["31400", "3400", "1"], ["31400", "3400", "1"], ["31401", "3401", "2"]]

xl := excelObjP6Data()
; xl2 := excelObjP6Data()
xl.app.visible := true

xl.app.Workbooks.Add()
; xl2.app.Workbooks.Add()

; MsgBox xl.app.activeWorkbook.Name
workbook := xl.app.activeWorkbook
savePath := A_ScriptDir "\exceltest\test-" FormatTime(, "dd-HH-mm-ss") ".xlsx" ; Saves in the same folder as the script
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

rakkenummer := 2
for Rakke in rakker
{
    for kolonneNummer, rakkeIndhold in Rakke
    {
        aktivCelle := activeSheet.cells(rakkenummer, kolonneNummer)
        aktivCelle.Value := StrTitle(rakkeIndhold)
        aktivCelle.Interior.Colorindex := 4
    }
    rakkenummer += 1

}

aktivCelle := activeSheet.cells(2, 3)
aktivCelle.Interior.Colorindex := 3
aktivCelle.addcomment()
aktivCelle.comment.text("Fejl: Blablalba")
; aktivCelle.comment.Visible := True
; workbook.

workbook.save()
xl.quit()
return