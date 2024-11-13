#Include ../modules/excelClass.ahk
#Include ../modules/excelClassP6Data.ahk


excelTest := excelObjP6Data()
excelPath := A_ScriptDir "\exceltest\test-" FormatTime(, "dd-HH-mm-ss") ".xlsx" ; Saves in the same folder as the script
excelTest.setWorkbookSavePath(excelPath)
excelTest.setAktivWorkbookDir(excelPath)
excelTest.lavNyWorkbook()
; excelTest.setAktivWorksheetName("Vognløbstest")

excelTest.skrivExcelKolonneNavn(excelTest.testKolonneNavnOgNummerArray)


MsgBox "test"
return