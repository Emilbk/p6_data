#Requires AutoHotkey v2.0
Persistent
; omskriv til bedre organisering
#Include excelObj.ahk
#Include dataGUI.ahk
#Include vlObj.ahk
#Include p6Navigering.ahk
#Include test.ahk


test := vlObj()

excel_fil := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
; DataGUI.excelData := excelIndlæsArr(excel_fil)


; vælgExcelFilTest()

; test.vlIndhentData(DataGUI.excelData, 2)
DataGUI.Show("AutoSize")

; vælgExcelFilTest()
; {
;     ; DataGUI.opt("+Disabled")
;     ; WinActivate(DataGUINavn)
;     ; valgtExcelFilLong := FileSelect()
;     ; if !valgtExcelFilLong
;     ; return
;     ; SplitPath(excel_fil, &valgtExcelFil)
;     excel_fil := "C:\Users\ebk\Trafikstyring V2\P6data\VL.xlsx"
;     indlæstExcelFilTekst := "Indlæst excel-fil: " . excel_fil
;     overskriftExcelfil.Text := indlæstExcelFilTekst
;     DataGUI.excelData := excelIndlæsArr(excel_fil)

;     ; listview
;     dataListview.Delete()
;     columnNumber := dataListview.GetCount("Col")
;     if columnNumber != 0
;         loop columnNumber
;             dataListview.DeleteCol(1)
;     for i, e in DataGUI.excelData[1]
;     {
;         dataListview.InsertCol(i, , e)
;     }

;     for i, e in DataGUI.excelData
;         if i > 1
;         {
;             dataListview.Insert(1, , e*)
;             ;dataListview.Insert(i, , DataGUI.excelData[i])
;         }
;     dataListview.ModifyCol()

;     return
; }

; vælgExcelFil()
; {
;     ; DataGUI.opt("+Disabled")
;     ; WinActivate(DataGUINavn)
;     valgtExcelFilLong := FileSelect()
;     if !valgtExcelFilLong
;         return
;     SplitPath(valgtExcelFilLong, &valgtExcelFil)
;     indlæstExcelFilTekst := "Indlæst excel-fil: " . valgtExcelFil
;     overskriftExcelfil.Text := indlæstExcelFilTekst
;     DataGUI.excelData := excelIndlæsArr(valgtExcelFilLong)

;     ; listview
;     dataListview.Delete()
;     columnNumber := dataListview.GetCount("Col")
;     if columnNumber != 0
;         loop columnNumber
;             dataListview.DeleteCol(1)
;     for i, e in DataGUI.excelData[1]
;     {
;         dataListview.InsertCol(i, , e)
;     }

;     for i, e in DataGUI.excelData
;         if i > 1
;         {
;             ; +1 for at indsætte fra bunden
;             dataListview.Insert(DataGUI.excelData.Length + 1, , e*)
;             ;dataListview.Insert(i, , DataGUI.excelData[i])
;         }
;     dataListview.ModifyCol()
;     ; lav bedre løsning
;     vl := vlObj()
;     vl.vlIndhentData(DataGUI.excelData, 2)
;     return
; }

; opdaterExcelRække()
; {

; }

; testfunk(*)
; {
;     loop DataGUI.excelData.Length
;     {
;         vl := vlObj()
;         if (A_Index >= 2 or A_Index <= (A_Index - 1))
;         {
;             vl.vlIndhentData(DataGUI.excelData, A_Index)
;             vl.p6IndlæsData()
;         }
;     }
; }