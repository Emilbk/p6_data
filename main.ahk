#Requires AutoHotkey v2.0
Persistent
; omskriv til bedre organisering
#Include modules
#Include excelClassP6Data.ahk
; #Include dataGUI.ahk
#Include vlClass.ahk
#Include p6.ahk
; #Include test.ahk

+escape:: ExitApp()

udrulÆndringer()
{

    excelpath := "C:\Users\ebk\makro\p6_data\VL.xlsx"
    excel := excelObjP6Data()
    excel.filVælgExcelFil(excelpath)
    excel.helperIndlæsAlt(1)
    excel.quit()

    vldata := excel.getData()

    testp6 := P6()
    testvl := VognløbObj()
    testp6.navAktiverP6Vindue()
    testp6.navLukVinduer()
    for vognløb in vldata
    {
        testvl.indhentVognløbsdata(vognløb)
        testvl.opretVognløbForHverDato()

        for vognløbsdag, vognløb in testvl.Vognløb
        {
            testp6.dataIndhentVlObj(vognløb)
            testp6.funkÆndrVognløb()
            testp6.funkTjekVognløb()
        }

        ; MsgBox "Færdig " vognløb["Vognløbsnummer"]
        ; testvl.IndlæsteVognløb.push(vognløb["Vognløbsnummer"])
    }
    MsgBox "Done!"


}

; udrulÆndringer()

; DataGUI.excelData := excelIndlæsArr(excel_fil)


; vælgExcelFilTest()

; test.vlIndhentData(DataGUI.excelData, 2)
; DataGUI.Show("AutoSize")

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
