#Include ../modules/excelClass.ahk
#Include ../modules/excelClassP6Data.ahk
#Include ../mock.ahk

test := excelIndlæsVlData()
excelFil := "C:\Users\ebk\makro\p6_data\assets\VL - Kopi.xlsx"
excelFilNy := "C:\Users\ebk\makro\p6_data\assets\VL - Kopi2.xlsx"

; test.vælgWorkbook(excelFil)
; test.LavNyWorkbook(excelFil)
; test.LavNyWorkbook(excelFil)
; test.åbenWorkbookReadonly(excelFil)

test.helperIndlæsAlt(excelFil, 1)

; test.app.Visible := "True"
test.quit()

nyExcel := excelObjP6Data()

nyExcel.åbenWorkbookReadWrite(excelFilNy)
nyexcel.app.visible := "True"
; nyExcel.LavNyWorkbook(excelFilNy)
nyExcel.aktivWorksheetKolonneNavnOgNummer := test.aktivWorksheetKolonneNavnOgNummer

nyExcel.kolonneNavnogNummerTilArray(nyExcel.aktivWorksheetKolonneNavnOgNummer)
nyExcel.setAktivRække(2)

excelFil := "C:\Users\ebk\makro\p6_data\assets\VL.xlsx"
excelArk := 1
excelobj := excelObjP6Data()
excelobj.get(excelFil, excelArk)
excelobj.quit()

; excelMock := mockExcelP6Data()
excelRækkeArray := excelobj.getRækkeData()
vlConstruct := VognløbConstructor()
vlContainer := vlConstruct.behandlVognløsbsdata(excelRækkeArray)
vognlob := vlContainer[1][1]

vlresultat := nyExcel.kolonneNavnogNummerTilArray(nyExcel.vlResultatKolonneNavnOgNummer)
nyExcel.skrivExcelKolonneNavn(vlresultat)

nyExcel.AktivRække := 2
for vlSæt in vlContainer
    for vl in vlsæt
{
    nyExcel.skrivExcelVognløb(vl, vlresultat)
    nyExcel.aktivrække += 1
    
}


; test.vælgAktivWorksheet(1)
; kolonne := Array("31400", "47")
; test.skrivExcelKolonneNavn(kolonne)

; test.gemWorkbook()


test.quit()

return