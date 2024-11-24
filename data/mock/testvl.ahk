#Requires AutoHotkey v2.0
#Include ../include.ahk

vinduehandle := 1706740
testVl := VognløbObj()

testVl.parametre.Vognløbsnummer.forventetIndhold := "31400"
testVl.parametre.Kørselsaftale.forventetIndhold := "3400"
testVl.parametre.Styresystem.forventetIndhold := "1"
testVl.parametre.Vognløbsdato.forventetIndhold := "MA"
testVl.parametre.VognløbsdatoSlut.forventetIndhold := "MA"
testVl.parametre.Starttid.forventetIndhold := "01:00"
testVl.parametre.Sluttid.forventetIndhold := "02:00"
testVl.parametre.Hjemzone.forventetIndhold := "Årh144"
testVl.parametre.Vognløbsnotering.forventetIndhold := "TESTVOGNLØB TESTVOGNLØB TESTVOGNLØB"
testVl.parametre.UndtagneTransporttyper.iBrug := 1
testVl.parametre.UndtagneTransporttyper.forventetIndhold := ["Crosser", "Barn2," "NJA", "Barn1", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]
testVl.parametre.KørerIkkeTransporttyper.forventetIndhold := ["Crosser", "Barn2," "NJA", "Barn1", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]
testVl.parametre.VognmandLinie1.forventetIndhold := "Vogmand1"


testVl.parametre.Vognløbsnummer.eksisterendeIndhold := "31400"
testVl.parametre.Kørselsaftale.eksisterendeIndhold := "3400"
testVl.parametre.Styresystem.eksisterendeIndhold := "2"
testVl.parametre.Vognløbsdato.eksisterendeIndhold := "MA"
testVl.parametre.VognløbsdatoSlut.eksisterendeIndhold := "MA"
testVl.parametre.Starttid.eksisterendeIndhold := "01:00"
testVl.parametre.Sluttid.eksisterendeIndhold := "02:00"
testVl.parametre.Hjemzone.eksisterendeIndhold := "Årh144"
testVl.parametre.Vognløbsnotering.eksisterendeIndhold := "TESTVOGNLØB TESTVOGNLØB TESTVOGNLØB"
testVl.parametre.UndtagneTransporttyper.iBrug := 1
testVl.parametre.UndtagneTransporttyper.eksisterendeIndhold := ["Nja", "CrosSER", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]
testVl.parametre.KørerIkkeTransporttyper.eksisterendeIndhold := ["Crosser", "Barn3," "NJA", "Barn1", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]



testVl.parametre.sorterUndtagneTransporttyperEksisterende()
testVl.parametre.sorterUndtagneTransporttyperForventet()
testVl.parametre.sorterKørerIkkeTransporttyperEksisterende()
testVl.parametre.sorterKørerIkkeTransporttyperForventet()

testVl.parametre.tjekUndtagneTransportTyperEns()
testVl.parametre.tjekKørerIkkeTransportTyperEns()

testVl.parametre.tjekAlleParameterForFejl()


; testExcelPath := A_ScriptDir "\exceltest\test-" FormatTime(, "dd-HH-mm-ss") ".xlsx" ; Saves in the same folder as the script
testExcelPath := A_ScriptDir "\exceltest\test2.xlsx" 

if !FileExist(testExcelPath)
    excelNyWorkbook := excelLavNyWorkbook(testExcelPath)

testExcel := udfyldTestExcelArk()
; testExcel.app.Visible := 1
testExcel.lavExcelTemplate()
testExcel.åbenWorkbookReadWrite(testExcelPath)
; testExcel.app.Worksheets.add()
; testExcel.app.Worksheets.add()
; testExcel.setAktivSheet(1)
; testExcel.navngivSheet(1, "Alle gyldige kolonnemuligheder")
; testExcel.navngivSheet(2, "Nyt ark2")
; testExcel.navngivSheet(3, "Nyt ark3")


; vl.parametre.tjekAlleParameterForFejl()
; MsgBox "test"
; return