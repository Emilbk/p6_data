#Include lib/AutoHotUnit/AutoHotUnit.ahk
#Include lib\ahktimer.ahk
#Include lib\json.ahk
#Include excel.test.ahk
#Include vl.test.ahk
; ahu.RegisterSuite(testExcelHentData)
; ahu.RegisterSuite(testExcelDataStruktur)
; ahu.RegisterSuite(testExcelVerificerData)
ahu.RegisterSuite(excel)


ahu.RunSuites()