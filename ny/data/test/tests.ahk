#Include lib/AutoHotUnit/AutoHotUnit.ahk
#Include lib\ahktimer.ahk
; #Include lib\json.ahk
#Include excel.test.ahk
#Include vl.test.ahk
ahu.RegisterSuite(testExcelHentData)
ahu.RegisterSuite(testExcelDataStruktur)
ahu.RegisterSuite(testExcelVerificerData)
ahu.RegisterSuite(parameterTest)
ahu.RegisterSuite(parameterFactoryTest)
ahu.RegisterSuite(vltest)


 
ahu.RunSuites()

