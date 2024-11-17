
teststr := "TI*"

if (IsInteger(SubStr(teststr,1, 2)) and InStr(teststr, "*"))
    MsgBox "ja"

år := SubStr(teststr, -4, 4)
mned := SubStr(teststr, 4, 2)
dag := SubStr(teststr, 1, 2)

datestr := år mned dag
msgbox FormatTime(datestr,"yyyy")

nydate := DateAdd(datestr,1,"Days")

teststrny := SubStr(teststr, 1, 2)

teststrny += 1

if IsInteger(teststrny)
    msgbox "ja"

return