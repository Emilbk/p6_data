#Requires AutoHotkey v2.0

; #Include ../include.ahk

tid1 := A_Now
tid2 := DateAdd(tid1, 39, "Minutes")
tid2 := DateAdd(tid2, 43, "Seconds")

MsgBox FormatTime(tid2, "mm:ss")

tidforskel := DateDiff(tid2, tid1, "Minutes")
tidforskelsec := DateDiff(tid2, tid1, "Seconds")

tidmin := tidforskelsec/60
tidsecrem := Mod(tidforskelsec,60)

MsgBox tidsecrem
MsgBox(Floor(tidforskelsec/60) ":" tidsecrem)
; forskelTid := tid1 - tid1
; forskelTidSecond := Round(forskelTid/1000, 1)

; MsgBox forskelTidSecond
; MsgBox tidslutdateTime
; MsgBox FormatTime(tidslut, "HH:mm:ss")

return
