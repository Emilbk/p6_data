#Requires AutoHotkey v2.0

test1 := Map("et", 1, "to", 2)
test2 := test1.Clone()
test2["et"] := "one"
return