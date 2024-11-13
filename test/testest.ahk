obj := {}
obj[] := Map()     ; Equivalent to obj.__Item := Map()
obj["base"] := 10
obj.ny := "test"

MsgBox obj.base = Object.prototype  ; True
MsgBox obj["base"]                  ; 10


return