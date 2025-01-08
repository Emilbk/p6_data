#Include ../test/lib/json.ahk
; #Include ../modules/includeModules.ahk

class test {

        static ny {

            get{
                data := Map("test1", 1)
        
                return data

            }
        }
}

t := test.ny
t2 := test

MsgBox t2 is test
MsgBox t["test1"]

return