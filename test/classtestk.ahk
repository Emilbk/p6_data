
class maincl {
    
    __New(vognløb) {
        
        this.vognløb := vognløb

        this.nested := maincl.nestedcl()
    }

    class nestedcl {

        ; __New(vognløb) {
           
        ;     this.vognløb := vognløb
        ; }

        testfunk(){

            MsgBox "test"
            ; MsgBox this.vognløb.Vognløbsnummer
        }
        
    }
}

; test := Object()

; vognløb := {Vognløbsnummer: "312400", Styresystem: "47"}
; main := maincl(vognløb)
; main.nested.testfunk()
; Define an object literal
obj := {Name: "Alice", Age: 30}

; Access properties using bracket notation
MsgBox(obj.Name)         ; Should display "Alice"
MsgBox(obj.Age)          ; Should display "30"

; Test dynamic addition of properties
obj.Occupation := "Engineer"
MsgBox(obj.Occupation)   ; Should display "Engineer"

; Attempt bracket notation access, standard in v2
MsgBox(obj["Name"])      ; Should display "Alice"

return