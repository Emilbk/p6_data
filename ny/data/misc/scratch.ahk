#Include ../modules/includeModules.ahk
; #Include ../modules/excelHentData.ahk
;#Include ../test/exelarrayMock.ahk
;#Include ../modules/gyldigeKolonner.ahk
;#Include ../modules/parameter.ahk



class Bank {
    ; Define a dynamic property with Get and Set methods
    AmountOfMoney {
        get {
            return this._amountOfMoney
        }
        set {
            if (value < 0)
                throw Error("Amount of money cannot be negative!")
            this._amountOfMoney := value
        }
    }

    ; Constructor to initialize the property
    __New(initialAmount := 0) {
        this.AmountOfMoney := initialAmount  ; Use the setter to initialize
    }
}

; Create an instance of the Bank class
myBank :=  Bank(1000)

; Get the AmountOfMoney property
MsgBox myBank.AmountOfMoney  ; Outputs: 1000

; Set the AmountOfMoney property
myBank.AmountOfMoney := 500
MsgBox myBank.AmountOfMoney  ; Outputs: 500

; Try to set a negative value
try {
    myBank.AmountOfMoney := -200  ; Throws an exception
} catch error as e {
    MsgBox e.Message  ; Outputs: Amount of money cannot be negative!
}

return