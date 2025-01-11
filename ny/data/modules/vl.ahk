#Requires AutoHotkey v2.0

class vognløb {
    __New(parametre) {
        this.parametre := parametre

    }

    vognløbsnummer{

        get{

            return this.parametre["Vognløbsnummer"].forventet
        }
    }

;     parameter {
;         set {

;             this._parameter := value
;         }
;     get => this._parameter
;    }
}
