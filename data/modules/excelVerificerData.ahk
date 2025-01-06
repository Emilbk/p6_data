#Include ../test/exelarrayMock.ahk
#Include includeModules.ahk
class excelVerificerData {

    __New(pexceldata) {
        this.excelData := pexceldata

    }
    parametre := parameter.data
    _gyldigeKolonner := gyldigeKolonner.data

    _ugyldigeKolonner := Map()

    verificerKolonner() {
        for kolonne in this.excelData[1]
            if !this._gyldigeKolonner.has(kolonne)
                this._ugyldigeKolonner.Set(kolonne, A_Index)
            else
                this._gyldigeKolonner[kolonne] := true


    }

    ugyldigeKolonner {
        get {
            this.verificerKolonner()
            return this._ugyldigeKolonner
        }
    }

    gyldigeKolonner {
        get {
            this.verificerKolonner()
            return this._gyldigeKolonner
        }
    }

}


tets := excelVerificerData(excelDataMock)

MsgBox tets.excelData[1][2]

return