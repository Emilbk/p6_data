#Include ../test/exelarrayMock.ahk

class excelVerificerData {

    __New(pexceldata) {
        this.excelData := pexceldata

    }
    _gyldigeKolonner := Map(
        "Budnummer", "",
        "Vognløbsnummer", "",
        "Kørselsaftale", "",
        "Styresystem", "",
        "Starttid", "",
        "Sluttid", "",
        "Startzone", "",
        "Slutzone", "",
        "Hjemzone", "",
        "VognmandLinie1", "",
        "VognmandLinie2", "",
        "VognmandLinie3", "",
        "VognmandLinie4", "",
        "VognmandKontaktnummer", "",
        "ChfKontaktNummer", "",
        "Vognløbskategori", "",
        "Planskema", "",
        "Økonomiskema", "",
        "Statistikgruppe", "",
        "Vognløbsnotering", "",
        "Ugedage", "",
        "UndtagneTransporttyper", "",
        "KørerIkkeTransporttyper", "",
    )
    
    _ugyldigeKolonner := Map()

    verificerKolonner(){
        for kolonne in this.excelData[1]
            if !this._gyldigeKolonner.has(kolonne)
                this._ugyldigeKolonner.Set(kolonne, A_Index)
            else
                this._gyldigeKolonner[kolonne] := true
                
            
    }
    
    ugyldigeKolonner{
        get{
            this.verificerKolonner()
            return this._ugyldigeKolonner
        }
    }

    gyldigeKolonner{
        get{
            this.verificerKolonner()
            return this._gyldigeKolonner
        }
    }

}


msgbox excelVerificerData(excelDataMock).gyldigeKolonner["Budnummer"]