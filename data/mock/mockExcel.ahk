class excelIndlæsVlDataMock {

    excelObj := Object()
    gyldigeKolonner :=
        [
            "Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone",
            "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer",
            "ChfKontaktnummer", "Vognløbskategori", "Økonomiskema", "Planskema", "Statistikgruppe", "Vognløbsnotering",
            "Ugedage", "UndtagneTransportTyper", "KørerIkkeTransporttyper"
        ]

    ugyldigeKolonner := []

    kolonnerIBrug := Map()
    sheetArray := [
        [
            "UgyldigTest", "Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone",
            "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer",
            "ChfKontaktnummer", "Vognløbskategori", "Økonomiskema", "Planskema", "Statistikgruppe", "Vognløbsnotering",
            "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage",
            "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper",
            "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper"
        ],
        [
            "UgyldigCelle", "212323", "31400", "3400", "47", "23:00", "05:05*", "årh143", "årh143", "årh143", "Vognmand 1", "Vognmand 2", "Vognmand 3 ", "Vognmand 4", "70112220", "70112210", "FV9", "31400", "31400", "2GVEL", "Vognløbsnotering",
            "25-11-2024", "26-11-2024", "ma", "ti", "on", "to", "fr", "lø", "sø",
            "LAV", "NJA", "TRANSPORT", "TMHJUL", "SYD24", "MIDT24", "FYN", "CROSSER",
            "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"
        ],
        [
            "UgyldigCelle", "212323", "31400", "3400", "47", "23:00", "05:05*", "årh143", "årh143", "årh143", "Vognmand 1", "Vognmand 2", "Vognmand 3 ", "Vognmand 4", "70112220", "70112210", "FV9", "31400", "31400", "2GVEL", "Vognløbsnotering",
            "25-11-2024", "26-11-2024", "ma", "ti", "on", "to", "fr", "lø", "sø",
            "LAV", "NJA", "TRANSPORT", "TMHJUL", "SYD24", "MIDT24", "FYN", "CROSSER",
            "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"
        ]
    ]

    setAktiveKolonner() {

        aktiveKolonner := this.kolonnerIBrug

        for kolonneIndex, kolonneNavn in this.sheetArray[1]
        {

            if this.tjekKolonneErGyldig(kolonneNavn)
            {
                aktiveKolonner.Set(kolonneNavn, {kolonneNavn: kolonneNavn, kolonneNummer: kolonneIndex})
            }
            else
            {
                this.ugyldigeKolonner.Push({ kolonneNavn: kolonneNavn, kolonneNummer: kolonneIndex })
            }

        }
    }

    tjekKolonneErGyldig(pKolonneNavnTilTjek) {

        fundet := 0

        for kolonneIndex, kolonneNavn in this.gyldigeKolonner
            if kolonneNavn = pKolonneNavnTilTjek
                fundet := 1
        return fundet
    }

    dataVerificerUgedage(pUgedagData){

        ugedag := pUgedagData

        if InStr(ugedag, "*")
        {
            
        }
        
        else
        {
            
        }
    }
}


test := excelIndlæsVlDataMock()
test.setAktiveKolonner()
return