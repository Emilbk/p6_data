#Include lib\json.ahk
; Fra VLMock.xlsx


class excelMock {

    static excelDataGyldig {
        get {
            return [["Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone", "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer", "ChfKontaktnummer", "Vognløbskategori", "Planskema", "Økonomiskema", "Statistikgruppe", "Vognløbsnotering", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper"],
                ["24-2267", "31400", "3400", "1", "06:00", "15:00", "årh143", "årh143", "årh143", "Vognmand 1 ApS", "Lukket pr. 12-11-24", "Gadenr 1", "8001 By", "70112220", "701122010", "FV8", "31400", "31400", "2GVEL", "Blabla", "17/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"],
                ["24-2268", "31401", "3401", "2", "23:00", "04:00*", "årh144", "årh144", "årh144", "Vognmand 2 ApS", "Lukket pr. 13-11-24", "Gadenr 2", "8002 By", "70112221", "2", "FV9", "31401", "31401", "3GVEL", "Blabla 2", "18/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"]]
        }
    }
    static excelDataUgyldigFlere {
        get {
            return [["Ugyldig", "Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone", "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer", "ChfKontaktnummer", "Vognløbskategori", "Planskema", "Økonomiskema", "Statistikgruppe", "Vognløbsnotering", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "ugyldig", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper"],
                ["ugyldig", "24-2267", "31400", "forMangeTegn", "1", "06:00", "15:00", "forLangtParameter", "årh143", "årh143", "Vognmand 1 ApS", "Lukket pr. 12-11-24", "Gadenr 1", "8001 By", "70112220", "70112211", "FV8", "31400", "31400", "2GVEL", "Blabla", "17/11/2024", "42/11/2024", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "ugyldig", "center", "tripstol", "forMangeArray", "forLangtParameterIArray" "", '', ''],
                ["ugyldig", "24-2268", "31401", "3401", "2", "06:01", "15:01", "forLangtParameter", "årh144", "årh144", "Vognmand 2 ApS", "Lukket pr. 13-11-24", "Gadenr 2", "8002 By", "70112221", "2", "FV9", "31401", "31401", "3GVEL", "Blabla 2", "18/11/2024", "ma", "ti", "on", "torsdag", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "ugyldig", "center", "tripstol", "forMangeArray", 'forLangtParameterIArray', '', '', '']]
        }
    }
    static excelDataUgyldigDato42112024 {
        get {
            return [["Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone", "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer", "ChfKontaktnummer", "Vognløbskategori", "Planskema", "Økonomiskema", "Statistikgruppe", "Vognløbsnotering", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "UndtagneTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper", "KørerIkkeTransportTyper"],
                ["24-2268", "31401", "3401", "2", "06:01", "15:01", "årh144", "årh144", "årh144", "Vognmand 2 ApS", "Lukket pr. 13-11-24", "Gadenr 2", "8002 By", "70112221", "2", "FV9", "31401", "31401", "3GVEL", "Blabla 2", "42/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"]]
        }
    }

    static excelMockfil {

        get {

            return A_ScriptDir "\data\test\assets\VLMock.xlsx"
        }
    }

}