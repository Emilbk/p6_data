class gyldigeKolonner{
    
    static _gyldigeKolonner := Map(
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
    
    static data{
        get{
            return gyldigeKolonner._gyldigeKolonner
        }
    }
}