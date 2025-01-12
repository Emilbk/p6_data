#Requires AutoHotkey v2.0

class vognløb {
    __New(parameter) {
        this.parameter := parameter
        
        ; this.Styresystem := { forventet: this.parameter["Styresystem"].forventet, faktisk: this.parameter["Styresystem"
        ;     ].faktisk }
        ; this.Kørselsaftale := { forventet: this.parameter["Kørselsaftale"].forventet, faktisk: this.parameter[
        ;     "Kørselsaftale"].faktisk }
        ; this.Vognløbsnummer := { forventet: this.parameter["Vognløbsnummer"].forventet, faktisk: this.parameter[
        ;     "Vognløbsnummer"].faktisk }
        ; this.vognløbsdato := {}
        ; this.Vognløbsdato := { forventet: this.parameter["Vognløbsdato"].forventet, faktisk: this.parameter[
        ;     "Vognløbsdato"].faktisk }
        ; this.Budnummer := { forventet: this.parameter["Budnummer"].forventet, faktisk: this.parameter["Budnummer"].faktisk }
        ; this.ChfKontaktnummer := { forventet: this.parameter["ChfKontaktnummer"].forventet, faktisk: this.parameter[
        ;     "ChfKontaktnummer"].faktisk }
        ; this.Hjemzone := { forventet: this.parameter["Hjemzone"].forventet, faktisk: this.parameter["Hjemzone"].faktisk }
        ; this.KørerIkkeTransportTyper := { forventet: this.parameter["KørerIkkeTransportTyper"].forventet, faktisk: this
        ;     .parameter["KørerIkkeTransportTyper"].faktisk }
        ; this.Planskema := { forventet: this.parameter["Planskema"].forventet, faktisk: this.parameter["Planskema"].faktisk }
        ; this.Sluttid := { forventet: this.parameter["Sluttid"].forventet, faktisk: this.parameter["Sluttid"].faktisk }
        ; this.Starttid := { forventet: this.parameter["Starttid"].forventet, faktisk: this.parameter["Starttid"].faktisk }
        ; this.Startzone := { forventet: this.parameter["Startzone"].forventet, faktisk: this.parameter["Startzone"].faktisk }
        ; this.Statistikgruppe := { forventet: this.parameter["Statistikgruppe"].forventet, faktisk: this.parameter[
        ;     "Statistikgruppe"].faktisk }
        ; this.UndtagneTransportTyper := { forventet: this.parameter["UndtagneTransportTyper"].forventet, faktisk: this.parameter[
        ;     "UndtagneTransportTyper"].faktisk }
        ; this.Vognløbskategori := { forventet: this.parameter["Vognløbskategori"].forventet, faktisk: this.parameter[
        ;     "Vognløbskategori"].faktisk }
        ; this.VognmandKontaktnummer := { forventet: this.parameter["VognmandKontaktnummer"].forventet, faktisk: this.parameter[
        ;     "VognmandKontaktnummer"].faktisk }
        ; this.VognmandLinie1 := { forventet: this.parameter["VognmandLinie1"].forventet, faktisk: this.parameter[
        ;     "VognmandLinie1"].faktisk }
        ; this.VognmandLinie2 := { forventet: this.parameter["VognmandLinie2"].forventet, faktisk: this.parameter[
        ;     "VognmandLinie2"].faktisk }
        ; this.VognmandLinie3 := { forventet: this.parameter["VognmandLinie3"].forventet, faktisk: this.parameter[
        ;     "VognmandLinie3"].faktisk }
        ; this.VognmandLinie4 := { forventet: this.parameter["VognmandLinie4"].forventet, faktisk: this.parameter[
        ;     "VognmandLinie4"].faktisk }
        ; this.Økonomiskema := { forventet: this.parameter["Økonomiskema"].forventet, faktisk: this.parameter[
        ;     "Økonomiskema"].faktisk }
    }
    vognløbsdatoForventet{
        set{
            this.parameter["Vognløbsdato"].forventet := Value
        }
        get => this.parameter["Vognløbsdato"].forventet
    }
    vognløbsdatoFaktisk{
        set{
            this.parameter["Vognløbsdato"].faktisk := Value
        }
        get => this.parameter["Vognløbsdato"].faktisk
    }
    
    StyresystemForventet{
        set{
            this.parameter["Styresystem"].forventet := Value
        }
        get => this.parameter["Styresystem"].forventet
    }
    StyresystemFaktisk{
        set{
            this.parameter["Styresystem"].faktisk := Value
        }
        get => this.parameter["Styresystem"].faktisk
    }
    
    KørselsaftaleForventet{
        set{
            this.parameter["Kørselsaftale"].forventet := Value
        }
        get => this.parameter["Kørselsaftale"].forventet
    }
    KørselsaftaleFaktisk{
        set{
            this.parameter["Kørselsaftale"].faktisk := Value
        }
        get => this.parameter["Kørselsaftale"].faktisk
    }
    
    VognløbsnummerForventet{
        set{
            this.parameter["Vognløbsnummer"].forventet := Value
        }
        get => this.parameter["Vognløbsnummer"].forventet
    }
    VognløbsnummerFaktisk{
        set{
            this.parameter["Vognløbsnummer"].faktisk := Value
        }
        get => this.parameter["Vognløbsnummer"].faktisk
    }
    
    BudnummerForventet{
        set{
            this.parameter["Budnummer"].forventet := Value
        }
        get => this.parameter["Budnummer"].forventet
    }
    BudnummerFaktisk{
        set{
            this.parameter["Budnummer"].faktisk := Value
        }
        get => this.parameter["Budnummer"].faktisk
    }
    
    ChfKonktaktNummerForventet{
        set{
            this.parameter["ChfKonktaktNummer"].forventet := Value
        }
        get => this.parameter["ChfKonktaktNummer"].forventet
    }
    ChfKonktaktNummerFaktisk{
        set{
            this.parameter["ChfKonktaktNummer"].faktisk := Value
        }
        get => this.parameter["ChfKonktaktNummer"].faktisk
    }
    
    HjemzoneForventet{
        set{
            this.parameter["Hjemzone"].forventet := Value
        }
        get => this.parameter["Hjemzone"].forventet
    }
    HjemzoneFaktisk{
        set{
            this.parameter["Hjemzone"].faktisk := Value
        }
        get => this.parameter["Hjemzone"].faktisk
    }
    
    KørerIkkeTransportTyperForventet{
        set{
            this.parameter["KørerIkkeTransportTyper"].forventet := Value
        }
        get => this.parameter["KørerIkkeTransportTyper"].forventet
    }
    KørerIkkeTransportTyperFaktisk{
        set{
            this.parameter["KørerIkkeTransportTyper"].faktisk := Value
        }
        get => this.parameter["KørerIkkeTransportTyper"].faktisk
    }
    
    PlanskemaForventet{
        set{
            this.parameter["Planskema"].forventet := Value
        }
        get => this.parameter["Planskema"].forventet
    }
    PlanskemaFaktisk{
        set{
            this.parameter["Planskema"].faktisk := Value
        }
        get => this.parameter["Planskema"].faktisk
    }
    
    SluttidForventet{
        set{
            this.parameter["Sluttid"].forventet := Value
        }
        get => this.parameter["Sluttid"].forventet
    }
    SluttidFaktisk{
        set{
            this.parameter["Sluttid"].faktisk := Value
        }
        get => this.parameter["Sluttid"].faktisk
    }
    
    StarttidForventet{
        set{
            this.parameter["Starttid"].forventet := Value
        }
        get => this.parameter["Starttid"].forventet
    }
    StarttidFaktisk{
        set{
            this.parameter["Starttid"].faktisk := Value
        }
        get => this.parameter["Starttid"].faktisk
    }
    
    StartzoneForventet{
        set{
            this.parameter["Startzone"].forventet := Value
        }
        get => this.parameter["Startzone"].forventet
    }
    StartzoneFaktisk{
        set{
            this.parameter["Startzone"].faktisk := Value
        }
        get => this.parameter["Startzone"].faktisk
    }
    
    StatistikGruppeForventet{
        set{
            this.parameter["StatistikGruppe"].forventet := Value
        }
        get => this.parameter["StatistikGruppe"].forventet
    }
    StatistikGruppeFaktisk{
        set{
            this.parameter["StatistikGruppe"].faktisk := Value
        }
        get => this.parameter["StatistikGruppe"].faktisk
    }
    
    UndtagneTransportTyperForventet{
        set{
            this.parameter["UndtagneTransportTyper"].forventet := Value
        }
        get => this.parameter["UndtagneTransportTyper"].forventet
    }
    UndtagneTransportTyperFaktisk{
        set{
            this.parameter["UndtagneTransportTyper"].faktisk := Value
        }
        get => this.parameter["UndtagneTransportTyper"].faktisk
    }
    
    VognløbskategoriForventet{
        set{
            this.parameter["Vognløbskategori"].forventet := Value
        }
        get => this.parameter["Vognløbskategori"].forventet
    }
    VognløbskategoriFaktisk{
        set{
            this.parameter["Vognløbskategori"].faktisk := Value
        }
        get => this.parameter["Vognløbskategori"].faktisk
    }
    
    VognmandKontaktnummerForventet{
        set{
            this.parameter["VognmandKontaktnummer"].forventet := Value
        }
        get => this.parameter["VognmandKontaktnummer"].forventet
    }
    VognmandKontaktnummerFaktisk{
        set{
            this.parameter["VognmandKontaktnummer"].faktisk := Value
        }
        get => this.parameter["VognmandKontaktnummer"].faktisk
    }
    
    VogmandLinie1Forventet{
        set{
            this.parameter["VogmandLinie1"].forventet := Value
        }
        get => this.parameter["VogmandLinie1"].forventet
    }
    VogmandLinie1Faktisk{
        set{
            this.parameter["VogmandLinie1"].faktisk := Value
        }
        get => this.parameter["VogmandLinie1"].faktisk
    }
    
    VognmandLinie2Forventet{
        set{
            this.parameter["VognmandLinie2"].forventet := Value
        }
        get => this.parameter["VognmandLinie2"].forventet
    }
    VognmandLinie2Faktisk{
        set{
            this.parameter["VognmandLinie2"].faktisk := Value
        }
        get => this.parameter["VognmandLinie2"].faktisk
    }
    
    VognmandLinie3Forventet{
        set{
            this.parameter["VognmandLinie3"].forventet := Value
        }
        get => this.parameter["VognmandLinie3"].forventet
    }
    VognmandLinie3Faktisk{
        set{
            this.parameter["VognmandLinie3"].faktisk := Value
        }
        get => this.parameter["VognmandLinie3"].faktisk
    }
    
    VognmandLinie4Forventet{
        set{
            this.parameter["VognmandLinie4"].forventet := Value
        }
        get => this.parameter["VognmandLinie4"].forventet
    }
    VognmandLinie4Faktisk{
        set{
            this.parameter["VognmandLinie4"].faktisk := Value
        }
        get => this.parameter["VognmandLinie4"].faktisk
    }
    
    ØkonomiskemaForventet{
        set{
            this.parameter["Økonomiskema"].forventet := Value
        }
        get => this.parameter["Økonomiskema"].forventet
    }
    ØkonomiskemaFaktisk{
        set{
            this.parameter["Økonomiskema"].faktisk := Value
        }
        get => this.parameter["Økonomiskema"].faktisk
    }
    
}
