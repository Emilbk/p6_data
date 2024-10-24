#Requires AutoHotkey v2.0

; Aktiverer P6-vindue, hvis ikke aktivt
/**
 */
class P6 extends class {

    vognløb := Array()

    setVognløb(pVognløb) {

        this.vognløb := pVognløb

        return
    }

    testVl() {

        MsgBox this.vognløb.tilIndlæsning.vognløbsnummer " - " this.vognløb.tilIndlæsning.vognløbsdato
        return
    }

    /** Metafunktioner */

    /** henter værdi fra P6-celle, eventuelt fra p6-msgbox hvis pHentMsgbox er sat
     * @param pKlipGenvej "appsKey" eller "ctrl", varierer fra felt til felt i P6
     * @param pHentMsgbox valgfri, hvis sat indhenter msgbox-besked 
     * @returns celleværdi eller msgbox-besked
     */
    kopierVærdi(pKlipGenvej, pHentMsgbox?, pNavigeringsSekvens?)
    {
        if (pKlipGenvej != "appsKey" and pKlipGenvej != "ctrl")
            throw Error("forkert genvejsinput")
        /** @var {Integer} clipwaitTid waittid ved første forsøg  */
        clipwaitTid := 0.4
        /** @var {Integer} clipwaitTidLoop waittid ved loop, når første mislykkes  */
        clipwaitTidLoop := 1.2
        clipwaitTidMsgbox := 0.5
        muligeKlipGenveje := Map("appsKey", "{appsKey}c", "ctrl", "^c")
        if (isset(pHentMsgbox) and pHentMsgbox != 0)
        {
            A_Clipboard := ""
            SendInput muligeKlipGenveje[pKlipGenvej]
            sleep 100
            clipwait clipwaitTidMsgbox
            sleep 100
            while a_clipboard = ""
            {
                if a_index > 1
                    return
                else
                {
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    sleep 200
                    ClipWait clipwaitTidMsgbox
                    sleep 300
                }
            }
            return a_clipboard
        }
        else
        {
            if IsSet(pNavigeringsSekvens)
            {
                Sendinput(pNavigeringsSekvens)
                sleep 20
            }
            A_Clipboard := ""
            SendInput muligeKlipGenveje[pKlipGenvej]
            sleep 100
            clipwait clipwaitTid
            sleep 100
            while a_clipboard = ""
            {
                if a_index > 10
                    throw (Error("Clipboardtimeout efter 10 forsøg"))
                else
                {
                    if IsSet(pNavigeringsSekvens)
                    {
                        Sendinput(pNavigeringsSekvens)
                        sleep 20
                    }
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    sleep 200
                    ClipWait clipwaitTidLoop
                    sleep 300
                }
            }
            return a_clipboard
        }
    }


    ;; P6-navigering
    navAktiverP6Vindue()
    {

        ; TODO bedre window-løsning? handle?
        HotIfWinNotActive "PLANET"
        {
            WinActivate "PLANET"
            WinWaitActive "PLANET"
            sleep 100
            SendInput "{esc}" ; registrerer ikke første tryk, når der skiftes til vindue
            ; sleep 300
            return true
        }
        return false
    }

    /**
     *  Aktiverer alt-menu i P6, tager op til to taste-sekvenser
     * @param tast1 
     * @param tast2 valgfri
     */
    navAltMenu(tast1, tast2?)
    {
        SendInput "{esc}{alt}"
        sleep 20
        Sendinput tast1
        if IsSet(tast2)
        {
            sleep 40
            SendInput tast2
            sleep 40
        }

        return
    }

    /** Lukker alle vinduer i P6 */
    navLukAlleVinduer()
    {
        SendInput "{esc}{alt}"
        sleep 20
        Sendinput "{v 2}{Down 2}{Enter}"
    }

    navLukAktivtVindue(){
        SendInput("^{F4}")
    }

    navVindueKørselsaftale()
    {
        ; this.navAktiverP6Vindue()
        this.navAltMenu("t", "k")
    }

    navVindueVognløb()
    {
        ; this.navAktiverP6Vindue()
        this.navAltMenu("t", "l")
        return
    }

    navVindueVognløbvognløbsnummer(){
        SendInput("!l")
    }

    navVindueVognløbvognløbsdato(){
        SendInput("!l{tab}")
    }

    ;; Data
    ændrVognløbsbilledeIndtastVognløbOgDato()
    {

        vognløbsnummer := this.vognløb.tilIndlæsning.Vognløbsnummer
        vognløbsdato := this.vognløb.tilIndlæsning.Vognløbsdato

        this.navAktiverP6Vindue()

        SendInput("^a")
        ; this.kopierVærdi("appsKey", 0, "!l")
        this.navVindueVognløbvognløbsnummer()
        SendInput(vognløbsnummer)
        ; this.kopierVærdi("ctrl", 0, "!l{tab}")
        this.navVindueVognløbvognløbsdato()
        SendInput(vognløbsdato)
        SendInput("{enter}")
        sleep 20

        this.kopierVærdi("ctrl", 1)
        if (InStr(A_Clipboard, "eksistere ikke"))
            throw Error("Vognløb ikke registreret - TODO")

        tjekAfIndtastningVognløbsnummer := this.kopierVærdi("appsKey", 0, "!l")
        tjekAfIndtastningVognløbsdato := this.kopierVærdi("ctrl", 0, "!l{tab}")

        if (tjekAfIndtastningVognløbsnummer != vognløbsnummer or tjekAfIndtastningVognløbsdato != vognløbsdato)
            throw (Error("Fejl i indtastning, vognløbsnummer eller dato er ikke korrekt"))

        return
    }

    ændrVognløbsbilledeÆndreVognløb()
    {
        SendInput ("^æ")
        sleep 20
        return
    }

    ændrVognløbsbilledeTjekKørselsaftaleOgStyresystem()
    {
        kørselsaftale := this.vognløb.tilIndlæsning.Kørselsaftale
        styresystem := this.vognløb.tilIndlæsning.Styresystem

        tjekEksisterendeKørselsaftale := this.kopierVærdi("appsKey", 0, "!k")
        tjekEksisterendeStyresystem := this.kopierVærdi("appsKey", 0, "!k{tab}")
        if (tjekEksisterendeKørselsaftale != kørselsaftale or tjekEksisterendeStyresystem != styresystem)
            throw (Error("Fejl i indlæsning, kørselsaftale eller styresystem er ikke det forventede"))

        SendInput("{enter}")
        this.kopierVærdi("ctrl", 1)
        if (InStr(A_Clipboard, "ikke registreret"))
            throw Error("Kørselsaftalen " kørselsaftale "_" styresystem " eksisterer ikke i P6.")

        return
    }

    ændrVognløbsbilledeIndtastÅbningstiderOgZone()
    {

        ; vognløbsdato := Format("{:U}", p_vl_obj["Dato"])
        vognløbsdato := this.vognløb.tilIndlæsning.Vognløbsdato
        starttid := this.vognløb.tilIndlæsning.Starttid
        sluttid := this.vognløb.tilIndlæsning.Sluttid
        startzone := this.vognløb.tilIndlæsning.Startzone
        slutzone := this.vognløb.tilIndlæsning.Slutzone
        hjemzone := this.vognløb.tilIndlæsning.Hjemzone

        ; this.navAktiverP6Vindue()
        this.kopierVærdi("ctrl")
        SendInput(vognløbsdato "{tab}")
        SendInput(starttid "{tab}")
        SendInput(vognløbsdato "{tab}")
        SendInput(sluttid "{tab}")
        SendInput(vognløbsdato "{tab}")
        SendInput(sluttid "{tab}")
        SendInput(startzone "{tab}")
        SendInput(slutzone "{tab}")
        SendInput(hjemzone "{tab}")
        SendInput("{enter}")
        p6_msgbox := this.kopierVærdi("ctrl", 1)
        if InStr(p6_msgbox, "Zone ikke registreret")
            throw (Error("Zonen findes ikke i P6"))
        if InStr(p6_msgbox, "Zone skal angives")
            throw (Error("Zonen er udfyldt tom"))

        return
    }

    ændrVognløbsbilledIndtastØvrige()
    {
        vognløbsnotering := this.vognløb.tilIndlæsning.Vognløbsnotering
        MobilnrChf := this.vognløb.tilIndlæsning.MobilnrChf
        Vognløbskategori := this.vognløb.tilIndlæsning.Vognløbskategori
        Planskema := this.vognløb.tilIndlæsning.Planskema
        Økonomiskema := this.vognløb.tilIndlæsning.Økonomiskema
        Statistikgruppe := this.vognløb.tilIndlæsning.Statistikgruppe
        UndtagneTransporttyper := this.vognløb.tilIndlæsning.UndtagneTransporttyper

        this.navAktiverP6Vindue()
        ; indlæsningstidstjek?
        if Vognløbsnotering
            SendInput("!p{tab 11}+{Up}" Vognløbsnotering)
        if MobilnrChf
            SendInput("!ø{tab 2}" MobilnrChf)
        if Vognløbskategori
            SendInput("!ø{tab 3}" Vognløbskategori)
        if Planskema
            SendInput("!ø{tab 6}" Planskema)
        if Økonomiskema
            SendInput("!ø{tab 8}" Økonomiskema)
        if Statistikgruppe
            SendInput("!ø{tab 9}" Statistikgruppe)

        return
    }

    ændrVognløbsbilledeIndtastTransporttyper() {

        UndtagneTransporttyper := this.vognløb.tilIndlæsning.UndtagneTransporttyper

        if UndtagneTransporttyper
        {
            SendInput("!u}")
            ; TODO #3 konsistens i antal slettede felter i transporttype
            sleep 20
            loop 19
            {
                SendInput("{delete}")
                sleep 25
                SendInput("{tab}")
                sleep 25

            }

            SendInput("{delete}")
            SendInput("!u}")
            for trtype in UndtagneTransporttyper
                SendInput("{tab}" trtype)
        }
    }

    ændrVognløbsbilledeAfslut()
    {
        SendInput("{enter}")

        p6_msgbox := this.kopierVærdi("ctrl", 1)
        if InStr(p6_msgbox, "Transporttypen")
            throw (Error("Transporttype findes ikke i P6"))
        if InStr(p6_msgbox, "Vløbsklasen")
            throw (Error("Vognløbskategorien findes ikke i P6"))
        ; kopierVærdi("shift")
        return
    }

    tjekVognløbsbiledeÅbningstiderogZone()
    {

        vognløbsdatoExcel := this.vognløb.tilIndlæsning.Vognløbsdato
        starttidExcel := this.vognløb.tilIndlæsning.Starttid
        sluttidExcel := this.vognløb.tilIndlæsning.Sluttid
        startzoneExcel := this.vognløb.tilIndlæsning.Startzone
        slutzoneExcel := this.vognløb.tilIndlæsning.Slutzone
        hjemzoneExcel := this.vognløb.tilIndlæsning.Hjemzone

        this.navAktiverP6Vindue()
        datoStartindlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        åbningsStartIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        datoNormaltSlutIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        åbningstidNormatlSlutIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        datoSidsteSlutIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        åbningstidSidsteSlutIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        StartzoneIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        SlutzoneIndlæst := this.kopierVærdi("ctrl")
        SendInput("{tab}")
        HjemzoneIndlæst := this.kopierVærdi("ctrl")
        SendInput("{enter}")

        ; TODO lav smartere tjek
        if datoStartindlæst != vognløbsdatoExcel
            throw Error("sdf")
        if åbningsStartIndlæst != starttidExcel
            throw Error("sdf")
        if datoNormaltSlutIndlæst != vognløbsdatoExcel
            throw Error("sdf")
        if åbningstidNormatlSlutIndlæst != sluttidExcel
            throw Error("sdf")
        if datoSidsteSlutIndlæst != vognløbsdatoExcel
            throw Error("sdf")
        if åbningstidSidsteSlutIndlæst != sluttidExcel
            throw Error("sdf")
        if StartzoneIndlæst != hjemzoneExcel
            throw Error("sdf")
        if SlutzoneIndlæst != hjemzoneExcel
            throw Error("sdf")
        if HjemzoneIndlæst != hjemzoneExcel
            throw Error("sdf")

        ; MsgBox "Alt i orden"

        return


    }

    funkÆndrVognløb()
    {
        this.navAktiverP6Vindue()
        ; this.navLukAlleVinduer()
        this.navVindueVognløb()
        this.ændrVognløbsbilledeIndtastVognløbOgDato()
        this.ændrVognløbsbilledeÆndreVognløb()
        this.ændrVognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.ændrVognløbsbilledeIndtastÅbningstiderOgZone()
        this.ændrVognløbsbilledIndtastØvrige()
        this.ændrVognløbsbilledeIndtastTransporttyper()
        this.ændrVognløbsbilledeAfslut()
        return
    }

    funkTjekVognløb()
    {
        this.navAktiverP6Vindue()
        this.navVindueVognløb()
        this.ændrVognløbsbilledeIndtastVognløbOgDato()
        this.ændrVognløbsbilledeÆndreVognløb()
        this.ændrVognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.tjekVognløbsbiledeÅbningstiderogZone()

    }
}

; comm
; TODO hvordan håndteres dato-array?
; p6_åben_vognløb()
; {
;     vognløbsnummer := this.["Vognløbsnummer"]
;     vognløbsdato := Format("{:U}", this.["Dato"])

;     P6_aktiver()
;     SendInput(vognløbsnummer)
;     SendInput "{tab}"
;     SendInput(vognløbsdato)
;     SendInput("{enter}")
;     sleep 20
;     kopierVærdi("ctrl", 1)
;     if (InStr(A_Clipboard, "eksistere ikke"))
;         throw Error("Ikke registreret - TODO")
;     ; tjek af korrekt vognløb
;     indlæst_vognløbsnummer := kopierVærdi("shift")
;     SendInput("{tab}")
;     indlæst_dato := kopierVærdi("ctrl")
;     ; lav fornuftigt system
;     if !(vognløbsnummer = indlæst_vognløbsnummer and vognløbsdato = indlæst_dato)
;         throw (Error("Fejl i indlæsning, ikke det forventede vognløbsnummer på forventet dato"))
;     return
; }
; ; TODO beslut hvordan hele vognløb-funktion skal struktureres
; p6_åben_vognløb_kørselsaftale(p_vl_obj)
; {
;     kørselsaftale := p_vl_obj["Kørselsaftale"]
;     styresystem := p_vl_obj["Styresystem"]

;     P6_aktiver()1
;     SendInput("^æ")
;     oprindelig_kørselsaftale := kopierVærdi("shift")
;     SendInput("{tab}")
;     oprindelig_styresystem := kopierVærdi("shift")
;     SendInput("{tab}")
;     if (kørselsaftale != oprindelig_kørselsaftale or styresystem != oprindelig_styresystem)
;         throw (Error("Fejl i indlæsning, åben kørselsaftale er ikke den forventede"))
;     SendInput(kørselsaftale)
;     SendInput "{tab}"
;     SendInput(styresystem)
;     SendInput("{tab}")
;     ; tjek af korrekt vognløb, omskriv
;     indlæst_kørselsaftale := kopierVærdi("shift")
;     SendInput("{tab}")
;     indlæst_styresystem := kopierVærdi("shift")
;     if (kørselsaftale = indlæst_kørselsaftale and styresystem = indlæst_styresystem)
;         korrekt := 1
;     SendInput("{enter}")
;     p6_msgbox := kopierVærdi("ctrl", 1)
;     if InStr(p6_msgbox, "ikke registreret")
;         throw (Error("Kørselsaftalen findes ikke i P6"))

;     return
; }

; ; lav modulær opbygning
; p6_åben_vognløb_åbningstider(p_vl_obj)
; {
;     vognløbsdato := Format("{:U}", p_vl_obj["Dato"])
;     starttid := p_vl_obj["Starttid"]
;     sluttid := p_vl_obj["Sluttid"]
;     startzone := p_vl_obj["Startzone"]
;     slutzone := p_vl_obj["Slutzone"]
;     hjemzone := p_vl_obj["Hjemzone"]

;     P6_aktiver()
;     kopierVærdi("ctrl")
;     SendInput(vognløbsdato "{tab}")
;     SendInput(starttid "{tab}")
;     SendInput(vognløbsdato "{tab}")
;     SendInput(sluttid "{tab}")
;     SendInput(vognløbsdato "{tab}")
;     SendInput(sluttid "{tab}")
;     SendInput(startzone "{tab}")
;     SendInput(slutzone "{tab}")
;     SendInput(hjemzone "{tab}")
;     SendInput("{enter}")
;     p6_msgbox := kopierVærdi("ctrl", 1)
;     if InStr(p6_msgbox, "Zone ikke registreret")
;         throw (Error("Zonen findes ikke i P6"))
;     if InStr(p6_msgbox, "Zone skal angives")
;         throw (Error("Zonen er udfyldt tom"))
;     return
; }
; ; modulær opbygning
; p6_åben_vognløb_resten(p_vl_obj)
; {

;     vognløbsnotering := p_vl_obj["Vognløbsnotering"]
;     MobilnrChf := p_vl_obj["MobilnrChf"]
;     Vognløbskategori := p_vl_obj["Vognløbskategori"]
;     Planskema := p_vl_obj["Planskema"]
;     Økonomiskema := p_vl_obj["Økonomiskema"]
;     Statistikgruppe := p_vl_obj["Statistikgruppe"]
;     UndtagneTransporttyper := p_vl_obj["Undtagne transporttyper"]

;     P6_aktiver()
;     if Vognløbsnotering
;         SendInput("!p{tab 11}+{Up}" Vognløbsnotering)
;     if MobilnrChf
;         SendInput("!ø{tab 2}" MobilnrChf)
;     if Vognløbskategori
;         SendInput("!ø{tab 3}" Vognløbskategori)
;     if Planskema
;         SendInput("!ø{tab 6}" Planskema)
;     if Økonomiskema
;         SendInput("!ø{tab 8}" Økonomiskema)
;     if Statistikgruppe
;         SendInput("!ø{tab 9}" Statistikgruppe)

;     if UndtagneTransporttyper
;     {
;         SendInput("!ø{tab 10}")
;         for trtype in UndtagneTransporttyper
;             SendInput("{tab}" trtype)
;     }
;     SendInput("{enter}")
;     return
; }
; p6_afslut_indlæsning_vognløb(p_vl_obj)
; {
;     ; P6_aktiver()
;     p6_msgbox := kopierVærdi("ctrl", 1)
;     if InStr(p6_msgbox, "Transporttypen")
;         throw (Error("Transporttype findes ikke i P6"))
;     if InStr(p6_msgbox, "Vløbsklasen")
;         throw (Error("Vognløbskategorien findes ikke i P6"))
;     ; kopierVærdi("shift")
;     return
; }

; p6_åben_kørselsaftale(p_vl_obj)
; {
;     P6_nav_kørselsaftale()
;     sleep 100
;     SendInput(p_vl_obj["Kørselsaftale"])
;     SendInput "{tab}"
;     SendInput(p_vl_obj["Styresystem"])
;     SendInput("{enter}")
;     p6_msgbox := kopierVærdi("ctrl", 1)
;     if (InStr(p6_msgbox, "ikke registreret"))
;         throw Error("Ikke registreret - TODO")
;     ; tjek af korrekt kørselsaftale
;     indlæst_kørselaftale := kopierVærdi("shift")
;     indlæst_kørselaftale := A_Clipboard
;     SendInput("{tab}")
;     indlæst_styresystem := kopierVærdi("shift")
;     SendInput("{tab}")
;     ; SendInput("^{F4}")
;     indlæst_styresystem := A_Clipboard
;     if (p_vl_obj["Kørselsaftale"] = indlæst_kørselaftale and p_vl_obj["Styresystem"] = indlæst_styresystem)
;     ; MsgBox "korrekt"
;         return
; }

; p6_indlæs_data_kørselsaftale_æ()
; {
;     SendInput("^æ")

;     return
; }

; p6_indlæs_data_kørselsaftale_planskema(p_vl_obj)
; {
;     SendInput("!p")
;     A_Clipboard := ""
;     tidligere_planskema := kopierVærdi("ctrl")
;     SendInput(p_vl_obj["Planskema"] "{tab}!p")
;     indlæst_planskema := kopierVærdi("ctrl")
;     if (p_vl_obj["Planskema"] = indlæst_planskema)
;         korrekt := 1
;     return
; }

; ; p6_indlæs_data_kørselsaftale_økonomiskema(p_vl_obj)
; ; {
; ;     SendInput("!p{tab 4}")
; ;     tidligere_planskema := kopierVærdi("ctrl")
; ;     SendInput(p_vl_obj["Planskema"] "{tab}!p{tab 4}")
; ;     tidligere_økonomiskema := kopierVærdi("ctrl")

; ;     if (p_vl_obj["Planskema"] = indlæst_planskema)
; ;         korrekt := 1
; ;     return
; ; }