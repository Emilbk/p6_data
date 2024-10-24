#Requires AutoHotkey v2.0

; Aktiverer P6-vindue, hvis ikke aktivt
/**
 */
class P6 extends class {

    vognløb := Map()

    /** Metafunktioner */

    /** henter værdi fra P6-celle, eventuelt fra p6-msgbox hvis pHentMsgbox er sat
     * @param pKlipGenvej "appsKey" eller "ctrl", varierer fra felt til felt i P6
     * @param pHentMsgbox valgfri, hvis sat indhenter msgbox-besked 
     * @returns celleværdi eller msgbox-besked
     */
    kopierVærdi(pKlipGenvej, pHentMsgbox?)
    {
        if (pKlipGenvej != "appsKey" and pKlipGenvej != "ctrl")
            throw Error("forkert genvejsinput")
        /** @var {Integer} clipwaitTid waittid ved første forsøg  */
        clipwaitTid := 0.4
        /** @var {Integer} clipwaitTidLoop waittid ved loop, når første mislykkes  */
        clipwaitTidLoop := 0.5
        clipwaitTidMsgbox := 0.5
        muligeKlipGenveje := Map("appsKey", "{appsKey}c", "ctrl", "^c")
        if IsSet(pHentMsgbox)
        {
            A_Clipboard := ""
            SendInput muligeKlipGenveje[pKlipGenvej]
            clipwait clipwaitTidMsgbox
            sleep 20
            while A_Clipboard = ""
            {
                if a_index > 1
                    return
                else
                {
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    ClipWait clipwaitTidMsgbox
                }
            }
            return A_Clipboard
        }
        if !IsSet(pHentMsgbox)
        {
            A_Clipboard := ""
            SendInput muligeKlipGenveje[pKlipGenvej]
            clipwait clipwaitTid
            while A_Clipboard = ""
            {
                if a_index > 10
                    throw (Error("Clipboardtimeout efter 10 forsøg"))
                else
                {
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    ClipWait clipwaitTidLoop
                }
            }
            return A_Clipboard
        }
    }

    ;; VL-data
    dataIndhentVlObj(pVlMap)
    {
        this.vognløb := pVlMap
        return
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
        return
    }

    navVindueVognløb()
    {
        ; this.navAktiverP6Vindue()
        this.navAltMenu("t", "l")
        return
    }

    ;; Data
    ændrVognløbsbilledeIndtastVognløbOgDato()
    {

        vognløbsnummer := this.vognløb["Vognløbsnummer"]
        vognløbsdato := this.vognløb["Vognløbsdato"]

        this.navAktiverP6Vindue()

        SendInput("^a")
        SendInput(vognløbsnummer)
        SendInput "{tab}"
        SendInput(vognløbsdato)
        SendInput("{enter}")
        sleep 20

        this.kopierVærdi("ctrl", 1)
        if (InStr(A_Clipboard, "eksistere ikke"))
            throw Error("Vognløb ikke registreret - TODO")

        tjekAfIndtastningVognløbsnummer := this.kopierVærdi("appsKey")
        SendInput("{tab}")
        tjekAfIndtastningVognløbsdato := this.kopierVærdi("ctrl")

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
        kørselsaftale := this.vognløb["Kørselsaftale"]
        styresystem := this.vognløb["Styresystem"]

        tjekEksisterendeKørselsaftale := this.kopierVærdi("appsKey")
        SendInput("{tab}")
        tjekEksisterendeStyresystem := this.kopierVærdi("appsKey")
        if (tjekEksisterendeKørselsaftale != this.vognløb["Kørselsaftale"] or tjekEksisterendeStyresystem != this.vognløb["Styresystem"])
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
        vognløbsdato := this.vognløb["Vognløbsdato"]
        starttid := this.vognløb["Starttid"]
        sluttid := this.vognløb["Sluttid"]
        startzone := this.vognløb["Startzone"]
        slutzone := this.vognløb["Slutzone"]
        hjemzone := this.vognløb["Hjemzone"]

        this.navAktiverP6Vindue()
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
        vognløbsnotering := this.vognløb["Vognløbsnotering"]
        MobilnrChf := this.vognløb["MobilnrChf"]
        Vognløbskategori := this.vognløb["Vognløbskategori"]
        Planskema := this.vognløb["Planskema"]
        Økonomiskema := this.vognløb["Økonomiskema"]
        Statistikgruppe := this.vognløb["Statistikgruppe"]
        UndtagneTransporttyper := this.vognløb["Undtagne transporttyper"]

        this.navAktiverP6Vindue()
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

        if UndtagneTransporttyper
        {
            SendInput("!u}")
            sleep 20
            loop 20
            {
                SendInput("{delete}")
                sleep 10
                SendInput("{tab}")

            }

            SendInput("!u}")
            for trtype in UndtagneTransporttyper
                SendInput("{tab}" trtype)
        }
        return
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
        
        vognløbsdatoExcel := this.vognløb["Vognløbsdato"]
        starttidExcel := this.vognløb["Starttid"]
        sluttidExcel := this.vognløb["Sluttid"]
        startzoneExcel := this.vognløb["Startzone"]
        slutzoneExcel := this.vognløb["Slutzone"]
        hjemzoneExcel := this.vognløb["Hjemzone"]

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
        this.navVindueVognløb()
        this.ændrVognløbsbilledeIndtastVognløbOgDato()
        this.ændrVognløbsbilledeÆndreVognløb()
        this.ændrVognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.ændrVognløbsbilledeIndtastÅbningstiderOgZone()
        this.ændrVognløbsbilledIndtastØvrige()
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
