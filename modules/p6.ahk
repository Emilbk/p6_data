/**
 * 
 */
#Requires AutoHotkey v2.0
; TODO opdel i navigering, datatjek og databehandling?


; Aktiverer P6-vindue, hvis ikke aktivt
/**
 */
class P6 extends class {

    vognløb := Object()
    ; dannes undervejs i parametertjek
    vognløb.tjekkedeParametre := p6Parameter()
    vognløb.indhentedeParametre := p6Parameter()
    vindueHandle := ""


    setP6Vindue(pvinduehandle?) {
        if IsSet(pvinduehandle)
            vindueHandle := pvinduehandle
        else {
            if InStr(WinGetTitle("A"), "PLANET")
                vindueHandle := WinExist("PLANET")

        }

        this.vindueHandle := vindueHandle
        return this.vindueHandle
    }

    /**
     * 
     * @param pVognløb @type {VognløbObj}
     */
    setVognløb(pVognløb) {

        if !Type(pVognløb) = VognløbObj
            throw TypeError("Er ikke vl-obj")
        this.vognløb := pVognløb

        return
    }


    /** Metafunktioner */

    ; clipboard-bug fikset af 2.0.18?
    /** henter værdi fra P6-celle, eventuelt fra p6-msgbox hvis pHentMsgbox er sat
     * @param pKlipGenvej "appsKey" eller "ctrl", varierer fra felt til felt i P6
     * @param pHentMsgbox valgfri, hvis sat indhenter msgbox-besked 
     * @returns celleværdi eller msgbox-besked
     */
    kopierVærdi(pKlipGenvej, pHentMsgbox?, pNavigeringsSekvens?, pVentIkkePåClipboard?)
    {
        pKlipGenvej := StrLower(pKlipGenvej)
        if (pKlipGenvej != "appskey" and pKlipGenvej != "ctrl")
            throw Error("forkert genvejsinput")
        /** @var {Integer} clipwaitTid waittid ved første forsøg  */
        clipwaitTid := 0.4
        /** @var {Integer} clipwaitTidLoop waittid ved loop, når første mislykkes  */
        clipwaitTidLoop := 1.2
        clipwaitTidMsgbox := 0.5
        muligeKlipGenveje := Map("appskey", "{appsKey}c", "ctrl", "^c")
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
            if IsSet(pVentIkkePåClipboard)
            {
                A_Clipboard := ""
                SendInput muligeKlipGenveje[pKlipGenvej]
                sleep 100
                clipwait clipwaitTid
                sleep 100
                while a_clipboard = ""
                {
                    if a_index > 2
                        return
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
                SendInput muligeKlipGenveje[pKlipGenvej]
                sleep 100
                clipwait clipwaitTid
                sleep 100

                return A_Clipboard
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

    enterOgTjekForMsgboxFejl() {
        SendInput("{Enter}")
        msgBoxFejl := this.kopierVærdi("ctrl", 1)

        return msgBoxFejl
    }

    ;; P6-navigering
    /**
     * Aktiver p6-vindue
     * @returns {Integer} 
     */
    navAktiverP6Vindue()
    {

        ; TODO bedre window-løsning? handle?
        if !WinActive("PLANET")
        {
            WinActivate("ahk_id" this.vindueHandle)
            WinWaitSuccess := WinWaitActive(this.vindueHandle, , 3)

            ; hvordan håndteres timeout?
            ; if !WinWaitSuccess
            ; {
            ;     sendinput("{escape}")
            ;     winpid := WinGetPID("A")
            ;     WinActivate("ahk_pid" winpid, "PLANET",)
            ;     WinWaitActive("ahk_pid" winpid, , , "PLANET")
            ;     msgbox := this.kopierVærdi("ctrl", 1)
            ;     if InStr(msgbox, "inaktiv i mere end")
            ;         throw P6Msgbox("P6-session timeout")
            ; }
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

    /**
     * Lukker alle vinduer i P6
     */
    navLukAlleVinduer()
    {
        SendInput "{esc}{alt}"
        sleep 20
        Sendinput "{v 2}{Down 2}{Enter}"
    }

    navLukAktivtVindue() {
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

    navVindueVognløbvognløbsnummer() {
        SendInput("!l")
    }

    navVindueVognløbvognløbsdato() {
        SendInput("!l{tab}")
    }

    ;; Data
    kørselsaftaleTjekKørselsaftaleOgStyresystem() {

        kørselsaftaleTilIndlæsning := this.vognløb.tilIndlæsning.Kørselsaftale
        styresystemTilIndlæsning := this.vognløb.tilIndlæsning.Styresystem
        kørselsaftaleTjek := this.kopierVærdi("ctrl")

        SendInput("{tab}")
        styresystemTjek := this.kopierVærdi("ctrl")

        if ((kørselsaftaleTjek != kørselsaftaleTilIndlæsning) or (styresystemTjek != styresystemTilIndlæsning))
            throw p6ForkertDataError("Forkert kørselsaftale indlæst", , , { forventetKørselsaftale: kørselsaftaleTilIndlæsning, faktiskKørselsaftale: kørselsaftaleTjek })
    }

    kørselsaftaleIndtastKørselsaftale() {

        kørselsaftaleTilIndlæsning := this.vognløb.tilIndlæsning.Kørselsaftale
        styresystemTilIndlæsning := this.vognløb.tilIndlæsning.Styresystem

        SendInput(kørselsaftaleTilIndlæsning "{tab}" styresystemTilIndlæsning)
        SendInput("{Enter}")

        mBoxtjek := this.kopierVærdi("ctrl", 1)
        if InStr(mBoxtjek, "registreret")
        {
            SendInput("{Enter}")
            throw P6MsgboxError("Kørselsaftalen findes ikke i P6")
        }
    }

    kørselsaftaleIndhentPlanskema() {
        SendInput("!p")
        planskema := this.kopierVærdi("ctrl", , , 1)

        planskemaP  := this.vognløb.indhentedeParametre.planskema
        this.vognløb.indhentedeParametre.setParameterEksisterende(planskemaP, planskema)
    }

    kørselsaftaleIndhentØkonomiskema() {
        SendInput("!p{tab 4}")
        Økonomiskema := this.kopierVærdi("ctrl", , , 1)

        ØkonomiskemaP  := this.vognløb.indhentedeParametre.Økonomiskema
        this.vognløb.indhentedeParametre.setParameterEksisterende(ØkonomiskemaP, Økonomiskema)
    }

    kørselsaftaleIndhentStatistikgruppe() {
        SendInput("!p{tab 6}")
        Statistikgruppe := this.kopierVærdi("ctrl", , , 1)

        StatistikgruppeP  := this.vognløb.indhentedeParametre.Statistikgruppe
        this.vognløb.indhentedeParametre.setParameterEksisterende(StatistikgruppeP, Statistikgruppe)
    }

    kørselsaftaleIndhentParameterVognmand() {
        SendInput("!p{tab 8}")
        ParameterVognmand := this.kopierVærdi("ctrl", , , 1)

        ParameterVognmandP  := this.vognløb.indhentedeParametre.ParameterVognmand
        this.vognløb.indhentedeParametre.setParameterEksisterende(ParameterVognmandP, ParameterVognmand)
    }

    kørselsaftaleIndhentObligatoriskVognmand() {
        SendInput("!m+{tab 8}")
        obligatoriskVognmand := this.kopierVærdi("ctrl", , , 1)

        obligatoriskVognmandP  := this.vognløb.indhentedeParametre.obligatoriskVognmand
        this.vognløb.indhentedeParametre.setParameterEksisterende(obligatoriskVognmandP, obligatoriskVognmand)
    }


    kørselsaftaleIndhentNormalHjemzone() {

        SendInput("!m+{tab 6}")
        normalHjemzone := this.kopierVærdi("ctrl", , , 1)

        normalHjemzoneP  := this.vognløb.indhentedeParametre.normalHjemzone
        this.vognløb.indhentedeParametre.setParameterEksisterende(normalHjemzoneP, normalHjemzone)
    }

    kørselsaftaleIndhentKørerIkkeTransportTyper() {
        SendInput("!k")
        kørerIkkeTransportTyperOprRækkefølge := Array()
        kørerIkkeTransportTyperOprRækkefølgeP := this.vognløb.indhentedeParametre.danParameterObj("kørerIkkeTransportTyperOprRækkefølge")
        this.vognløb.indhentedeParametre.setParameterEksisterende(kørerIkkeTransportTyperOprRækkefølgeP, kørerIkkeTransportTyperOprRækkefølge)
        loop 10
        {
            kørerIkkeTransportTyperOprRækkefølge.push(this.kopierVærdi("ctrl", , , 1))
            SendInput("{tab}")
        }

    }

    kørselsaftaleIndhentPauseRegel() {
        SendInput("!r")
        pauseregel := this.kopierVærdi("ctrl", , , 1)

        pauseregelP := this.vognløb.indhentedeParametre.kørerIkkeTransportTyperOprRækkefølgeP
        this.vognløb.indhentedeParametre.setParameterEksisterende(pauseregelP, pauseregel)
    }

    kørselsaftaleIndhentPauseDynamisk() {
        SendInput("!r{tab}")
        pauseDynamisk := this.kopierVærdi("ctrl", , , 1)

        pauseDynamiskP  := this.vognløb.indhentedeParametre.pauseDynamisk
        this.vognløb.indhentedeParametre.setParameterEksisterende(pauseDynamiskP, pauseDynamisk)
    }

    kørselsaftaleIndhentPauseStart() {
        SendInput("!r{tab 3}")
        pauseStart := this.kopierVærdi("ctrl", , , 1)

        pauseStartP  := this.vognløb.indhentedeParametre.pauseStart
        this.vognløb.indhentedeParametre.setParameterEksisterende(pauseStartP, pauseStart)
    }

    kørselsaftaleIndhentPauseSlut() {
        SendInput("!r{tab 4}")
        pauseSlut := this.kopierVærdi("ctrl", , , 1)

        pauseSlutP  := this.vognløb.indhentedeParametre.pauseSlut
        this.vognløb.indhentedeParametre.setParameterEksisterende(pauseSlutP, pauseSlut)
    }

    kørselsaftaleIndhentVognmandNavn() {
        SendInput("!a")
        vognmandNavn := this.kopierVærdi("ctrl", , , 1)

        vognmandNavnP  := this.vognløb.indhentedeParametre.vognmandNavn
        if Type(vognmandNavn) = "string" and vognmandNavn != ""
            this.vognløb.indhentedeParametre.setParameterEksisterende(vognmandNavnP, vognmandNavn)

    }

    kørselsaftaleIndhentVognmandCO() {
        SendInput("!a{tab}")
        vognmandCO := this.kopierVærdi("ctrl", , , 1)

        vognmandCOP  := this.vognløb.indhentedeParametre.vognmandCO
        if Type(vognmandCO) = "string" and vognmandCO != ""
            this.vognløb.indhentedeParametre.setParameterEksisterende(vognmandCOP, vognmandCO)
    }

    kørselsaftaleIndhentVognmandAdresse() {
        SendInput("!a{tab 2}")
        vognmandAdresse := this.kopierVærdi("ctrl", , , 1)

        vognmandAdresseP  := this.vognløb.indhentedeParametre.vognmandAdresse
        if Type(vognmandAdresse) = "string" and vognmandAdresse != ""
            this.vognløb.indhentedeParametre.setParameterEksisterende(vognmandAdresseP, vognmandAdresse)
    }

    kørselsaftaleIndhentVognmandPostNr() {
        SendInput("!a{tab 3}")
        vognmandPostNr := this.kopierVærdi("ctrl", , , 1)

        vognmandPostNrP  := this.vognløb.indhentedeParametre.vognmandPostNr
        if Type(vognmandPostNr) = "string" and vognmandPostNr != ""
            this.vognløb.indhentedeParametre.setParameterEksisterende(vognmandPostNrP, vognmandPostNr)
    }

    kørselsaftaleIndhentVognmandTelefon() {
        SendInput("!a{tab 4}")
        vognmandTelefon := this.kopierVærdi("ctrl", , , 1)

        vognmandTelefonP  := this.vognløb.indhentedeParametre.vognmandTelefon
        if Type(vognmandTelefon) = "string" and vognmandTelefon != ""
            this.vognløb.indhentedeParametre.setParameterEksisterende(vognmandTelefonP, vognmandTelefon)
    }
    kørselsaftaleÆndr() {

        SendInput("^æ")
    }


    kørselsaftaleAfbryd() {
        SendInput("^a")
    }

    kørselsaftaleIndtastPlansskemaOgØkonomiskema() {
        ;planskema !p
        ;økonomi !p{tab 4}

    }

    kørselsaftaleIndtastStatistikgruppe() {
        ;stat !p{tab 6}
    }

    kørselsaftaleIndtastNormalHjemzone() {
        ;normHjemzone !m{tab 6}
    }
    kørselsaftaleIndtastVognmandNavn() {
        ;vmnavn !a
    }

    kørselsaftaleIndtastVognmanCO() {
        ;vmCo !a{tab}
    }

    kørselsaftaleIndtastHjemzoneAdresse() {
        ;vmAdr !a{tab 2}
    }

    kørselsaftaleIndtastHjemzonePostnr() {

        ;  !a{tab 3}
    }

    kørselsaftaleIndtastVMKontaktnummer() {
        ; !a{tab 4}
    }


    kørselsaftaleIndtastKørerIkkeTransporttyper() {
        ;!k
    }


    vognløbsbilledeIndtastVognløbOgDato()
    {

        vognløbsnummerTilindlæsning := this.vognløb.tilIndlæsning.Vognløbsnummer
        vognløbsdatoTilIndlæsning := this.vognløb.tilIndlæsning.Vognløbsdato

        this.navAktiverP6Vindue()

        SendInput("^a")
        this.navVindueVognløbvognløbsnummer()
        SendInput(vognløbsnummerTilindlæsning)
        ; this.kopierVærdi("ctrl", 0, "!l{tab}")
        this.navVindueVognløbvognløbsdato()
        SendInput(vognløbsdatoTilIndlæsning)
        SendInput("{enter}")
        sleep 20

        mBoxFejl := this.kopierVærdi("ctrl", 1)
        if (InStr(mBoxFejl, "eksistere ikke"))
            throw P6MsgboxError("Vognløb findes ikke på dato " vognløbsdatoTilIndlæsning, , mBoxFejl)

        ; TODO separat funk
        tjekAfIndtastningVognløbsnummer := this.kopierVærdi("appsKey", 0, "!l")
        tjekAfIndtastningVognløbsdato := this.kopierVærdi("ctrl", 0, "!l{tab}")

        if (tjekAfIndtastningVognløbsnummer != vognløbsnummerTilindlæsning or tjekAfIndtastningVognløbsdato != vognløbsdatoTilIndlæsning)
        {
            indtastningObj := { indtastetVognløbsnummer: tjekAfIndtastningVognløbsnummer, indtastetVognløbsdato: tjekAfIndtastningVognløbsdato, fejlType: "Indtastning af kørselsaftale" }
            throw (P6Indtastningsfejl("Fejl i indtastning, vognløbsnummer eller dato er ikke det forventede", , indtastningObj))
        }
        return
    }

    vognløbsbilledeÆndrVognløb()
    {
        SendInput ("^æ")
        sleep 20
        return
    }

    vognløbsbilledeAfbryd()
    {
        SendInput ("^a")
        sleep 20
        return
    }
    vognløbsbilledeTjekKørselsaftaleOgStyresystem()
    {
        kørselsaftaleTilIndlæsning := this.vognløb.tilIndlæsning.Kørselsaftale
        styresystemTilIndlæsning := this.vognløb.tilIndlæsning.Styresystem

        kørselsaftaleEksisterende := this.kopierVærdi("appsKey", 0, "!k")
        styresystemEksisterende := this.kopierVærdi("appsKey", 0, "!k{tab}")

        if kørselsaftaleEksisterende != kørselsaftaleTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", kørselsaftaleEksisterende, styresystemEksisterende, "kørselsaftale")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    fejlIParameter: "kørselsaftaleVognløbsbillede" })
        if styresystemEksisterende != styresystemTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", styresystemTilIndlæsning, styresystemEksisterende, "styresystem")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    FejlIParameter: "KørselsaftaleVognløbsbillede" })

        mBoxFejl := this.enterOgTjekForMsgboxFejl()
        if (InStr(mBoxFejl, "ikke registreret"))
            throw P6MsgboxError("Kørselsaftalen eksisterer ikke i P6.", , mBoxFejl,)

    }

    vognløbsbilledeIndtastÅbningstiderOgZone()
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
            throw (P6MsgboxError("Zonen findes ikke i P6"))
        if InStr(p6_msgbox, "Zone skal angives")
            throw (P6MsgboxError("Zonen er udfyldt tom"))

    }

    vognløbsbilledeIndtastØvrige()
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

    vognløbsbilledeIndtastTransporttyper() {

        UndtagneTransporttyper := this.vognløb.tilIndlæsning.UndtagneTransporttyper

        if UndtagneTransporttyper
        {
            SendInput("!u")
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
            sleep 400
            SendInput("!u}")
            for trtype in UndtagneTransporttyper
                SendInput(trtype "{tab}")
        }
    }

    ændrVognløbsbilledeAfslut()
    {
        SendInput("{enter}")

        p6_msgbox := this.kopierVærdi("ctrl", 1)
        if InStr(p6_msgbox, "Transporttypen")
            throw (P6MsgboxError("Transporttype findes ikke i P6"))
        if InStr(p6_msgbox, "Vløbsklasen")
            throw (P6MsgboxError("Vognløbskategorien findes ikke i P6"))
        ; kopierVærdi("shift")
        return
    }


    vognløbsbilledeTjekÅbningstiderOgZone()
    {

        vognløbsdatoStartTilIndlæsning := this.vognløb.tilIndlæsning.Vognløbsdato
        ; TODO lav tjek for slutdato over midnat i vognløbsconstructor
        ; nemmest at definere i excelark?
        vognløbsdatoSlutTilIndlæsning := this.vognløb.tilIndlæsning.Vognløbsdato
        starttidTilIndlæsning := this.vognløb.tilIndlæsning.Starttid
        slutTidTilIndlæsning := this.vognløb.tilIndlæsning.Sluttid
        startZoneTilIndlæsning := this.vognløb.tilIndlæsning.Startzone
        slutzoneTilIndlæsning := this.vognløb.tilIndlæsning.Slutzone
        hjemzoneTilIndlæsning := this.vognløb.tilIndlæsning.Hjemzone

        ; TODO START HER
        this.navAktiverP6Vindue()
        if vognløbsdatoStartTilIndlæsning
        {
            startdato := this.kopierVærdi("ctrl")
            this.vognløb.tjekkedeParametre.skabOgTestParameter("Vognløbsdato", vognløbsdatoStartTilIndlæsning, startdato)
            SendInput("{tab}")
        }
        else
            SendInput("{tab}")

        if starttidTilIndlæsning
        {
            startTid := this.kopierVærdi("ctrl")
            this.vognløb.tjekkedeParametre.skabOgTestParameter(starttidTilIndlæsning)
            startTid := this.tjekParameter(starttidTilIndlæsning, "Starttid", "ctrl")
        }
        else
            SendInput("{tab}")

        if vognløbsdatoSlutTilIndlæsning
            normalSlutDato := this.tjekParameter(vognløbsdatoSlutTilIndlæsning, "Sluttid", "ctrl")
        else
            SendInput("{tab}")

        if slutTidTilIndlæsning
            normalSlutTid := this.tjekParameter(slutTidTilIndlæsning, "Sluttid", "ctrl")
        else
            SendInput("{tab}")

        if vognløbsdatoSlutTilIndlæsning
            sidsteSlutDato := this.tjekParameter(vognløbsdatoSlutTilIndlæsning, "Vognløbsdato", "ctrl")
        else
            SendInput("{tab}")

        if slutTidTilIndlæsning
            sidsteSlutTid := this.tjekParameter(slutTidTilIndlæsning, "Sluttid", "ctrl")
        else
            SendInput("{tab}")

        if startZoneTilIndlæsning
            startZone := this.tjekParameter(startZoneTilIndlæsning, "Startzone", "ctrl")
        else
            SendInput("{tab}")

        if slutzoneTilIndlæsning
            slutzone := this.tjekParameter(slutzoneTilIndlæsning, "Slutzone", "ctrl")
        else
            SendInput("{tab}")
        if hjemzoneTilIndlæsning
            hjemzone := this.tjekParameter(hjemzoneTilIndlæsning, "Hjemzone", "ctrl")

        SendInput("{enter}")
        return


    }

    vognløbsbilledeIndhentÅbningstiderogZone() {

        tjekForAktivtVindue := this.kopierVærdi("ctrl")
        startDato := this.kopierVærdi("Ctrl")
        startDatoPar := this.vognløb.indhentedeParametre.startDato
        this.vognløb.indhentedeParametre.setParameterEksisterende(startDatoPar, startDato)

        startTid := this.kopierVærdi("Ctrl")
        startTidPar := this.vognløb.indhentedeParametre.startTid
        this.vognløb.tjekkedeParametre.setParameterEksisterende(startTidPar, startTid)
        SendInput("{tab}")

        normalSlutDato := this.kopierVærdi("Ctrl")
        normalSlutDatoPar := this.vognløb.indhentedeParametre.normalSlutDato
        this.vognløb.indhentedeParametre.setParameterEksisterende(normalSlutDatoPar, normalSlutDato)
        SendInput("{tab}")

        normalSluttid := this.kopierVærdi("Ctrl")
        normalSluttidPar := this.vognløb.indhentedeParametre.normalSluttid
        this.vognløb.indhentedeParametre.setParameterEksisterende(normalSluttidPar, normalSluttid)
        SendInput("{tab}")

        sidsteSlutDato := this.kopierVærdi("Ctrl")
        sidsteSlutDatoPar := this.vognløb.indhentedeParametre.sidsteSlutDato
        this.vognløb.indhentedeParametre.setParameterEksisterende(sidsteSlutDatoPar, sidsteSlutDato)
        SendInput("{tab}")

        sidsteSlutTid := this.kopierVærdi("Ctrl")
        sidsteSlutTidPar := this.vognløb.indhentedeParametre.sidsteSlutTid
        this.vognløb.indhentedeParametre.setParameterEksisterende(sidsteSlutTidPar, sidsteSlutTid)
        SendInput("{tab}")

        startzone := this.kopierVærdi("Ctrl")
        startzonePar := this.vognløb.indhentedeParametre.startzone
        this.vognløb.indhentedeParametre.setParameterEksisterende(startzonePar, startzone)
        SendInput("{tab}")

        slutzone := this.kopierVærdi("Ctrl")
        slutzonePar := this.vognløb.indhentedeParametre.slutzone
        this.vognløb.indhentedeParametre.setParameterEksisterende(slutzonePar, slutzone)
        SendInput("{tab}")

        hjemzone := this.kopierVærdi("Ctrl")
        hjemzonePar := this.vognløb.indhentedeParametre.hjemzone
        this.vognløb.indhentedeParametre.setParameterEksisterende(hjemzonePar, hjemzone)
        SendInput("{enter}")

        p6_msgbox := this.kopierVærdi("ctrl", 1)
        if InStr(p6_msgbox, "Zone ikke registreret")
            throw (P6MsgboxError("Zonen findes ikke i P6"))
        if InStr(p6_msgbox, "Zone skal angives")
            throw (P6MsgboxError("Zonen er udfyldt tom"))

    }

    vognløbsbilledeIndhentØvrige() {
        SendInput("!v+{Up}")
        Vognløbsnotering := this.kopierVærdi("ctrl")
        VognløbsnoteringPar := this.vognløb.indhentedeParametre.Vognløbsnotering
        this.vognløb.indhentedeParametre.setParameterEksisterende(VognløbsnoteringPar, Vognløbsnotering)


        SendInput("!ø{tab 2}")
        MobilnrChf := this.kopierVærdi("appsKey")
        MobilnrChfPar := this.vognløb.indhentedeParametre.MobilnrChf
        this.vognløb.indhentedeParametre.setParameterEksisterende(MobilnrChfPar, MobilnrChf)

        SendInput("!ø{tab 3}")
        Vognløbskategori := this.kopierVærdi("appsKey")
        VognløbskategoriPar := this.vognløb.indhentedeParametre.Vognløbskategori
        this.vognløb.indhentedeParametre.setParameterEksisterende(VognløbskategoriPar, Vognløbskategori)

        SendInput("!ø{tab 6}")
        Planskema := this.kopierVærdi("appsKey")
        PlanskemaPar := this.vognløb.indhentedeParametre.Planskema
        this.vognløb.indhentedeParametre.setParameterEksisterende(PlanskemaPar, Planskema)

        SendInput("!ø{tab 8}")
        Økonomiskema := this.kopierVærdi("appsKey")
        ØkonomiskemaPar := this.vognløb.indhentedeParametre.Økonomiskema
        this.vognløb.indhentedeParametre.setParameterEksisterende(ØkonomiskemaPar, Økonomiskema)

        SendInput("!ø{tab 9}")
        Statistikgruppe := this.kopierVærdi("appsKey")
        StatistikgruppePar := this.vognløb.indhentedeParametre.Statistikgruppe
        this.vognløb.indhentedeParametre.setParameterEksisterende(StatistikgruppePar, Statistikgruppe)
    }

    vognløbsbilledeIndhentTransporttyper() {


        undtagneTransportTyper := Array()
        transportTyperPar := this.vognløb.indhentedeParametre.transportTyper
        this.vognløb.indhentedeParametre.transporttyper.eksisterendeIndhold := undtagneTransportTyper
        SendInput("!u")
        loop 19
        {
            tType := this.kopierVærdi("ctrl", , , 1)
            if !tType
                break
            undtagneTransportTyper.Push(tType)
            SendInput("{tab}")

        }
    }

    kørselsaftaleIndhent() {

    }
    funkÆndrVognløb()
    {
        this.navAktiverP6Vindue()
        ; this.navLukAlleVinduer()
        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeIndtastÅbningstiderOgZone()
        this.vognløbsbilledeIndtastØvrige()
        this.vognløbsbilledeIndtastTransporttyper()
        this.ændrVognløbsbilledeAfslut()
        return
    }

    funkTjekVognløb()
    {
        this.navAktiverP6Vindue()
        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeTjekÅbningstiderOgZone()

    }
}

class p6Parameter {

    vognløbDatoStart := { navn: "vognløbDatoStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbDatoSlut := { navn: "vognløbDatoSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbTidStart := { navn: "vognløbTidStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbTidSlut := { navn: "vognløbTidSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    startzone := { navn: "startzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    slutzone := { navn: "slutzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    hjemzone := { navn: "hjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbsNotering := { navn: "vognløbsNotering", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbsKategori := { navn: "vognløbsKategori", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    undtagneTransportTyper := { navn: "undtagneTransportTyper", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    mobilNrChf := { navn: "mobilNrChf", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognmandNavn := { navn: "vognmandNavn", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognmandCO := { navn: "vognmandCO", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognmandAdresse := { navn: "vognmandAdresse", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognmandPostNr := { navn: "vognmandPostNr", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognmandTelefon := { navn: "vognmandTelefon", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    pauseRegel := { navn: "pauseRegel", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    pauseDynamisk := { navn: "pauseDynamisk", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    pauseStart := { navn: "pauseStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    pauseSlut := { navn: "pauseSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    kørerIkkeTransportTyperOprindeligRækkefølge := { navn: "kørerIkkeTransportTyperOprindeligRækkefølge", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    normalHjemzone := { navn: "normalHjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    parameterVognmand := { navn: "parameterVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    obligatoriskVognmand := { navn: "obligatoriskVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    statistikgruppe := { navn: "statistikgruppe", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    økonomiskema := { navn: "økonomiskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    planskema := { navn: "planskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    kørselsaftale := { navn: "kørselsaftale", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    styresystem := { navn: "styresystem", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbsDato := { navn: "vognløbsDato", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
    vognløbsNummer := { navn: "vognløbsNummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }

    danParameterObj(pParameterNavn) {

        if this.HasOwnProp(pParameterNavn)
            return

        this.%pParameterNavn% := { navn: pParameterNavn, forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }

        return this.%pParameterNavn%
    }

    setParameterForventet(pParameterObj, pForventet) {


        pParameterObj.forventetIndhold := pForventet


    }
    setParameterEksisterende(pParameterObj, pEksisterende) {


        pParameterObj.eksisterendeIndhold := pEksisterende


    }

    /**
     * 
     * @param pParamaetObj 
     * @returns Bool
     */
    tjekParameterForFejl(pParamaetObj) {


        forventetParameterIndhold := pParamaetObj.forventetIndhold
        eksisterendeParameterIndhold := pParamaetObj.eksisterendeIndhold
        fundetFejl := pParamaetObj.fejl

        if fundetFejl
            return

        if forventetParameterIndhold != eksisterendeParameterIndhold
            pParamaetObj.fejl := 1
        else
            pParamaetObj.fejl := 0
    }

    skabOgTestParameter(pParameterNavn, pForventetIndhold, pEksisterendeIndhold) {


        parameter := this.danParameterObj(pParameterNavn)
        this.setParameterForventet(parameter, pForventetIndhold)
        this.setParameterEksisterende(parameter, pEksisterendeIndhold)
        this.tjekParameterForFejl(parameter)
    }
}


class p6Mock extends P6 {

    vognløb := Object()
    tjekkedeParametre := p6ParameterMock()
    vognløb.tjekkedeParametre := this.tjekkedeParametre


    vognløbsbilledeTjekKørselsaftaleOgStyresystem()
    {
        kørselsaftaleTilIndlæsning := this.vognløb.tilIndlæsning.Kørselsaftale
        styresystemTilIndlæsning := this.vognløb.tilIndlæsning.Styresystem

        kørselsaftaleEksisterende := this.kopierVærdi("appsKey", 0, "!k")
        styresystemEksisterende := this.kopierVærdi("appsKey", 0, "!k{tab}")

        if kørselsaftaleEksisterende != kørselsaftaleTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", kørselsaftaleEksisterende, styresystemEksisterende, "kørselsaftale")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    fejlIParameter: "kørselsaftaleVognløbsbillede" })
        if styresystemEksisterende != styresystemTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", styresystemTilIndlæsning, styresystemEksisterende, "styresystem")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    FejlIParameter: "KørselsaftaleVognløbsbillede" })

        mBoxFejl := this.enterOgTjekForMsgboxFejl()
        if (InStr(mBoxFejl, "ikke registreret"))
            throw P6MsgboxError("Kørselsaftalen eksisterer ikke i P6.", , mBoxFejl,)

        return
    }
    funkTjekVognløb()
    {
        ; this.navAktiverP6Vindue()
        ; this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        ; this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeTjekÅbningstiderOgZone()

    }
}

class p6ParameterMock extends p6Parameter {


}