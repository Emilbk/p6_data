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
    vognløb.tjekkedeParametre := parameterClass()
    vognløb.parametre := parameterClass()
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

        kørselsaftaleTilIndlæsning := this.vognløb.parametre.Kørselsaftale.forventetIndhold
        styresystemTilIndlæsning := this.vognløb.parametre.Styresystem.forventetIndhold
        kørselsaftaleTjek := this.kopierVærdi("ctrl")

        SendInput("{tab}")
        styresystemTjek := this.kopierVærdi("ctrl")

        if ((kørselsaftaleTjek != kørselsaftaleTilIndlæsning) or (styresystemTjek != styresystemTilIndlæsning))
            throw p6ForkertDataError("Forkert kørselsaftale indlæst", , , { forventetKørselsaftale: kørselsaftaleTilIndlæsning, faktiskKørselsaftale: kørselsaftaleTjek })
    }

    kørselsaftaleIndtastKørselsaftale() {

        kørselsaftaleTilIndlæsning := this.vognløb.parametre.Kørselsaftale.forventetIndhold
        styresystemTilIndlæsning := this.vognløb.parametre.Styresystem.forventetIndhold

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

        planskemaP := this.vognløb.parametre.planskema
        this.vognløb.parametre.setParameterEksisterende(planskemaP, planskema)
    }

    kørselsaftaleIndhentØkonomiskema() {
        SendInput("!p{tab 4}")
        Økonomiskema := this.kopierVærdi("ctrl", , , 1)

        ØkonomiskemaP := this.vognløb.parametre.Økonomiskema
        this.vognløb.parametre.setParameterEksisterende(ØkonomiskemaP, Økonomiskema)
    }

    kørselsaftaleIndhentStatistikgruppe() {
        SendInput("!p{tab 6}")
        Statistikgruppe := this.kopierVærdi("ctrl", , , 1)

        StatistikgruppeP := this.vognløb.parametre.Statistikgruppe
        this.vognløb.parametre.setParameterEksisterende(StatistikgruppeP, Statistikgruppe)
    }

    kørselsaftaleIndhentParameterVognmand() {
        SendInput("!p{tab 8}")
        ParameterVognmand := this.kopierVærdi("ctrl", , , 1)

        ParameterVognmandP := this.vognløb.parametre.ParameterVognmand
        this.vognløb.parametre.setParameterEksisterende(ParameterVognmandP, ParameterVognmand)
    }

    kørselsaftaleIndhentObligatoriskVognmand() {
        SendInput("!m+{tab 8}")
        obligatoriskVognmand := this.kopierVærdi("ctrl", , , 1)

        obligatoriskVognmandP := this.vognløb.parametre.obligatoriskVognmand
        this.vognløb.parametre.setParameterEksisterende(obligatoriskVognmandP, obligatoriskVognmand)
    }


    kørselsaftaleIndhentNormalHjemzone() {

        SendInput("!m+{tab 6}")
        normalHjemzone := this.kopierVærdi("ctrl", , , 1)

        normalHjemzoneP := this.vognløb.parametre.normalHjemzone
        this.vognløb.parametre.setParameterEksisterende(normalHjemzoneP, normalHjemzone)
    }

    kørselsaftaleIndhentKørerIkkeTransportTyper() {
        SendInput("!k")
        kørerIkkeTransportTyperOprRækkefølge := Array()
        kørerIkkeTransportTyperOprRækkefølgeP := this.vognløb.parametre.danParameterObj("kørerIkkeTransportTyperOprRækkefølge")
        this.vognløb.parametre.setParameterEksisterende(kørerIkkeTransportTyperOprRækkefølgeP, kørerIkkeTransportTyperOprRækkefølge)
        loop 10
        {
            kørerIkkeTransportTyperOprRækkefølge.push(this.kopierVærdi("ctrl", , , 1))
            SendInput("{tab}")
        }

    }

    kørselsaftaleIndhentPauseRegel() {
        SendInput("!r")
        pauseregel := this.kopierVærdi("ctrl", , , 1)

        pauseregelP := this.vognløb.parametre.pauseRegel
        this.vognløb.parametre.setParameterEksisterende(pauseregelP, pauseregel)
    }

    kørselsaftaleIndhentPauseDynamisk() {
        SendInput("!r{tab}")
        pauseDynamisk := this.kopierVærdi("ctrl", , , 1)

        pauseDynamiskP := this.vognløb.parametre.pauseDynamisk
        this.vognløb.parametre.setParameterEksisterende(pauseDynamiskP, pauseDynamisk)
    }

    kørselsaftaleIndhentPauseStart() {
        SendInput("!r{tab 3}")
        pauseStart := this.kopierVærdi("ctrl", , , 1)

        pauseStartP := this.vognløb.parametre.pauseStart
        this.vognløb.parametre.setParameterEksisterende(pauseStartP, pauseStart)
    }

    kørselsaftaleIndhentPauseSlut() {
        SendInput("!r{tab 4}")
        pauseSlut := this.kopierVærdi("ctrl", , , 1)

        pauseSlutP := this.vognløb.parametre.pauseSlut
        this.vognløb.parametre.setParameterEksisterende(pauseSlutP, pauseSlut)
    }

    kørselsaftaleIndhentVognmandNavn() {
        SendInput("!a")
        vognmandNavn := this.kopierVærdi("ctrl", , , 1)

        vognmandNavnP := this.vognløb.parametre.VognmandLinie1
        if Type(vognmandNavn) = "string" and vognmandNavn != ""
            this.vognløb.parametre.setParameterEksisterende(vognmandNavnP, vognmandNavn)

    }

    kørselsaftaleIndhentVognmandCO() {
        SendInput("!a{tab}")
        vognmandCO := this.kopierVærdi("ctrl", , , 1)

        vognmandCOP := this.vognløb.parametre.VognmandLinie2
        if Type(vognmandCO) = "string" and vognmandCO != ""
            this.vognløb.parametre.setParameterEksisterende(vognmandCOP, vognmandCO)
    }

    kørselsaftaleIndhentVognmandAdresse() {
        SendInput("!a{tab 2}")
        vognmandAdresse := this.kopierVærdi("ctrl", , , 1)

        vognmandAdresseP := this.vognløb.parametre.VognmandLinie3
        if Type(vognmandAdresse) = "string" and vognmandAdresse != ""
            this.vognløb.parametre.setParameterEksisterende(vognmandAdresseP, vognmandAdresse)
    }

    kørselsaftaleIndhentVognmandPostNr() {
        SendInput("!a{tab 3}")
        vognmandPostNr := this.kopierVærdi("ctrl", , , 1)

        vognmandPostNrP := this.vognløb.parametre.VognmandLinie4
        if Type(vognmandPostNr) = "string" and vognmandPostNr != ""
            this.vognløb.parametre.setParameterEksisterende(vognmandPostNrP, vognmandPostNr)
    }

    kørselsaftaleIndhentVognmandTelefon() {
        SendInput("!a{tab 4}")
        vognmandTelefon := this.kopierVærdi("ctrl", , , 1)

        vognmandTelefonP := this.vognløb.parametre.vognmandTelefon
        if Type(vognmandTelefon) = "string" and vognmandTelefon != ""
            this.vognløb.parametre.setParameterEksisterende(vognmandTelefonP, vognmandTelefon)
    }
    kørselsaftaleÆndr() {

        SendInput("^æ")
    }


    kørselsaftaleAfbryd() {
        SendInput("^a")
    }

    kørselsaftaleAfslut() {
        SendInput("^g")
    }

    kørselsaftaleIndtastPlansskema() {


        if !this.vognløb.parametre.planskema.forventetIndhold
            return
        planskema := this.vognløb.parametre.planskema.forventetIndhold

        SendInput("!p")
        SendInput(planskema)

    }
    kørselsaftaleIndtastØkonomiskema() {
        if !this.vognløb.parametre.økonomiskema.forventetIndhold
            return
        økonomiskema := this.vognløb.parametre.økonomiskema.forventetIndhold

        SendInput("!p {tab 4}")
        SendInput(økonomiskema)

    }

    kørselsaftaleIndtastStatistikgruppe() {
        if !this.vognløb.parametre.statistikgruppe.forventetIndhold
            return
        statistikgruppe := this.vognløb.parametre.statistikgruppe.forventetIndhold

        SendInput("!p {tab 6}")
        SendInput(statistikgruppe)

    }

    kørselsaftaleIndtastNormalHjemzone() {
        if !this.vognløb.parametre.Hjemzone.forventetIndhold
            return
        normalHjemzone := this.vognløb.parametre.Hjemzone.forventetIndhold
        SendInput("!m +{tab 6}")
        SendInput(normalHjemzone)

    }

    kørselsaftaleIndtastVognmandLinie1() {
        if !this.vognløb.parametre.vognmandLinie1.forventetIndhold
            return
        vognmandLinie1 := this.vognløb.parametre.vognmandLinie1.forventetIndhold

        SendInput("!a")
        SendInput(vognmandLinie1)

    }

    kørselsaftaleIndtastVognmandLinie2() {
        if !this.vognløb.parametre.vognmandLinie2.forventetIndhold
            return
        vognmandLinie2 := this.vognløb.parametre.vognmandLinie2.forventetIndhold

        SendInput("!a")
        sleep 20
        SendInput("{tab}")
        SendInput(vognmandLinie2)

    }
    kørselsaftaleIndtastVognmandLinie3() {
        if !this.vognløb.parametre.vognmandLinie3.forventetIndhold
            return
        vognmandLinie3 := this.vognløb.parametre.vognmandLinie3.forventetIndhold

        SendInput("!a")
        sleep 20
        SendInput("{tab 2}")
        SendInput(vognmandLinie3)

    }
    kørselsaftaleIndtastVognmandLinie4() {
        if !this.vognløb.parametre.vognmandLinie4.forventetIndhold
            return
        vognmandLinie4 := this.vognløb.parametre.vognmandLinie4.forventetIndhold

        SendInput("!a")
        sleep 20
        SendInput("{tab 3}")
        SendInput(vognmandLinie4)

    }
    kørselsaftaleIndtastVognmandKontaktnummer() {
        if !this.vognløb.parametre.vognmandKontaktnummer.forventetIndhold
            return
        vognmandKontaktnummer := this.vognløb.parametre.vognmandKontaktnummer.forventetIndhold

        SendInput("!a")
        sleep 20
        SendInput("{tab 4}")
        SendInput(vognmandKontaktnummer)

    }

    kørselsaftaleIndtastKørerIkkeTransporttyper() {
        if !this.vognløb.parametre.kørerIkkeTransporttyyper.forventetIndhold
            return
        kørerIkkeTransporttyyper := this.vognløb.parametre.kørerIkkeTransporttyyper.forventetIndhold

        SendInput("!p {tab 4}")
        for transporttype in kørerIkkeTransporttyyper
            SendInput(transporttype)

    }


    vognløbsbilledeIndtastVognløbOgDato()
    {

        vognløbsnummerTilindlæsning := this.vognløb.parametre.Vognløbsnummer.forventetIndhold
        vognløbsdatoTilIndlæsning := this.vognløb.parametre.Vognløbsdato.forventetIndhold

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
            throw P6MsgboxError("Vognløb findes ikke på dato " vognløbsdatoTilIndlæsning, , mBoxFejl, { test: "test" })
        if (InStr(mBoxFejl, "Planet Version 6"))
            throw P6MsgboxError("Ukendt Fejl", , mBoxFejl)

        ; TODO separat funk
        tjekAfIndtastningVognløbsnummer := this.kopierVærdi("appsKey", 0, "!l")
        tjekAfIndtastningVognløbsdato := this.kopierVærdi("ctrl", 0, "!l{tab}")

        if (tjekAfIndtastningVognløbsnummer != vognløbsnummerTilindlæsning or tjekAfIndtastningVognløbsdato != vognløbsdatoTilIndlæsning)
        {
            indtastningObj := { indtastetVognløbsnummer: tjekAfIndtastningVognløbsnummer, indtastetVognløbsdato: tjekAfIndtastningVognløbsdato, fejlType: "Indtastning af kørselsaftale" }
            throw (P6Indtastningsfejl("Fejl i indtastning, vognløbsnummer eller dato er ikke det forventede", , indtastningObj))
        
        }

        ; TODO omskriv når setparameter omskrevet
        this.vognløb.parametre.Vognløbsnummer.eksisterendeIndhold := tjekAfIndtastningVognløbsnummer
        this.vognløb.parametre.vognløbsdato.eksisterendeIndhold := tjekAfIndtastningVognløbsdato

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
        kørselsaftaleTilIndlæsning := this.vognløb.parametre.Kørselsaftale.forventetIndhold
        styresystemTilIndlæsning := this.vognløb.parametre.Styresystem.forventetIndhold

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

        this.vognløb.parametre.Kørselsaftale.eksisterendeIndhold := kørselsaftaleEksisterende
        this.vognløb.parametre.Styresystem.eksisterendeIndhold := styresystemEksisterende
    }

    vognløbsbilledeIndtastÅbningstiderOgZone()
    {

        ; vognløbsdato := Format("{:U}", p_vl_obj["Dato"])
        vognløbsdato := this.vognløb.parametre.Vognløbsdato.forventetIndhold
        vognløbsdatoSlut := this.vognløb.parametre.vognløbsdatoSlut.forventetIndhold
        starttid := this.vognløb.parametre.Starttid.forventetIndhold
        sluttid := this.vognløb.parametre.Sluttid.forventetIndhold
        startzone := this.vognløb.parametre.Startzone.forventetIndhold
        slutzone := this.vognløb.parametre.Slutzone.forventetIndhold
        hjemzone := this.vognløb.parametre.Hjemzone.forventetIndhold

        ; this.navAktiverP6Vindue()
        this.kopierVærdi("ctrl")
        if vognløbsdato
            SendInput(vognløbsdato)
        SendInput("{tab}")
        if starttid
            SendInput(starttid)
        SendInput("{tab}")
        if vognløbsdatoSlut
            SendInput(vognløbsdatoSlut)
        SendInput("{tab}")
        if sluttid
            SendInput(sluttid)
        SendInput("{tab}")
        if vognløbsdatoSlut
            SendInput(vognløbsdatoSlut)
        SendInput("{tab}")
        if sluttid
            SendInput(sluttid)
        SendInput("{tab}")
        if startzone
            SendInput(startzone)
        SendInput("{tab}")
        if slutzone
            SendInput(slutzone)
        SendInput("{tab}")
        if hjemzone
            SendInput(hjemzone)
        SendInput("{tab}")
        SendInput("{enter}")
        p6_msgbox := this.kopierVærdi("ctrl", 1)
        if InStr(p6_msgbox, "Zone ikke registreret")
            throw (P6MsgboxError("Zonen findes ikke i P6"))
        if InStr(p6_msgbox, "Zone skal angives")
            throw (P6MsgboxError("Zonen er udfyldt tom"))
        if (InStr(p6_msgbox, "for langt for modellen"))
            throw P6MsgboxError("Vognløbet er for langt for modellen", , p6_msgbox)

    }

    vognløbsbilledeIndtastØvrige()
    {
        vognløbsnotering := this.vognløb.parametre.Vognløbsnotering.forventetIndhold
        chfKontaktNummer := this.vognløb.parametre.chfKontaktNummer.forventetIndhold
        Vognløbskategori := this.vognløb.parametre.Vognløbskategori.forventetIndhold
        Planskema := this.vognløb.parametre.Planskema.forventetIndhold
        Økonomiskema := this.vognløb.parametre.Økonomiskema.forventetIndhold
        Statistikgruppe := this.vognløb.parametre.Statistikgruppe.forventetIndhold
        UndtagneTransporttyper := this.vognløb.parametre.UndtagneTransporttyper.forventetIndhold

        this.navAktiverP6Vindue()
        ; indlæsningstidstjek?
        if Vognløbsnotering
            SendInput("!p{tab 11}+{Up}" Vognløbsnotering)
        if chfKontaktNummer
            SendInput("!ø{tab 2}" chfKontaktNummer)
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

        if !this.vognløb.parametre.undtagneTransportTyper.iBrug
            return

        UndtagneTransporttyper := this.vognløb.parametre.UndtagneTransporttyper.forventetIndhold

        if UndtagneTransporttyper
        {
            SendInput("!u")

            for trtype in UndtagneTransporttyper
                SendInput(trtype "{tab}"), sleep(10)
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

        vognløbsdatoStartTilIndlæsning := this.vognløb.parametre.Vognløbsdato.forventetIndhold
        ; TODO lav tjek for slutdato over midnat i vognløbsconstructor
        ; nemmest at definere i excelark?
        vognløbsdatoSlutTilIndlæsning := this.vognløb.parametre.VognløbsdatoSlut.forventetIndhold
        starttidTilIndlæsning := this.vognløb.parametre.Starttid.forventetIndhold
        slutTidTilIndlæsning := this.vognløb.parametre.Sluttid.forventetIndhold
        startZoneTilIndlæsning := this.vognløb.parametre.Startzone.forventetIndhold
        slutzoneTilIndlæsning := this.vognløb.parametre.Slutzone.forventetIndhold
        hjemzoneTilIndlæsning := this.vognløb.parametre.Hjemzone.forventetIndhold

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
    ;; Vognløbsbillede Indhent

    vognløbsBilledeIndhentAlleÅbneVognløbsdatoer() {

        ;implementer
    }

    vognløbsbilledeIndhentÅbningstiderogZone() {

        if this.vognløb.gyldigeKolonner.startTid.iBrug
        {
            startDato := this.kopierVærdi("Ctrl")
            startDatoPar := this.vognløb.parametre.vognløbsdatoStart
            this.vognløb.parametre.setParameterEksisterende(startDatoPar, startDato)
            SendInput("{tab}")

            startTid := this.kopierVærdi("Ctrl")
            startTidPar := this.vognløb.parametre.startTid
            this.vognløb.parametre.setParameterEksisterende(startTidPar, startTid)
            SendInput("{tab}")

            normalSlutDato := this.kopierVærdi("Ctrl")
            normalSlutDatoPar := this.vognløb.parametre.VognløbsdatoNormalSlut
            this.vognløb.parametre.setParameterEksisterende(normalSlutDatoPar, normalSlutDato)
            SendInput("{tab}")

            normalSluttid := this.kopierVærdi("Ctrl")
            normalSluttidPar := this.vognløb.parametre.normalSluttid
            this.vognløb.parametre.setParameterEksisterende(normalSluttidPar, normalSluttid)
            SendInput("{tab}")

            sidsteSlutDato := this.kopierVærdi("Ctrl")
            sidsteSlutDatoPar := this.vognløb.parametre.VognløbsdatoSidsteSlut
            this.vognløb.parametre.setParameterEksisterende(sidsteSlutDatoPar, sidsteSlutDato)
            SendInput("{tab}")

            sidsteSlutTid := this.kopierVærdi("Ctrl")
            sidsteSlutTidPar := this.vognløb.parametre.sidsteSlutTid
            this.vognløb.parametre.setParameterEksisterende(sidsteSlutTidPar, sidsteSlutTid)
            SendInput("{tab}")
        }
        else
            SendInput("{tab 6}")

        if this.vognløb.gyldigeKolonner.hjemzone.iBrug
        {
            startzone := this.kopierVærdi("Ctrl")
            startzonePar := this.vognløb.parametre.startzone
            this.vognløb.parametre.setParameterEksisterende(startzonePar, startzone)
            SendInput("{tab}")

            slutzone := this.kopierVærdi("Ctrl")
            slutzonePar := this.vognløb.parametre.slutzone
            this.vognløb.parametre.setParameterEksisterende(slutzonePar, slutzone)
            SendInput("{tab}")

            hjemzone := this.kopierVærdi("Ctrl")
            hjemzonePar := this.vognløb.parametre.hjemzone
            this.vognløb.parametre.setParameterEksisterende(hjemzonePar, hjemzone)
            SendInput("{tab}")
        }
            SendInput("{Enter}")

            p6_msgbox := this.kopierVærdi("ctrl", 1)
            if InStr(p6_msgbox, "Zone ikke registreret")
                throw (P6MsgboxError("Zonen findes ikke i P6"))
            if InStr(p6_msgbox, "Zone skal angives")
                throw (P6MsgboxError("Zonen er udfyldt tom"))

        }

        vognløbsbilledeIndhentØvrige() {

            if this.vognløb.gyldigeKolonner.Vognløbsnotering.iBrug
            {
                SendInput("!v+{Up}")
                Vognløbsnotering := this.kopierVærdi("ctrl",,,1)
                VognløbsnoteringPar := this.vognløb.parametre.Vognløbsnotering
                this.vognløb.parametre.setParameterEksisterende(VognløbsnoteringPar, Vognløbsnotering)
            }
            if this.vognløb.gyldigeKolonner.chfKontaktNummer.iBrug
            {
                SendInput("!ø{tab 2}")
                chfKontaktNummer := this.kopierVærdi("appsKey")
                chfKontaktNummerPar := this.vognløb.parametre.chfKontaktNummer
                this.vognløb.parametre.setParameterEksisterende(chfKontaktNummerPar, chfKontaktNummer)
            }
            if this.vognløb.gyldigeKolonner.Vognløbskategori.iBrug
            {
                SendInput("!ø{tab 3}")
                Vognløbskategori := this.kopierVærdi("appsKey")
                VognløbskategoriPar := this.vognløb.parametre.Vognløbskategori
                this.vognløb.parametre.setParameterEksisterende(VognløbskategoriPar, Vognløbskategori)
            }
            if this.vognløb.gyldigeKolonner.Planskema.iBrug
            {
                SendInput("!ø{tab 6}")
                Planskema := this.kopierVærdi("appsKey")
                PlanskemaPar := this.vognløb.parametre.Planskema
                this.vognløb.parametre.setParameterEksisterende(PlanskemaPar, Planskema)
            }
            if this.vognløb.gyldigeKolonner.Økonomiskema.iBrug
            {
                SendInput("!ø{tab 8}")
                Økonomiskema := this.kopierVærdi("appsKey")
                ØkonomiskemaPar := this.vognløb.parametre.Økonomiskema
                this.vognløb.parametre.setParameterEksisterende(ØkonomiskemaPar, Økonomiskema)
            }
            if this.vognløb.gyldigeKolonner.Statistikgruppe.iBrug
            {
                SendInput("!ø{tab 9}")
                Statistikgruppe := this.kopierVærdi("appsKey")
                StatistikgruppePar := this.vognløb.parametre.Statistikgruppe
                this.vognløb.parametre.setParameterEksisterende(StatistikgruppePar, Statistikgruppe)
            }
        }

        vognløbsbilledeIndhentTransporttyper() {
            if this.vognløb.gyldigeKolonner.undtagneTransportTyper.iBrug
            {

                undtagneTransportTyper := Array()
                transportTyperPar := this.vognløb.parametre.undtagneTransportTyper
                this.vognløb.parametre.undtagneTransportTyper.eksisterendeIndhold := undtagneTransportTyper

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
        }

        kørselsaftaleIndhent() {

        }


        funkKørselsaftaleÆndrHjemzone() {

            this.navAktiverP6Vindue()
            this.navLukAlleVinduer()
            this.navVindueKørselsaftale()
            this.kørselsaftaleIndtastKørselsaftale()
            this.kørselsaftaleTjekKørselsaftaleOgStyresystem()
            this.kørselsaftaleÆndr()
            this.kørselsaftaleIndtastNormalHjemzone()
            this.kørselsaftaleIndtastVognmandLinie1()
            this.kørselsaftaleIndtastVognmandLinie2()
            this.kørselsaftaleIndtastVognmandLinie3()
            this.kørselsaftaleIndtastVognmandLinie4()
            this.kørselsaftaleAfslut()
        }
        funkVognløbsbilledeÆndrHjemzone() {

            this.navVindueVognløb()
            this.vognløbsbilledeIndtastVognløbOgDato()
            this.vognløbsbilledeÆndrVognløb()
            this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
            this.vognløbsbilledeIndtastÅbningstiderOgZone()
            this.ændrVognløbsbilledeAfslut()
        }
        funkÆndrVognløb()
        {
            this.navAktiverP6Vindue()
            this.navLukAlleVinduer()
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

        funkIndhentData() {

            this.funkIndhentVognløbsbillede()
        }

        funkIndhentVognløbsbillede() {

            ; this.navAktiverP6Vindue()
            ; this.navLukAlleVinduer()
            this.navVindueVognløb()
            this.vognløbsbilledeIndtastVognløbOgDato()
            this.vognløbsbilledeÆndrVognløb()
            this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
            this.vognløbsbilledeIndhentÅbningstiderogZone()
            this.vognløbsbilledeIndhentØvrige()
            this.vognløbsbilledeIndhentTransporttyper()
            this.ændrVognløbsbilledeAfslut()
        }
        funkIndhentKørselsaftale() {

            this.navAktiverP6Vindue()
            this.navLukAlleVinduer()
            this.navVindueKørselsaftale()
            this.kørselsaftaleIndtastKørselsaftale()
            this.kørselsaftaleTjekKørselsaftaleOgStyresystem()
            this.kørselsaftaleÆndr()
            this.kørselsaftaleIndhentPlanskema()
            this.kørselsaftaleIndhentØkonomiskema()
            this.kørselsaftaleIndhentStatistikgruppe()
            this.kørselsaftaleIndhentNormalHjemzone()
            this.kørselsaftaleIndhentKørerIkkeTransportTyper()
            this.kørselsaftaleIndhentObligatoriskVognmand()
            this.kørselsaftaleIndhentPauseRegel()
            this.kørselsaftaleIndhentPauseDynamisk()
            this.kørselsaftaleIndhentPauseStart()
            this.kørselsaftaleIndhentPauseSlut()
            this.kørselsaftaleIndhentVognmandNavn()
            this.kørselsaftaleIndhentVognmandCO()
            this.kørselsaftaleIndhentVognmandAdresse()
            this.kørselsaftaleIndhentVognmandPostNr()
            this.kørselsaftaleIndhentVognmandTelefon()
        }

    }

    class parameterClass {

        Budnummer := { navn: "Budnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Vognløbsnummer := { navn: "Vognløbsnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Vognløbsdato := { navn: "Vognløbsdato", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognløbsdatoStart := { navn: "VognløbsdatoStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognløbsdatoSlut := { navn: "VognløbsdatoSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognløbsdatoNormalSlut := { navn: "VognløbsdatoNormalslut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognløbsdatoSidsteSlut := { navn: "VognløbsdatoSidsteSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Kørselsaftale := { navn: "Kørselsaftale", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Styresystem := { navn: "Styresystem", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Starttid := { navn: "Starttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Sluttid := { navn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        NormalSluttid := { navn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        SidsteSluttid := { navn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Hjemzone := { navn: "Hjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Startzone := { navn: "Startzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Slutzone := { navn: "Slutzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        NormalHjemzone := { navn: "NormalHjemzone", forventetIndhold: this.Hjemzone.forventetIndhold, eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        chfKontaktNummer := { navn: "chfKontaktNummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Vognløbskategori := { navn: "Vognløbskategori", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Planskema := { navn: "Planskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Økonomiskema := { navn: "Økonomiskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Statistikgruppe := { navn: "Statistikgruppe", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Vognløbsnotering := { navn: "Vognløbsnotering", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognmandLinie1 := { navn: "VognmandLinie1", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognmandLinie2 := { navn: "VognmandLinie2", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognmandLinie3 := { navn: "VognmandLinie3", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognmandLinie4 := { navn: "VognmandLinie4", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        VognmandKontaktnummer := { navn: "VognmandKontaktnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        ObligatoriskVognmand := { navn: "ObligatoriskVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        KørselsaftaleVognmand := { navn: "KørselsaftaleVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0 }
        Ugedage := { navn: "Ugedage", forventetIndhold: Array(), eksisterendeIndhold: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0 }
        UndtagneTransporttyper := { navn: "UndtagneTransporttyper", forventetIndhold: Array(), eksisterendeIndhold: Array(), ForventetMenIkkeIEksisterende: Array(), EksisterendeMenIkkeIForventet: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0 }
        KørerIkkeTransporttyper := { navn: "KørerIkkeTransporttyper", forventetIndhold: Array(), eksisterendeIndhold: Array(), ForventetMenIkkeIEksisterende: Array(), EksisterendeMenIkkeIForventet: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0 }


        ; kun i P6
        pauseRegel := { navn: "pauseRegel", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
        pauseDynamisk := { navn: "pauseDynamisk", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
        pauseStart := { navn: "pauseStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
        pauseSlut := { navn: "pauseSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }
        kørerIkkeTransportTyperOprindeligRækkefølge := { navn: "kørerIkkeTransportTyperOprindeligRækkefølge", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }


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

            if !forventetParameterIndhold
                return

            if fundetFejl
                return

            if forventetParameterIndhold != eksisterendeParameterIndhold
                pParamaetObj.fejl := 1
            else
                pParamaetObj.fejl := 0
        }

        tjekAlleParameterForFejl() {

            undtagetArray := Map("UndtagneTransporttyper", 0, "KørerIkkeTransporttyper", 0, "Ugedage", 0)
            for parameternavn, parameterobj in this.OwnProps()
                if !undtagetArray.Has(parameternavn)
                    this.tjekParameterForFejl(parameterobj)

        }
        skabOgTestParameter(pParameterNavn, pForventetIndhold, pEksisterendeIndhold) {


            parameter := this.danParameterObj(pParameterNavn)
            this.setParameterForventet(parameter, pForventetIndhold)
            this.setParameterEksisterende(parameter, pEksisterendeIndhold)
            this.tjekParameterForFejl(parameter)
        }

        sorterArrayAlfabetisk(pArrayTilSortering) {
            sorteretStr := ""

            for arrayIndhold in pArrayTilSortering
            {
                if arrayIndhold != A_Space
                {
                    arrayIndhold := StrUpper(arrayIndhold)
                    sorteretStr .= arrayIndhold ","
                }
            }

            sorteretStr := SubStr(sorteretStr, 1, -1)
            sorteretStr := Sort(sorteretStr, "d,")
            sorteretArray := StrSplit(sorteretStr, ",")

            return sorteretArray
        }

        sorterUndtagneTransporttyperForventet() {


            sorteretArray := this.sorterArrayAlfabetisk(this.UndtagneTransporttyper.ForventetIndhold)

            this.UndtagneTransporttyper.ForventetIndhold := sorteretArray

            return
        }
        sorterUndtagneTransporttyperEksisterende() {


            sorteretArray := this.sorterArrayAlfabetisk(this.UndtagneTransporttyper.eksisterendeIndhold)

            this.UndtagneTransporttyper.eksisterendeIndhold := sorteretArray

            return
        }
        sorterKørerIkkeTransporttyperEksisterende() {


            sorteretArray := this.sorterArrayAlfabetisk(this.KørerIkkeTransporttyper.eksisterendeIndhold)

            this.KørerIkkeTransporttyper.eksisterendeIndhold := sorteretArray

            return
        }
        sorterKørerIkkeTransporttyperForventet() {


            sorteretArray := this.sorterArrayAlfabetisk(this.KørerIkkeTransporttyper.ForventetIndhold)

            this.KørerIkkeTransporttyper.ForventetIndhold := sorteretArray

            return
        }

        arrayTilMap(pArrayTilMap) {

            MapFraArray := Map()

            for value in pArrayTilMap
                MapFraArray.Set(value, "")


            return MapFraArray
        }

        tjekUndtagneTransportTyperEns() {

            tjekTransporttyperEksisterende := this.UndtagneTransporttyper.eksisterendeIndhold
            tjekTransporttyperForventet := this.UndtagneTransporttyper.ForventetIndhold
            tjekTransporttyperEksisterendeMenIkkeForventet := Array()
            tjekTransporttyperForventetMenIkkeEksisterende := Array()

            MapTransportTyperEksisterende := this.arrayTilMap(tjekTransporttyperEksisterende)
            MapTransportTyperForventet := this.arrayTilMap(tjekTransporttyperForventet)

            for mapName, mapValue in MapTransportTyperEksisterende
                if !MapTransportTyperForventet.has(mapName)
                    tjekTransporttyperEksisterendeMenIkkeForventet.Push(mapName)

            for mapName, mapValue in MapTransportTyperForventet
                if !MapTransportTyperEksisterende.Has(mapname)
                    tjekTransporttyperForventetMenIkkeEksisterende.Push(mapname)

            if tjekTransporttyperEksisterendeMenIkkeForventet.Length or tjekTransporttyperForventet.Length
                this.UndtagneTransportTyper.fejl := 1

            this.UndtagneTransporttyper.EksisterendeMenIkkeIForventet := tjekTransporttyperEksisterendeMenIkkeForventet
            this.UndtagneTransporttyper.ForventetMenIkkeIEksisterende := tjekTransporttyperForventetMenIkkeEksisterende
        }
        tjekKørerIkkeTransportTyperEns() {

            tjekTransporttyperEksisterende := this.KørerIkkeTransporttyper.eksisterendeIndhold
            tjekTransporttyperForventet := this.KørerIkkeTransporttyper.ForventetIndhold
            tjekTransporttyperEksisterendeMenIkkeForventet := Array()
            tjekTransporttyperForventetMenIkkeEksisterende := Array()

            MapTransportTyperEksisterende := this.arrayTilMap(tjekTransporttyperEksisterende)
            MapTransportTyperForventet := this.arrayTilMap(tjekTransporttyperForventet)

            for mapName, mapValue in MapTransportTyperEksisterende
                if !MapTransportTyperForventet.has(mapName)
                    tjekTransporttyperEksisterendeMenIkkeForventet.Push(mapName)

            for mapName, mapValue in MapTransportTyperForventet
                if !MapTransportTyperEksisterende.Has(mapname)
                    tjekTransporttyperForventetMenIkkeEksisterende.Push(mapname)

            if tjekTransporttyperEksisterendeMenIkkeForventet.Length or tjekTransporttyperForventet.Length
                this.KørerIkkeTransportTyper.fejl := 1

            this.KørerIkkeTransporttyper.EksisterendeMenIkkeIForventet := tjekTransporttyperEksisterendeMenIkkeForventet
            this.KørerIkkeTransporttyper.ForventetMenIkkeIEksisterende := tjekTransporttyperForventetMenIkkeEksisterende
        }
    }

    class p6Mock extends P6 {

        vognløb := Object()
        tjekkedeParametre := p6ParameterMock()
        vognløb.tjekkedeParametre := this.tjekkedeParametre


        vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        {
            kørselsaftaleTilIndlæsning := this.vognløb.parametre.Kørselsaftale.forventetIndhold
            styresystemTilIndlæsning := this.vognløb.parametre.Styresystem.forventetIndhold

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

    class p6ParameterMock extends parameterClass {


    }