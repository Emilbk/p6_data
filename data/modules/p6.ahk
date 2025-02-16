/**
 * 
 */
#Requires AutoHotkey v2.0
; TODO opdel i navigering, datatjek og databehandling?

; TODO tilpas ny datastruktur (kolonnenavn og parameternavn)

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

    tryClipwait(pWaitTid) {

        try {

            clipwait pWaitTid
        } catch Error as e {

            ; filpath := A_ScriptDir "\clipWaitfejl" FormatTime(, "ddMM-HHmmss") ".txt"
            ; sleep 100
            ; FileAppend(Format("-----`nLine: {1}`nMessage: {2}`nWhat: {3}`nStack: {4}`n------", e.Line, e.Message, e.What, e.Stack))
            ClipWait pWaitTid

        }
    }

    tryClipRead() {
        try {
            clip := A_Clipboard
            return clip

        } catch Error as e {
            loop 20 {
                try {
                    sleep 40
                    clip := A_Clipboard
                    if clip != ""
                        return clip
                } catch Error as e {
                    continue
                }
            }
            MsgBox "Fatal clipRead-fejl"
        }
    }
    tryClipWrite(str) {
        try {
            A_Clipboard := str
        } catch Error as e {
            loop 20 {
                try {
                    sleep 40
                    A_Clipboard := str
                    clip := this.tryClipRead()
                    if clip := str
                        return
                } catch Error as e {
                    continue
                }
            }
            MsgBox "Fatal clipWrite-fejl"
        }
    }

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
        clipwaitTid := 0.2
        /** @var {Integer} clipwaitTidLoop waittid ved loop, når første mislykkes  */
        clipwaitTidLoop := 0.2
        clipwaitTidMsgbox := 0.2
        muligeKlipGenveje := Map("appskey", "{appsKey}c", "ctrl", "^c")
        if (isset(pHentMsgbox) and pHentMsgbox != 0)
        {
            this.tryClipWrite("")
            SendInput muligeKlipGenveje[pKlipGenvej]
            this.tryClipwait(clipwaitTid)
            clip := this.tryClipRead()
            tidSatIgang := A_Now
            sekunderGået := 0
            while clip = ""
            {
                sekunderGået := DateDiff(A_Now, tidSatIgang, "Seconds")
                if sekunderGået > 1
                    return
                else
                {
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    sleep 20
                    this.tryClipwait clipwaitTidMsgbox
                    clip := this.tryClipRead()
                }
            }
            return clip
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
                this.tryClipWrite("")
                SendInput muligeKlipGenveje[pKlipGenvej]
                sleep 20
                this.tryClipwait clipwaitTid
                clip := this.tryClipRead()
                tidSatIgang := A_Now
                while clip = ""
                {
                    sekunderGået := DateDiff(A_Now, tidSatIgang, "Seconds")
                    if sekunderGået > 1
                        return
                    else
                    {
                        if IsSet(pNavigeringsSekvens)
                        {
                            Sendinput(pNavigeringsSekvens)
                            sleep 20
                        }
                        SendInput muligeKlipGenveje[pKlipGenvej]
                        sleep 20
                        this.tryClipwait clipwaitTidLoop
                        clip := this.tryClipRead()
                    }
                }
                return clip
            }
            this.tryClipWrite("")
            SendInput muligeKlipGenveje[pKlipGenvej]
            sleep 40
            this.tryClipwait(clipwaitTid)
            clip := this.tryClipRead()
            sleep 40
            tidSatIgang := A_Now
            while clip = ""
            {
                sleep 40
                sekunderGået := DateDiff(A_Now, tidSatIgang, "Seconds")
                if sekunderGået > 10
                    throw (Error("Clipboard-timeout efter 10 sekunder, er P6 aktiveret i korrekt felt og er Citrix-udkllipsholderen tilgængelig? "))
                else
                {
                    if IsSet(pNavigeringsSekvens)
                    {
                        Sendinput(pNavigeringsSekvens)
                        sleep 20
                    }
                    SendInput muligeKlipGenveje[pKlipGenvej]
                    sleep 40
                    this.tryClipwait clipwaitTidLoop
                    clip := this.tryClipRead()
                }
            }
            return clip
        }
    }

    enterOgHentMsgboxFejl() {
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
            try {
                WinActivate("ahk_id" this.vindueHandle)
                WinWaitSuccess := WinWaitActive(this.vindueHandle, , 3)

            } catch Error as e {
                MsgBox "P6-vindue er ikke valgt!"
            }

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
            throw P6MsgboxError("Kørselsaftalen " kørselsaftaleTilIndlæsning "_" styresystemTilIndlæsning " findes ikke i P6")
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
        sleep 50

        mBoxFejl := this.kopierVærdi("ctrl", 1)
        this.tjekP6Msgbox(mBoxFejl)

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

    vognløbsbilledeIndtastVognløbOgDatoTilIndlæg()
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
        SendInput("^i")
        sleep 300

        mBoxFejl := this.kopierVærdi("ctrl", 1)
        this.tjekP6Msgbox(mBoxFejl)

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
    vognløbsbilledeIndtastKørselsaftaleOgStyresystemTilIndlæg()
    {
        kørselsaftaleTilIndlæsning := this.vognløb.parametre.Kørselsaftale.forventetIndhold
        styresystemTilIndlæsning := this.vognløb.parametre.Styresystem.forventetIndhold

        SendInput("!k" kørselsaftaleTilIndlæsning)
        SendInput("!k{tab}" styresystemTilIndlæsning)

        kørselsaftaleEksisterende := this.kopierVærdi("appsKey", 0, "!k")
        styresystemEksisterende := this.kopierVærdi("appsKey", 0, "!k{tab}")

        if kørselsaftaleEksisterende != kørselsaftaleTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", kørselsaftaleTilIndlæsning, kørselsaftaleEksisterende, "kørselsaftale")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    fejlIParameter: "kørselsaftaleVognløbsbillede" })
        if styresystemEksisterende != styresystemTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}. Forventet {3}: {1}, Eksisterende {3}: {2}", styresystemTilIndlæsning, styresystemEksisterende, "styresystem")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    FejlIParameter: "KørselsaftaleVognløbsbillede" })


        mBoxFejl := this.enterOgHentMsgboxFejl()
        this.tjekP6Msgbox(mBoxFejl)

        this.vognløb.parametre.Kørselsaftale.eksisterendeIndhold := kørselsaftaleEksisterende
        this.vognløb.parametre.Styresystem.eksisterendeIndhold := styresystemEksisterende
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
                Format("Fejl i indlæsning af {3}`nForventet {3}: {1}`nEksisterende {3}: {2}", kørselsaftaleTilIndlæsning, kørselsaftaleEksisterende, "kørselsaftale")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    fejlIParameter: "kørselsaftaleVognløbsbillede" })
        if styresystemEksisterende != styresystemTilIndlæsning
            throw p6ForkertDataError(
                Format("Fejl i indlæsning af {3}. Forventet {3}: {1}, Eksisterende {3}: {2}", styresystemTilIndlæsning, styresystemEksisterende, "styresystem")
                , , , { forventetParameter: kørselsaftaleTilIndlæsning,
                    fundetParameter: kørselsaftaleEksisterende,
                    FejlIParameter: "KørselsaftaleVognløbsbillede" })


        mBoxFejl := this.enterOgHentMsgboxFejl()
        this.tjekP6Msgbox(mBoxFejl)

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
        if !starttid
        {
            SendInput("{Tab 6}")
        }
        else
        {
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
        }
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
        this.tjekP6Msgbox(p6_msgbox)

    }

    vognløbsbilledeIndtastStatistikgruppe() {
        Statistikgruppe := this.vognløb.parametre.Statistikgruppe.forventetIndhold

        if Statistikgruppe
            SendInput("!ø{tab 9}" Statistikgruppe)

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

        ; fiks tjek, er altid aktiv når kolonnen er navngivet
        ; if !this.vognløb.parametre.undtagneTransportTyper.iBrug
        ;     return
        vtype := this.vognløb.parametre.Vogntype.forventetIndhold
        trType := []

        vogntype2 := [
            "høj",
            "NJA",
            "barn1",
            "barn2",
            "barn3",
            "liftnet",
            "center",
            "liggende",
            "selepude",
            "cykel",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        vogntype3 := [
            "høj",
            "NJA",
            "barn1",
            "barn2",
            "barn3",
            "liftnet",
            "center",
            "liggende",
            "selepude",
            "cykel",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        vogntype4 := [
            "høj",
            "lav",
            "NJA",
            "transport",
            "tripstol",
            "barn1",
            "barn2",
            "barn3",
            "center",
            "selepude",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        vogntype5 := [
            "lav",
            "NJA",
            "barn1",
            "barn2",
            "barn3",
            "liftnet",
            "tmhjul",
            "tmlarve",
            "selepude",
            "tripstol",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        vogntype6 := [
            "lav",
            "NJA",
            "barn1",
            "barn2",
            "barn3",
            "liftnet",
            "tmhjul",
            "tmlarve",
            "selepude",
            "tripstol",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        vogntype7 := [
            "lav",
            "NJA",
            "barn1",
            "barn2",
            "barn3",
            "liftnet",
            "tmhjul",
            "tmlarve",
            "selepude",
            "tripstol",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " ",
            " "
        ]
        if this.vognløb.parametre.Vogntype.forventetIndhold
        {

            switch vtype
            {
                case 2:
                    trType := vogntype2
                case 3:
                    trType := vogntype3
                case 4:
                    trType := vogntype4
                case 5:
                    trType := vogntype5
                case 6:
                    trType := vogntype6
                case 7:
                    trType := vogntype7

            }

            {
                sendinput("!u")

                for transtype in trType
                    sendinput(transtype "{tab}"), sleep(10)
            }

        }
        else
        {
            UndtagneTransporttyper := this.vognløb.parametre.UndtagneTransporttyper.forventetIndhold

            if undtagnetransporttyper
            {
                sendinput("!u")

                for trtype in undtagnetransporttyper
                    sendinput(trtype "{tab}"), sleep(10)
            }

        }
    }
    vognløbsbilledeIndtastKapaciteter() {
        if !this.vognløb.parametre.Vogntype.iBrug
            return
        vtype := this.vognløb.parametre.Vogntype.forventetIndhold
        kap := {}

        vogntype2 := {
            fors: "1,2",
            bags: "3,2",
            bagage: "7,0",
            højsæde: "0",
            kplads: "0",
            hkplads: "0",
            liggende: "0",
            rest: "0",
            forshøj: "0",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "0",
            a_højsæde: "0",
            a_kplads: "0",
            a_hkplads: "0",
            a_liggende: "0",
            a_rest: "0",
            a_forshøj: "0",
        }

        vogntype3 := {
            fors: "0",
            bags: "0",
            bagage: "7,0",
            højsæde: "3,6",
            kplads: "0",
            hkplads: "0",
            liggende: "0",
            rest: "1",
            forshøj: "1,2",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "0",
            a_højsæde: "0",
            a_kplads: "0",
            a_hkplads: "0",
            a_liggende: "0",
            a_rest: "0",
            a_forshøj: "0",
        }

        vogntype4 := {
            fors: "1,2",
            bags: "2,4",
            bagage: "5,0",
            højsæde: "0",
            kplads: "0",
            hkplads: "1,0",
            liggende: "0",
            rest: "0",
            forshøj: "0",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "0",
            a_højsæde: "0",
            a_kplads: "0",
            a_hkplads: "0",
            a_liggende: "0",
            a_rest: "0",
            a_forshøj: "0",
        }

        vogntype5 := {
            fors: "0",
            bags: "0",
            bagage: "7",
            højsæde: "3,6",
            kplads: "0",
            hkplads: "2,4",
            liggende: "0",
            rest: "1,0",
            forshøj: "1,2",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "0",
            a_højsæde: "0",
            a_kplads: "0",
            a_hkplads: "0",
            a_liggende: "0",
            a_rest: "0",
            a_forshøj: "0",
        }
        vogntype6 := {
            fors: "0",
            bags: "0",
            bagage: "7,0",
            højsæde: "3,6",
            kplads: "0",
            hkplads: "2,4",
            liggende: "0",
            rest: "1",
            forshøj: "1,2",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "0",
            a_højsæde: "0",
            a_kplads: "0",
            a_hkplads: "0",
            a_liggende: "0",
            a_rest: "0",
            a_forshøj: "0",
        }
        vogntype7 := {
            fors: "1,2",
            bags: "0",
            bagage: "7,0",
            højsæde: "3,6",
            kplads: "0",
            hkplads: "1,9",
            liggende: "0",
            rest: "1",
            forshøj: "1,2",
            a_fors: "0",
            a_bags: "0",
            a_bagage: "7",
            a_højsæde: "3,6",
            a_kplads: "0",
            a_hkplads: "1,9",
            a_liggende: "8",
            a_rest: "0",
            a_forshøj: "1,9",
        }

        switch vtype
        {
            case 2:
                kap := vogntype2
            case 3:
                kap := vogntype3
            case 4:
                kap := vogntype4
            case 5:
                kap := vogntype5
            case 6:
                kap := vogntype6
            case 7:
                kap := vogntype7
        }

        SendInput("!k")
        sleep 40
        SendInput(kap.fors "{tab}")
        SendInput(kap.a_fors "{tab}")
        SendInput(kap.bags "{tab}")
        SendInput(kap.a_bags "{tab}")
        SendInput(kap.bagage "{tab}")
        SendInput(kap.a_bagage "{tab}")
        SendInput(kap.højsæde "{tab}")
        SendInput(kap.a_højsæde "{tab}")
        SendInput(kap.kplads "{tab}")
        SendInput(kap.a_kplads "{tab}")
        SendInput(kap.hkplads "{tab}")
        SendInput(kap.a_hkplads "{tab}")
        SendInput(kap.liggende "{tab}")
        SendInput(kap.a_liggende "{tab}")
        SendInput(kap.rest "{tab}")
        SendInput(kap.a_rest "{tab}")
        SendInput(kap.forshøj "{tab}")
        SendInput(kap.a_forshøj "{tab}")
    }
    ændrVognløbsbilledeAfslut()
    {
        SendInput("{enter}")

        p6_msgbox := this.kopierVærdi("ctrl", 1)
        this.tjekP6Msgbox(p6_msgbox)

        return
    }


    vognløbsbilledeIndhentStatistikgruppe() {

        SendInput("!ø {tab 9}")

        statistikGruppe := this.kopierVærdi("ctrl")
        this.vognløb.parametre.setParameterEksisterende("Statistikgruppe", statistikGruppe)

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

    vognløbsbilledeIndhentHjemzone() {

        SendInput("{tab 6}")

        startzone := this.kopierVærdi("Ctrl")
        startzonePar := this.vognløb.parametre.startzone
        this.vognløb.parametre.setParameterEksisterende("Startzone", startzone)
        SendInput("{tab}")

        slutzone := this.kopierVærdi("Ctrl")
        slutzonePar := this.vognløb.parametre.slutzone
        this.vognløb.parametre.setParameterEksisterende("Slutzone", slutzone)
        SendInput("{tab}")

        hjemzone := this.kopierVærdi("Ctrl")
        hjemzonePar := this.vognløb.parametre.hjemzone
        this.vognløb.parametre.setParameterEksisterende("Hjemzone", hjemzone)

        SendInput("{Enter}")

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
        this.tjekP6Msgbox(p6_msgbox)


    }

    vognløbsbilledeIndhentØvrige() {

        if this.vognløb.gyldigeKolonner.Vognløbsnotering.iBrug
        {
            SendInput("!v+{Up}")
            Vognløbsnotering := this.kopierVærdi("ctrl", , , 1)
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

    tjekP6Msgbox(pMsgbox) {
        if (InStr(pMsgBox, "eksistere ikke"))
            throw P6MsgboxError(this.vognløb.parametre.Vognløbsnummer.forventetIndhold " - " this.vognløb.parametre.vognløbsdato.forventetIndhold ": Vognløb findes ikke på dato.")
        if (InStr(pMsgBox, "Kan ikke nå frem til første opgave"))
            throw P6MsgboxError(this.vognløb.parametre.vognløbsnummer.forventetIndhold " - " this.vognløb.parametre.vognløbsdato.forventetIndhold ": Vognløb kan ikke nå første køreordre.")
        if (InStr(pMsgBox, "Køretidsfaktorerne"))
            throw P6MsgboxError(this.vognløb.parametre.vognløbsnummer.forventetIndhold " - " this.vognløb.parametre.vognløbsdato.forventetIndhold ": Vognløb kan ikke nå sidste køreordre.")
        if (InStr(pMsgbox, "Forkert tidspunkt"))
            throw P6MsgboxError("Fejl i vognløbstidspunkter.",)
        if (InStr(pMsgbox, "kørselsaftalen ikke registreret"))
            throw P6MsgboxError("Kørselsaftalen eksisterer ikke i P6.",)
        if (InStr(pMsgbox, "vognløbet er allerede"))
            throw P6MsgboxError("Vognløbet eksisterer allerede på dato.",)
        if InStr(pMsgBox, "Zone ikke registreret")
            throw (P6MsgboxError("Zonen findes ikke i P6"))
        if InStr(pMsgBox, "Zone skal angives")
            throw (P6MsgboxError("Zone kan ikke angives tom"))
        if (InStr(pMsgBox, "for langt for modellen"))
            throw P6MsgboxError("Vognløbet er for langt for modellen")
        if InStr(pMsgBox, "samme transporttype")
            throw (P6MsgboxError("Transporttype er opført to gange"))
        if InStr(pMsgBox, "Transporttypen")
            throw (P6MsgboxError("Transporttype findes ikke i P6"))
        if InStr(pMsgBox, "Vløbsklasen")
            throw (P6MsgboxError("Vognløbskategorien findes ikke i P6"))
        if InStr(pMsgBox, "Planskema ikke registreret")
            throw (P6MsgboxError("Planskema findes ikke i P6"))
        if InStr(pMsgBox, "Planskema skal angives")
            throw (P6MsgboxError("Planskema kan ikke angives tomt"))
        if InStr(pMsgBox, "Økonomiskema ikke registreret")
            throw (P6MsgboxError("Økonomiskema findes ikke i P6"))
        if InStr(pMsgBox, "Økonomiskema skal angives")
            throw (P6MsgboxError("Økonomiskema kan ikke angives tomt"))
        if InStr(pMsgBox, "Vognens kapacitet")
            throw (P6MsgboxError("Kapacitetfejl på vognløb"))

        if (InStr(pMsgBox, "----"))
            throw P6MsgboxError(this.vognløb.parametre.Vognløbsnummer.ForventetIndhold " - " this.vognløb.parametre.vognløbsdato.ForventetIndhold ": Ikke-kategoriseret fejl på vognløb.")
    }

    funkKørselsaftaleÆndrHjemzone() {

        this.navAktiverP6Vindue()
        this.navLukAlleVinduer()
        this.navVindueKørselsaftale()
        this.kørselsaftaleIndtastKørselsaftale()
        this.kørselsaftaleTjekKørselsaftaleOgStyresystem()
        this.kørselsaftaleÆndr()
        this.kørselsaftaleIndtastStatistikgruppe()
        this.kørselsaftaleIndtastNormalHjemzone()
        this.kørselsaftaleIndtastVognmandLinie1()
        this.kørselsaftaleIndtastVognmandLinie2()
        this.kørselsaftaleIndtastVognmandLinie3()
        this.kørselsaftaleIndtastVognmandLinie4()
        this.kørselsaftaleAfslut()
    }
    funkVognløbsbilledeIndhentHjemzone() {

        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeIndhentHjemzone()
        this.vognløbsbilledeIndhentStatistikgruppe()
        this.ændrVognløbsbilledeAfslut()
    }
    funkVognløbsbilledeÆndrHjemzone() {

        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeIndtastÅbningstiderOgZone()
        this.vognløbsbilledeIndtastStatistikgruppe()
        this.ændrVognløbsbilledeAfslut()
    }
    funkÆndrVognløb()
    {
        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDato()
        this.vognløbsbilledeÆndrVognløb()
        this.vognløbsbilledeTjekKørselsaftaleOgStyresystem()
        this.vognløbsbilledeIndtastÅbningstiderOgZone()
        this.vognløbsbilledeIndtastØvrige()
        this.vognløbsbilledeIndtastTransporttyper()
        this.vognløbsbilledeIndtastKapaciteter()
        this.ændrVognløbsbilledeAfslut()
        return
    }
    funkIndlægVognløb(){
        this.navVindueVognløb()
        this.vognløbsbilledeIndtastVognløbOgDatoTilIndlæg()
        this.vognløbsbilledeIndtastKørselsaftaleOgStyresystemTilIndlæg()
        this.vognløbsbilledeIndtastÅbningstiderOgZone()
        this.vognløbsbilledeIndtastØvrige()
        this.vognløbsbilledeIndtastTransporttyper()
        this.vognløbsbilledeIndtastKapaciteter()
        this.ændrVognløbsbilledeAfslut()
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

    Budnummer := { parameterNavn: "Budnummer", kolonneNavn: "Budnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    Vognløbsnummer := { parameterNavn: "Vognløbsnummer", kolonneNavn: "Vognløbsnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 5 }
    Vogntype := { parameterNavn: "Vogntype", kolonneNavn: "Vogntype", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 1 }
    Vognløbsdato := { parameterNavn: "Vognløbsdato", kolonneNavn: "Ugedage", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    VognløbsdatoStart := { parameterNavn: "VognløbsdatoStart", kolonneNavn: "Ugedage", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    VognløbsdatoSlut := { parameterNavn: "VognløbsdatoSlut", kolonneNavn: "Ugedage", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    VognløbsdatoNormalSlut := { parameterNavn: "VognløbsdatoNormalslut", kolonneNavn: "Ugedage", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    VognløbsdatoSidsteSlut := { parameterNavn: "VognløbsdatoSidsteSlut", kolonneNavn: "Ugedage", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    Kørselsaftale := { parameterNavn: "Kørselsaftale", kolonneNavn: "Kørselsaftale", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 4 }
    Styresystem := { parameterNavn: "Styresystem", kolonneNavn: "Styresystem", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 2 }
    Starttid := { parameterNavn: "Starttid", kolonneNavn: "Starttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    Sluttid := { parameterNavn: "Sluttid", kolonneNavn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    NormalSluttid := { parameterNavn: "Sluttid", kolonneNavn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    SidsteSluttid := { parameterNavn: "Sluttid", kolonneNavn: "Sluttid", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    Hjemzone := { parameterNavn: "Hjemzone", kolonneNavn: "Hjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    Startzone := { parameterNavn: "Startzone", kolonneNavn: "Hjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    Slutzone := { parameterNavn: "Slutzone", kolonneNavn: "Hjemzone", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    NormalHjemzone := { parameterNavn: "NormalHjemzone", kolonneNavn: "Hjemzone", forventetIndhold: this.Hjemzone.forventetIndhold, eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    chfKontaktNummer := { parameterNavn: "chfKontaktNummer", kolonneNavn: "chfKontaktNummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 8 }
    Vognløbskategori := { parameterNavn: "Vognløbskategori", kolonneNavn: "Vognløbskategori", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 6 }
    Planskema := { parameterNavn: "Planskema", kolonneNavn: "Planskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 10 }
    Økonomiskema := { parameterNavn: "Økonomiskema", kolonneNavn: "Økonomiskema", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 10 }
    Statistikgruppe := { parameterNavn: "Statistikgruppe", kolonneNavn: "Statistikgruppe", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 10 }
    StatistikgruppeKørselsaftale := { parameterNavn: "StatistikgruppeKørselsaftale", kolonneNavn: "Statistikgruppe", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 10 }
    Vognløbsnotering := { parameterNavn: "Vognløbsnotering", kolonneNavn: "Vognløbsnotering", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 240 }
    VognmandLinie1 := { parameterNavn: "VognmandLinie1", kolonneNavn: "VognmandLinie1", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 25 }
    VognmandLinie2 := { parameterNavn: "VognmandLinie2", kolonneNavn: "VognmandLinie2", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 24 }
    VognmandLinie3 := { parameterNavn: "VognmandLinie3", kolonneNavn: "VognmandLinie3", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 29 }
    VognmandLinie4 := { parameterNavn: "VognmandLinie4", kolonneNavn: "VognmandLinie4", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 20 }
    VognmandKontaktnummer := { parameterNavn: "VognmandKontaktnummer", kolonneNavn: "VognmandKontaktnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 11 }
    ObligatoriskVognmand := { parameterNavn: "ObligatoriskVognmand", kolonneNavn: "ObligatoriskVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: 5 }
    KørselsaftaleVognmand := { parameterNavn: "KørselsaftaleVognmand", kolonneNavn: "KørselsaftaleVognmand", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    Ugedage := { parameterNavn: "Ugedage", kolonneNavn: "Ugedage", forventetIndhold: Array(), eksisterendeIndhold: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    UndtagneTransporttyper := { parameterNavn: "UndtagneTransporttyper", kolonneNavn: "UndtagneTransporttyper", forventetIndhold: Array(), eksisterendeIndhold: Array(), ForventetMenIkkeIEksisterende: Array(), EksisterendeMenIkkeIForventet: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "", maxArray: 20 }
    KørerIkkeTransporttyper := { parameterNavn: "KørerIkkeTransporttyper", kolonneNavn: "KørerIkkeTransporttyper", forventetIndhold: Array(), eksisterendeIndhold: Array(), ForventetMenIkkeIEksisterende: Array(), EksisterendeMenIkkeIForventet: Array(), fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "", maxArray: 10 }


    ; kun i P6
    PauseRegel := { parameterNavn: "PauseRegel", kolonneNavn: "PauseRegel", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    PauseDynamisk := { parameterNavn: "PauseDynamisk", kolonneNavn: "PauseDynamisk", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    PauseStart := { parameterNavn: "PauseStart", kolonneNavn: "PauseStart", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }
    PauseSlut := { parameterNavn: "PauseSlut", kolonneNavn: "PauseSlut", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }


    danParameterObj(pParameterNavn) {

        if this.HasOwnProp(pParameterNavn)
            return

        this.%pParameterNavn% := { navn: pParameterNavn, forventetIndhold: "", eksisterendeIndhold: "", fejl: 0 }

        return this.%pParameterNavn%
    }

    setParameterForventet(pParameterObj, pForventet) {


        pParameterObj.forventetIndhold := pForventet


    }

    /**
     * 
     * @param pParamaetObj 
     * @returns Bool
     */
    tjekParameterForFejl(pParameterNavn) {

        parameterObj := this.%pParameterNavn%

        forventetParameterIndhold := StrLower(parameterObj.forventetIndhold)
        eksisterendeParameterIndhold := strlower(parameterObj.eksisterendeIndhold)
        fundetFejl := parameterObj.fejl

        if !forventetParameterIndhold
            return

        if fundetFejl
            return

        if forventetParameterIndhold != eksisterendeParameterIndhold
            parameterObj.fejl := 1
        else
            parameterObj.fejl := 0
    }

    setParameterEksisterende(pParameterNavn, pParameterIndhold) {

        this.%pParameterNavn%.eksisterendeIndhold := pParameterIndhold

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

        mBoxFejl := this.enterOgHentMsgboxFejl()
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