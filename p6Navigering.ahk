#Requires AutoHotkey v2.0

; Aktiverer P6-vindue, hvis ikke aktivt

p6_clipwait(p_valg, p_type?)
{
    clipwaitTid := 0.3
    clipwaitTid2 := 0.5
    clipwaitTidMsgbox := 1
    input := Map("shift", "+{F10}c", "ctrl", "^c")
    if IsSet(p_type)
    {
        A_Clipboard := ""
        SendInput input[p_valg]
        clipwait clipwaitTidMsgbox
        return A_Clipboard
    }
    if !IsSet(p_type)
    {
        A_Clipboard := ""
        SendInput input[p_valg]
        clipwait clipwaitTid
        while A_Clipboard = ""
        {
            if a_index > 10
                throw (Error("Clipboardfejl"))
            else
            {
                SendInput input[p_valg]
                ClipWait clipwaitTid2
            }
        }
        return A_Clipboard
    }
}

P6_aktiver()
{

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

; Aktiverer alt-menu i P6, tager op til to taste-sekvenser
P6_alt_menu(tast1, tast2?)
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


; Aktiverer alt-menu i P6, tager op til to taste-sekvenser
P6_luk_vinduer()
{
    SendInput "{esc}{alt}"
    sleep 20
    Sendinput "{v 2}{Down 2}{Enter}"

    return
}

P6_nav_kørselsaftale()
{
    P6_aktiver()
    P6_alt_menu("t", "k")


    return
}


P6_nav_vognløbsbillede()
{
    P6_aktiver()
    P6_alt_menu("t", "l")


    return
}

; TODO hvordan håndteres dato-array?
p6_åben_vognløb(p_vl_obj)
{
    P6_aktiver()
    SendInput(p_vl_obj.vl_data["Vognløbsnummer"])
    SendInput "{tab}"
    SendInput(p_vl_obj.vl_data["Dato"])
    SendInput("{enter}")
    sleep 20
    p6_clipwait("ctrl", 1)
    if (InStr(A_Clipboard, "eksistere ikke"))
        throw Error("Ikke registreret - TODO")
    ; tjek af korrekt vognløb
    indlæst_vognløbsnummer := p6_clipwait("shift")
    SendInput("{tab}")
    indlæst_dato := p6_clipwait("ctrl")
    ; lav fornuftigt system
    if !(p_vl_obj.vl_data["Vognløbsnummer"] = indlæst_vognløbsnummer and p_vl_obj.vl_data["Dato"] = indlæst_dato)
        throw (Error("Fejl i indlæsning, ikke det forventede vognløbsnummer på forventet dato"))
    return
}
; TODO beslut hvordan hele vognløb-funktion skal struktureres
p6_åben_vognløb_kørselsaftale(p_vl_obj)
{
    P6_aktiver()
    SendInput("^æ")
    oprindelig_kørselsaftale := p6_clipwait("shift")
    SendInput("{tab}")
    oprindelig_styresystem := p6_clipwait("shift")
    SendInput("{tab}")
    if (p_vl_obj.vl_data["Kørselsaftale"] != oprindelig_kørselsaftale or p_vl_obj.vl_data["Styresystem"] != oprindelig_styresystem)
        throw (Error("Fejl i indlæsning, åben kørselsaftale er ikke den forventede"))
    ; SendInput(p_vl_obj.vl_data["Kørselsaftale"])
    SendInput "{tab}"
    SendInput(p_vl_obj.vl_data["Styresystem"])
    SendInput("{tab}")
    ; tjek af korrekt vognløb, omskriv
    indlæst_kørselsaftale := p6_clipwait("shift")
    SendInput("{tab}")
    indlæst_styresystem := p6_clipwait("shift")
    if (p_vl_obj.vl_data["Kørselsaftale"] = indlæst_kørselsaftale and p_vl_obj.vl_data["Styresystem"] = indlæst_styresystem)
        korrekt := 1
    ; MsgBox "korrekt"
    SendInput("{enter}")
    ; p6_msgbox := p6_clipwait("ctrl", 1)
    ; if InStr(p6_msgbox, "ikke registreret")
        ; throw (Error("Kørselsaftalen findes ikke i P6"))

    return
}

; lav modulær opbygning
p6_åben_vognløb_åbningstider(p_vl_obj)
{
    ; hvorfor breaker den script her?
    ; P6_aktiver()
    ; p6_clipwait("ctrl")
    ; P6_aktiver()
    SendInput(p_vl_obj.vl_data["Dato"] "{tab}")
    SendInput(p_vl_obj.vl_data["Starttid"] "{tab}")
    SendInput(p_vl_obj.vl_data["Dato"] "{tab}")
    SendInput(p_vl_obj.vl_data["Sluttid"] "{tab}")
    SendInput(p_vl_obj.vl_data["Dato"] "{tab}")
    SendInput(p_vl_obj.vl_data["Sluttid"] "{tab}")
    SendInput(p_vl_obj.vl_data["Startzone"] "{tab}")
    SendInput(p_vl_obj.vl_data["Slutzone"] "{tab}")
    SendInput(p_vl_obj.vl_data["Hjemzone"] "{tab}")
    SendInput("{enter}")
    p6_msgbox := p6_clipwait("ctrl", 1)
    if InStr(p6_msgbox, "Zone ikke registreret")
        throw (Error("Zonen findes ikke i P6"))
    if InStr(p6_msgbox, "Zone skal angives")
        throw (Error("Zonen er udfyldt tom"))
    return
}
; modulær opbygning
p6_åben_vognløb_resten(p_vl_obj)
{
    P6_aktiver()
    if (p_vl_obj.vl_data["Vognløbsnotering"])
        SendInput("!p{tab 11}+{Up}" p_vl_obj.vl_data["Vognløbsnotering"])
    if (p_vl_obj.vl_data["MobilnrChf"])
        SendInput("!ø{tab 2}" p_vl_obj.vl_data["MobilnrChf"])
    if (p_vl_obj.vl_data["Vognløbskategori"])
        SendInput("!ø{tab 3}" p_vl_obj.vl_data["Vognløbskategori"])
    if (p_vl_obj.vl_data["Planskema"])
        SendInput("!ø{tab 6}" p_vl_obj.vl_data["Planskema"])
    if (p_vl_obj.vl_data["Økonomiskema"])
        SendInput("!ø{tab 8}" p_vl_obj.vl_data["Økonomiskema"])
    if (p_vl_obj.vl_data["Statistikgruppe"])
        SendInput("!ø{tab 9}" p_vl_obj.vl_data["Statistikgruppe"])

    if (p_vl_obj.vl_data["Undtagne transporttyper"])
    {
        SendInput("!ø{tab 10}")
        for trtype in p_vl_obj.vl_data["Undtagne transporttyper"]
            SendInput("{tab}" trtype)
    }
    SendInput("{enter}")
    return
}
p6_afslut_indlæsning_vognløb(p_vl_obj)
{
    ; P6_aktiver()
    p6_msgbox := p6_clipwait("ctrl", 1)
    if InStr(p6_msgbox, "Transporttypen")
        throw (Error("Transporttype findes ikke i P6"))
    ; p6_clipwait("shift")
    return A_Clipboard
}

p6_åben_kørselsaftale(p_vl_obj)
{
    P6_nav_kørselsaftale()
    sleep 100
    SendInput(p_vl_obj.vl_data["Kørselsaftale"])
    SendInput "{tab}"
    SendInput(p_vl_obj.vl_data["Styresystem"])
    SendInput("{enter}")
    p6_msgbox := p6_clipwait("ctrl", 1)
    if (InStr(p6_msgbox, "ikke registreret"))
        throw Error("Ikke registreret - TODO")
    ; tjek af korrekt kørselsaftale
    indlæst_kørselaftale := p6_clipwait("shift")
    indlæst_kørselaftale := A_Clipboard
    SendInput("{tab}")
    indlæst_styresystem := p6_clipwait("shift")
    SendInput("{tab}")
    ; SendInput("^{F4}")
    indlæst_styresystem := A_Clipboard
    if (p_vl_obj.vl_data["Kørselsaftale"] = indlæst_kørselaftale and p_vl_obj.vl_data["Styresystem"] = indlæst_styresystem)
    ; MsgBox "korrekt"
        return
}

p6_indlæs_data_kørselsaftale_æ()
{
    SendInput("^æ")

    return
}

p6_indlæs_data_kørselsaftale_planskema(p_vl_obj)
{
    SendInput("!p")
    A_Clipboard := ""
    tidligere_planskema := p6_clipwait("ctrl")
    SendInput(p_vl_obj.vl_data["Planskema"] "{tab}!p")
    indlæst_planskema := p6_clipwait("ctrl")
    if (p_vl_obj.vl_data["Planskema"] = indlæst_planskema)
        korrekt := 1
    return
}

; p6_indlæs_data_kørselsaftale_økonomiskema(p_vl_obj)
; {
;     SendInput("!p{tab 4}")
;     A_Clipboard := ""
;     SendInput("^c")
;     clipwait clipwaitTid
;     tidligere_planskema := A_Clipboard
;     SendInput(p_vl_obj.vl_data["Planskema"] "{tab}!p{tab 4}")
;     A_Clipboard := ""
;     SendInput("^c")
;     clipwait clipwaitTid
;     indlæst_planskema := A_Clipboard
;     if (p_vl_obj.vl_data["Planskema"] = indlæst_planskema)
;         korrekt := 1
;     return
; }
