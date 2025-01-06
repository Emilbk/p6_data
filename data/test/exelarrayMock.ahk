#Include lib\json.ahk
; Fra VLMock.xlsx

excelDataMock := [["Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone", "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer", "ChfKontaktNummer", "Vognløbskategori", "Planskema", "Økonomiskema", "Statistikgruppe", "Vognløbsnotering", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper"], ["24-2267", "31400", "3400", "1", "06:00", "15:00", "årh143", "årh143", "årh143", "Vognmand 1 ApS", "Lukket pr. 12-11-24", "Gadenr 1", "8001 By", "70112220", "701122010", "FV8", "31400", "31400", "2GVEL", "Blabla", "17/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"], ["24-2268", "31401", "3401", "2", "06:01", "15:01", "årh144", "årh144", "årh144", "Vognmand 2 ApS", "Lukket pr. 13-11-24", "Gadenr 2", "8002 By", "70112221", "2", "FV9", "31401", "31401", "3GVEL", "Blabla 2", "18/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "center", "tripstol"]]
excelDataUgyldigMock := [["Ugyldig", "Budnummer", "Vognløbsnummer", "Kørselsaftale", "Styresystem", "Starttid", "Sluttid", "Startzone", "Slutzone", "Hjemzone", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3", "VognmandLinie4", "VognmandKontaktnummer", "ChfKontaktNummer", "Vognløbskategori", "Planskema", "Økonomiskema", "Statistikgruppe", "Vognløbsnotering", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "Ugedage", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "UndtagneTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper", "ugyldig", "KørerIkkeTransporttyper", "KørerIkkeTransporttyper"], ["ugyldig", "24-2267", "31400", "3400", "1", "06:00", "15:00", "årh143", "årh143", "årh143", "Vognmand 1 ApS", "Lukket pr. 12-11-24", "Gadenr 1", "8001 By", "70112220", "70112211", "FV8", "31400", "31400", "2GVEL", "Blabla", "17/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "ugyldig", "center", "tripstol"], ["ugyldig", "24-2268", "31401", "3401", "2", "06:01", "15:01", "årh144", "årh144", "årh144", "Vognmand 2 ApS", "Lukket pr. 13-11-24", "Gadenr 2", "8002 By", "70112221", "2", "FV9", "31401", "31401", "3GVEL", "Blabla 2", "18/11/2024", "ma", "ti", "on", "to", "fr", "lø", "sø", "LAV", "NJA", "TRANSPORT", "TMHJUL", "TMLARVE", "FYN24", "SYD24", "MIDT24", "FYN", "CROSSER", "høj", "nja", "barn1", "barn2", "barn3", "liftnet", "selepude", "liggende", "ugyldig", "center", "tripstol"]]


kolonneNavne := Map()
rIndhol := Array()

for rækkeIndex, kolonneNavn in excelDataMock
{
    rIndhol.Push(Map())
    for kolonneIndex, rækkeIndhold in kolonneNavn
    {
        if rækkeIndex = 1
            kolonneNavne.Set(rækkeIndhold, kolonneIndex)
        ; else
            ; rIndhol[rækkeIndex].set(excelDataMock[1][kolonneIndex], rækkeIndhold)
    }
}

kolonneNavne["Ugedage"] := Array()
kolonneNavne["UndtagneTransporttyper"] := Array()
kolonneNavne["KørerIkkeTransporttyper"] := Array()

for kolonneNavn in excelDataMock[1]
{
    if kolonneNavn = "Ugedage"
        kolonneNavne["Ugedage"].Push(A_Index)
    if kolonneNavn = "UndtagneTransporttyper"
        kolonneNavne["UndtagneTransporttyper"].Push(A_Index)
    if kolonneNavn = "KørerIkkeTransporttyper"
        kolonneNavne["KørerIkkeTransporttyper"].Push(A_Index)
}
rIndhol.RemoveAt(1)
return