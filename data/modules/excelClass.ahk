/************************************************************************
 * @description Excel-Class
 * @author ebk
 * @date 2024/10/18
 * @version 0.1.0
 * 
 * Indlæser excelark
 * Datastruktur:
 * Hver række defineres separat i array aktivWorksheet.SheetArray
 * Herunder defineres hver celle i en map, med kolonnenavn som key og celleværdi som value
 * Hvis flere kolonner har samme navn oprettes i stedet for celleværdien et array med de samlede celleværdier
 * worksheetArrayRække(alle rækker)[en specifik række]{en specific celle knyttet til kolonnenavn}([array hvis flere af den samme kolonne])
 ***********************************************************************/

#Requires AutoHotkey v2.0
#SingleInstance Force
#include ../include.ahk

class excel {

    app := ComObject("Excel.Application")

    quit() {

        this.app.quit()
    }
}

class excelIndlæsVlData extends excel {
    __New(pExcelFil, pArkNavnEllerNummer) {
        this.app := ComObject("Excel.Application")

        this.helperIndlæsAlt(pExcelFil, pArkNavnEllerNummer)

        this.quit()
    }

    aktivWorkbook := Object()
    aktivWorksheet := Object()
    aktivWorksheet.SheetArray := Array()
    aktivWorksheet.KolonneNavnOgNummer := Map()

    ugyldigeKolonneNavne := {}

    gyldigeKolonneNavnOgNummer := staticGyldigeKolonneNavnOgNummer

    åbenWorkbookReadonly(pExcelFil) {


        this.excelFilNavnLong := pExcelFil
        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        this.aktivWorkbookComObj := this.app.Workbooks.open(pExcelFil, "ReadOnly" = true)
        return
    }

    setAktivRækkeNummer(pAktivRække) {

        this.aktivWorksheet.AktivRække := pAktivRække

    }


    setAktivWorksheet(pSheetNummerEllerNavn) {
        this.aktivWorksheetComObj := this.aktivWorkbookComObj.Sheets(pSheetNummerEllerNavn)

        return
    }


    getGyldigeKolonner() {

        return this.gyldigeKolonneNavnOgNummer
    }
    dataFindBrugtExcelRangeIAktivWorksheet() {

        this.aktivWorksheet.RækkerEnd := this.aktivWorksheetComObj.usedrange.rows.count
        this.aktivWorksheet.KolonnerEnd := this.aktivWorksheetComObj.usedrange.columns.count

        return
    }

    dataIndlæsAktivRangetilArray() {
        this.aktivWorksheet.SheetArray := this.aktivWorksheetComObj.usedrange.value

        return
    }

    erGyldigKolonne(pKolonneTilTjek) {


        if this.gyldigeKolonneNavnOgNummer.HasOwnProp(pKolonneTilTjek)
            return 1
        else
            return 0
    }


    dataIndlæsKolonneNavnogNummerTilMap() {
        if not this.aktivWorksheet.SheetArray is ComObjArray
            throw Error("aktivWorksheet.SheetArray er ikke indlæst")

        rækkeKolonneNavn := 1
        loop this.aktivWorksheet.KolonnerEnd {
            kolonneNummer := A_Index
            nuværendeKolonneNavn := this.aktivWorksheet.SheetArray[rækkeKolonneNavn, kolonneNummer]
            if this.erGyldigKolonne(nuværendeKolonneNavn)
            {
                if Type(this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.kolonneNummer) != "Array"
                {
                    this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.iBrug := 1
                    this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.kolonneNummer := kolonneNummer
                }

                if Type(this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.kolonneNummer) = "Array"
                {
                    this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.iBrug := 1
                    this.gyldigeKolonneNavnOgNummer.%nuværendeKolonneNavn%.kolonneNummer.push(kolonneNummer)
                }
            }
            else
                this.ugyldigeKolonneNavne.%nuværendeKolonneNavn% := { kolonneNavn: nuværendeKolonneNavn, kolonneNummer: kolonneNummer }
        }
    }


    dataIndlæsRækkeArrayMinusKolonneNavne() {

        arrayKolonneNavne := map("KørerIkkeTransporttyper", 0, "UndtagneTransporttyper", 0, "Ugedage", 0,)
        this.vlArray := Array()
        loop this.aktivWorksheet.RækkerEnd {
            rækkenummer := A_Index
            kolonneNavnRække := 1
            this.vlArray.Push(Map("KørerIkkeTransporttyper", Array(), "UndtagneTransporttyper", Array(), "Ugedage", Array()))
            loop this.aktivWorksheet.KolonnerEnd {
                kolonneNummer := A_Index
                kolonneNavn := this.aktivWorksheet.SheetArray[kolonneNavnRække, kolonneNummer]
                celleIndhold := this.aktivWorksheet.SheetArray[rækkenummer, kolonneNummer]
                if this.gyldigeKolonneNavnOgNummer.HasProp(kolonneNavn)
                {
                    if Type(celleIndhold) = "Float"
                        celleIndhold := String(Floor(celleIndhold))
                    if arrayKolonneNavne.Has(kolonneNavn)
                        this.vlArray[rækkenummer][kolonneNavn].Push(celleIndhold)
                    else
                        this.vlArray[rækkenummer][kolonneNavn] := celleIndhold
                }
            }
        }
        this.vlArray.RemoveAt(1)
        return
    }

    dataVerificerInputTidspunkt() {

               if !this.gyldigeKolonneNavnOgNummer.Starttid.iBrug or !this.gyldigeKolonneNavnOgNummer.Sluttid.iBrug
            return

        datestr := "20241116"
        loop this.aktivWorksheet.RækkerEnd {

            rækkeNummer := A_Index
            if rækkeNummer = 1
                continue
            kolonneStartTid := this.gyldigeKolonneNavnOgNummer.Starttid.kolonneNummer
            kolonneSluttTid := this.gyldigeKolonneNavnOgNummer.Sluttid.kolonneNummer

            celleStartTid := this.aktivWorksheet.SheetArray[rækkeNummer, kolonneStartTid]
            celleSlutTid := this.aktivWorksheet.SheetArray[rækkeNummer, kolonneSluttTid]

            if !celleSlutTid or !celleStartTid
                return

            if !InStr(celleStartTid, ":") or !InStr(celleSlutTid, ":")
                throw Error("Forkert format i række " A_Index " - tidspunkt skal angives tt:mm")

            if InStr(celleSlutTid, "*")
                testStrSlut := SubStr(celleSlutTid, 1, 5)
            else
                testStrSlut := celleSlutTid

            testStrStartArray := StrSplit(celleStartTid, ":")
            testStrStart := testStrStartArray[1] . testStrStartArray[2]
            testStrSlutArray := StrSplit(testStrSlut, ":")
            testStrSlut := testStrSlutArray[1] . testStrSlutArray[2]

            dateTestStart := datestr . testStrStart
            dateTestSlut := datestr . testStrSlut


            if !IsTime(dateTestStart) or !IsTime(dateTestSlut)
                throw Error("Forkert format i række " A_Index " - tidspunkt skal angives tt:mm eller tt:mm*")
        }

    }

    helperIndlæsAlt(pExcelFil, pArkNavnEllerNummer) {
        this.åbenWorkbookReadonly(pExcelFil)
        this.setAktivWorksheet(pArkNavnEllerNummer)
        this.dataFindBrugtExcelRangeIAktivWorksheet()
        this.dataIndlæsAktivRangetilArray()
        this.dataIndlæsKolonneNavnogNummerTilMap()
        this.dataIndlæsRækkeArrayMinusKolonneNavne()
        this.dataVerificerInputTidspunkt()
        return
    }

    getKolonneNavnOgNummer()
    {
        return this.aktivWorksheet.KolonnenavneOgNummer
    }

    getVlArray() {
        return this.vlArray
    }

    quit() {
        this.app.quit()
        return
    }

}


class excelLavNyWorkbook extends excel {

    __New(pPath) {
        this.LavNyWorkbook(pPath)
    }

    LavNyWorkbook(pExcelFil) {

        this.aktivWorkbookComObj := this.app.Workbooks.add()

        this.excelFilNavnLong := pExcelFil
        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        this.aktivWorkbookComObj.Saveas(pExcelFil)
        this.quit()
        return
    }

}

class excelBehandlWorkbook extends excel {

    åbenWorkbookReadWrite(pExcelFil) {
        if !FileExist(pExcelFil)
            throw Error("Excel-fil findes ikke")

        this.excelFilNavnLong := pExcelFil

        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        this.aktivWorkbookComObj := this.app.Workbooks.open(this.excelFilNavnLong, "ReadOnly" = false)

    }

    gyldigKolonneNavnOgNummer := {}
    aktivSheet := ""

    setAktivSheet(pSheetNavnEllerNummer) {

        this.aktivSheet := this.app.Worksheets(pSheetNavnEllerNummer)

        this.aktivSheet.activate()
    }
    setGyldigeKolonner(pGyldigKolonne) {

        this.gyldigKolonneNavnOgNummer := pGyldigKolonne
    }
    gemWorkbook() {

        this.aktivWorkbookComObj.Save()
    }

    kolonneNavnogNummerTilArray(pGyldigKolonne) {


        outputArray := this.KolonneNavnOgNummerArray

        kolonneNummerHojest := this.kolonneNummerArrayFindHøjestePlads(pGyldigKolonne)

        loop kolonneNummerHojest
            outputArray.Push("")

        outputArray := this.kolonneNavnogNummerUdfyldArray(outputArray, pGyldigKolonne)

        this.kolonneNavnogNummerArray := outputArray

        return outputArray
    }

    kolonneNummerArrayFindHøjestePlads(pGyldigKolonne) {

        kolonneNummerHojest := 0
        KolonneNavnOgNummer := pGyldigKolonne


        for kolonneNavn, kolonneObj in KolonneNavnOgNummer.OwnProps()
        {
            if type(kolonneObj.KolonneNummer) = "Integer"
                if kolonneObj.KolonneNummer >= kolonneNummerHojest
                    kolonneNummerHojest := kolonneObj.KolonneNummer
            if Type(kolonneObj.KolonneNummer) = "Array"
                for kolonneNummerArray in kolonneObj.KolonneNummer
                    if kolonneNummerArray >= kolonneNummerHojest
                        kolonneNummerHojest := kolonneNummerArray
        }
        return kolonneNummerHojest

    }

    kolonneNavnogNummerUdfyldArray(pInputArray, pGyldigKolonne) {

        KolonneNavnOgNummer := pGyldigKolonne


        for kolonneNavn, kolonneObj in KolonneNavnOgNummer.OwnProps()
        {
            if Type(kolonneObj.kolonneNummer) = "Integer"
                pInputArray[kolonneObj.kolonneNummer] := kolonneNavn
            if Type(kolonneObj.kolonneNummer) = "Array"
                for kolonneNummerArray in kolonneObj.kolonneNummer
                    pInputArray[kolonneNummerArray] := kolonneNavn
        }

        return pInputArray
    }

    udfyldVognløbRækker(pVl, pRækkenummer) {

        pVl.parametre.sorterUndtagneTransporttyperEksisterende()
        pVl.parametre.sorterUndtagneTransporttyperForventet()
        pVl.parametre.sorterKørerIkkeTransporttyperEksisterende()
        pVl.parametre.sorterKørerIkkeTransporttyperForventet()

        for parameterNavn, parameterObj in pVl.parametre.OwnProps()
        {
            if !parameterObj.eksisterendeIndhold
                continue
            parameterObj := pVl.parametre.%parameterNavn%
            if parameterObj.eksisterendeIndhold
                if this.gyldigKolonneNavnOgNummer.HasOwnProp(parameterNavn)
                {
                    if type(parameterObj.eksisterendeIndhold) != "Array"
                    {
                        parmameterKolonne := parameterobj.navn
                        kolonneNummer := this.gyldigKolonneNavnOgNummer.%parmameterKolonne%.kolonneNummer

                        aktivCelle := this.aktivSheet.cells(pRækkenummer, kolonneNummer)
                        aktivCelle.value := parameterObj.eksisterendeIndhold
                    }
                    if type(parameterObj.eksisterendeIndhold) = "Array"
                    {
                        kolonneNummer := this.gyldigKolonneNavnOgNummer.%parameterobj.navn%.kolonneNummer[1]
                        for celleIndhold in parameterObj.eksisterendeIndhold
                        {
                            aktivCelleIndhold := celleIndhold
                            parmameterKolonne := parameterobj.navn

                            aktivCelle := this.aktivSheet.cells(pRækkenummer, kolonneNummer)
                            aktivCelle.value := aktivCelleIndhold
                            kolonneNummer += 1
                        }
                    }
                }
        }
        this.aktivSheet.columns().AutoFit
        ; kolonneNavn := pvl.parametre.vognløbsnummer.navn
        ; parameterObj := pvl.parametre.%kolonneNavn%


    }

}

class udfyldTestExcelArk extends excel {

    testVl := vognløbObj()

    testVl.parametre.Budnummer.eksisterendeIndhold := "24-2267"
    testVl.parametre.Vognløbsnummer.eksisterendeIndhold := "31400"
    testVl.parametre.Planskema.eksisterendeIndhold := "31400"
    testVl.parametre.Økonomiskema.eksisterendeIndhold := "31400"
    testVl.parametre.Kørselsaftale.eksisterendeIndhold := "3400"
    testVl.parametre.KørselsaftaleVognmand.eksisterendeIndhold := "3VOGNM"
    testVl.parametre.ObligatoriskVognmand.eksisterendeIndhold := "3BAR"
    testVl.parametre.Vognløbsnotering.eksisterendeIndhold := "Type 2, GV 19-02, Autostol 0-13"
    testVl.parametre.Statistikgruppe.eksisterendeIndhold := "2GVEL"
    testVl.parametre.Styresystem.eksisterendeIndhold := "1"
    testVl.parametre.Starttid.eksisterendeIndhold := "19:00"
    testVl.parametre.Sluttid.eksisterendeIndhold := "02:00*"
    testVl.parametre.Slutzone.eksisterendeIndhold := "Årh142"
    testVl.parametre.Startzone.eksisterendeIndhold := "Årh142"
    testVl.parametre.Hjemzone.eksisterendeIndhold := "Årh142"
    testVl.parametre.ChfKontaktNummer.eksisterendeIndhold := "70112210"
    testVl.parametre.VognmandKontaktnummer.eksisterendeIndhold := "70112220"
    testVl.parametre.VognmandLinie1.eksisterendeIndhold := "Vognmand ApS"
    testVl.parametre.VognmandLinie2.eksisterendeIndhold := "Ny hjemzone pr. 01-12-2024"
    testVl.parametre.VognmandLinie3.eksisterendeIndhold := "Hjemzonegade 101"
    testVl.parametre.VognmandLinie4.eksisterendeIndhold := "Hjemzoneby, 8000"
    testVl.parametre.Vognløbskategori.eksisterendeIndhold := "FG9"
    testVl.parametre.UndtagneTransporttyper.eksisterendeIndhold := ["Nja", "CrosSER", "Barn1", "Barn2", "TTJHjul", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]
    testVl.parametre.KørerIkkeTransporttyper.eksisterendeIndhold := ["Crosser", "Barn3," "NJA", "Barn1", A_Space, A_Space, A_Space, A_Space, A_Space, A_Space]
    testVl.parametre.ugedage.eksisterendeIndhold := ["30-11-2024", "MA", "TI", "ON", "TO", "FR", "LØ", "SØ"]


    TestKolonneNavnOgNummer := {
        Budnummer: { kolonneNavn: "Budnummer", kolonneNummer: 1, iBrug: 0, kolonneKommentar: 0 },
        Vognløbsnummer: { kolonneNavn: "Vognløbsnummer", kolonneNummer: 2, iBrug: 0, kolonneKommentar: "Påkrævet" },
        Kørselsaftale: { kolonneNavn: "Kørselsaftale", kolonneNummer: 3, iBrug: 0, kolonneKommentar: "Påkrævet" },
        Styresystem: { kolonneNavn: "Styresystem", kolonneNummer: 4, iBrug: 0, kolonneKommentar: "Påkrævet" },
        Starttid: { kolonneNavn: "Starttid", kolonneNummer: 5, iBrug: 0, kolonneKommentar: "Vognløbets starttid i tekstformat `"tt:mm`"." },
        Sluttid: { kolonneNavn: "Sluttid", kolonneNummer: 6, iBrug: 0, overMidnat: 0, kolonneKommentar: "Vognløbets sluttid i tekstformat `"tt:mm`". Hvis vognløbet løber over midnat tilføjes *, f. eks. `"01:30*`" (nødvendigt for at vognløbsdato defineres korrekt)." },
        Startzone: { kolonneNavn: "Startzone", kolonneNummer: 7, iBrug: 0, kolonneKommentar: 0 },
        Slutzone: { kolonneNavn: "Slutzone", kolonneNummer: 8, iBrug: 0, kolonneKommentar: 0 },
        Hjemzone: { kolonneNavn: "Hjemzone", kolonneNummer: 9, iBrug: 0, kolonneKommentar: 0 },
        ChfKontaktNummer: { kolonneNavn: "ChfKontaktNummer", kolonneNummer: 10, iBrug: 0, kolonneKommentar: 0 },
        VognmandKontaktNummer: { kolonneNavn: "VognmandKontaktNummer", kolonneNummer: 11, iBrug: 0, kolonneKommentar: 0 },
        Vognløbskategori: { kolonneNavn: "Vognløbskategori", kolonneNummer: 12, iBrug: 0, kolonneKommentar: 0 },
        Planskema: { kolonneNavn: "Planskema", kolonneNummer: 13, iBrug: 0, kolonneKommentar: 0 },
        Økonomiskema: { kolonneNavn: "Økonomiskema", kolonneNummer: 14, iBrug: 0, kolonneKommentar: 0 },
        Statistikgruppe: { kolonneNavn: "Statistikgruppe", kolonneNummer: 15, iBrug: 0, kolonneKommentar: 0 },
        Vognløbsnotering: { kolonneNavn: "Vognløbsnotering", kolonneNummer: 16, iBrug: 0, kolonneKommentar: "Fast notat på vognløb. Bruges pt. på alle vognløbsdage." },
        VognmandLinie1: { kolonneNavn: "VognmandLinie1", kolonneNummer: 17, iBrug: 0, kolonneKommentar: "Første linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie2: { kolonneNavn: "VognmandLinie2", kolonneNummer: 18, iBrug: 0, kolonneKommentar: "Anden linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie3: { kolonneNavn: "VognmandLinie3", kolonneNummer: 19, iBrug: 0, kolonneKommentar: "Tredje linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie4: { kolonneNavn: "VognmandLinie4", kolonneNummer: 20, iBrug: 0, kolonneKommentar: "Fjerde linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        ObligatoriskVognmand: { kolonneNavn: "ObligatoriskVognmand", kolonneNummer: 21, iBrug: 0, kolonneKommentar: 0 },
        KørselsaftaleVognmand: { kolonneNavn: "KørselsaftaleVognmand", kolonneNummer: 22, iBrug: 0, kolonneKommentar: "Vognmandsparameter defineret i kørselsaftalen." },
        Ugedage: { kolonneNavn: "Ugedage", kolonneNummer: Array(23, 24, 25, 26, 27, 28, 29), iBrug: 0, kolonneKommentar: "Faste ugedage i P6-format (MA, TI osv.) Kan også tage konkret dato i formatet `"dd-mm-åååå`". Én dato pr. kolonne, så mange kolonner som ønsket" },
        UndtagneTransporttyper: { kolonneNavn: "UndtagneTransporttyper", kolonneNummer: Array(30, 31, 32, 33, 34, 35), iBrug: 0, kolonneKommentar: "Undtagne transporttyper som defineret i vognløbet. Definer op til 20 stk. Én transporttype pr. kolonne" },
        KørerIkkeTransporttyper: { kolonneNavn: "KørerIkkeTransporttyper", kolonneNummer: Array(36, 37, 38, 39, 40, 41, 42, 43, 44,), iBrug: 0, kolonneKommentar: "Undtagne transporttyper som defineret i kørselsaftalen. Definer op til 10 stk. Én transporttype pr. kolonne." },
    }

    tjekketVLKolonneNavnOgNummer := {
        Vognløbsnummer: { kolonneNavn: "Vognløbsnummer", kolonneNummer: 1, iBrug: 0, kolonneKommentar: 0 },
        Vognløbsdato: { kolonneNavn: "Vognløbsdato", kolonneNummer: 2, iBrug: 0, kolonneKommentar: 0 },
        Kørselsaftale: { kolonneNavn: "Kørselsaftale", kolonneNummer: 3, iBrug: 0, kolonneKommentar: 0 },
        Styresystem: { kolonneNavn: "Styresystem", kolonneNummer: 4, iBrug: 0, kolonneKommentar: 0 },
        Starttid: { kolonneNavn: "Starttid", kolonneNummer: 5, iBrug: 0, kolonneKommentar: "Vognløbets starttid i tekstformat `"tt:mm`"." },
        Sluttid: { kolonneNavn: "Sluttid", kolonneNummer: 6, iBrug: 0, overMidnat: 0, kolonneKommentar: "Vognløbets sluttid i tekstformat `"tt:mm`". Hvis vognløbet løber over midnat tilføjes *, f. eks. `"01:30*`" (nødvendigt for at vognløbsdato defineres korrekt)." },
        Startzone: { kolonneNavn: "Startzone", kolonneNummer: 7, iBrug: 0, kolonneKommentar: 0 },
        Slutzone: { kolonneNavn: "Slutzone", kolonneNummer: 8, iBrug: 0, kolonneKommentar: 0 },
        Hjemzone: { kolonneNavn: "Hjemzone", kolonneNummer: 9, iBrug: 0, kolonneKommentar: 0 },
        ChfKontaktNummer: { kolonneNavn: "ChfKontaktNummer", kolonneNummer: 10, iBrug: 0, kolonneKommentar: 0 },
        VognmandKontaktNummer: { kolonneNavn: "VognmandKontaktNummer", kolonneNummer: 11, iBrug: 0, kolonneKommentar: 0 },
        Vognløbskategori: { kolonneNavn: "Vognløbskategori", kolonneNummer: 12, iBrug: 0, kolonneKommentar: 0 },
        Planskema: { kolonneNavn: "Planskema", kolonneNummer: 13, iBrug: 0, kolonneKommentar: 0 },
        Økonomiskema: { kolonneNavn: "Økonomiskema", kolonneNummer: 14, iBrug: 0, kolonneKommentar: 0 },
        Statistikgruppe: { kolonneNavn: "Statistikgruppe", kolonneNummer: 15, iBrug: 0, kolonneKommentar: 0 },
        Vognløbsnotering: { kolonneNavn: "Vognløbsnotering", kolonneNummer: 16, iBrug: 0, kolonneKommentar: "Fast notat på vognløb. Bruges pt. på alle vognløbsdage." },
        VognmandLinie0: { kolonneNavn: "VognmandLinie1", kolonneNummer: 18, iBrug: 0, kolonneKommentar: "Første linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie1: { kolonneNavn: "VognmandLinie2", kolonneNummer: 19, iBrug: 0, kolonneKommentar: "Anden linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie2: { kolonneNavn: "VognmandLinie3", kolonneNummer: 20, iBrug: 0, kolonneKommentar: "Tredje linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        VognmandLinie3: { kolonneNavn: "VognmandLinie4", kolonneNummer: 21, iBrug: 0, kolonneKommentar: "Fjerde linie af `"Ansvarlig`"-feltet defineret i kørselsaftalen." },
        ObligatoriskVognmand: { kolonneNavn: "ObligatoriskVognmand", kolonneNummer: 22, iBrug: 0, kolonneKommentar: 0 },
        KørselsaftaleVognmand: { kolonneNavn: "KørselsaftaleVognmand", kolonneNummer: 23, iBrug: 0, kolonneKommentar: "Vognmandsparameter defineret i kørselsaftalen." },
        UndtagneTransporttyper: { kolonneNavn: "UndtagneTransporttyper", kolonneNummer: Array(24, 25, 26, 27, 28, 29), iBrug: 0, kolonneKommentar: "Undtagne transporttyper som defineret i vognløbet. Definer op til 20 stk. Én transporttype pr. kolonne" },
        KørerIkkeTransporttyper: { kolonneNavn: "KørerIkkeTransporttyper", kolonneNummer: Array(30, 31, 32, 33), iBrug: 0, kolonneKommentar: "Undtagne transporttyper som defineret i kørselsaftalen. Definer op til 10 stk. Én transporttype pr. kolonne." },
    }

    TestKolonneNavnOgNummerArray := Array()

    åbenWorkbookReadWrite(pExcelFil) {
        if !FileExist(pExcelFil)
            throw Error("Excel-fil findes ikke")

        this.excelFilNavnLong := pExcelFil

        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension
        this.aktivWorkbookComObj := this.app.Workbooks.open(this.excelFilNavnLong, "ReadOnly" = false)
        this.aktivSheet := this.app.ActiveWorkbook.activeSheet

    }

    gemWorkbook() {

        this.aktivWorkbookComObj.save()
        this.quit()
    }

    navngivSheet(pSheet, pNavn) {

        this.app.Worksheets(pSheet).name := pNavn
    }

    setAktivSheet(pSheetNavnEllerNummer) {

        this.aktivSheet := this.app.Worksheets(pSheetNavnEllerNummer)

        this.aktivSheet.activate()
    }

    kolonneNavnogNummerTilArray() {


        outputArray := this.TestKolonneNavnOgNummerArray

        kolonneNummerHojest := this.kolonneNummerArrayFindHøjestePlads()

        loop kolonneNummerHojest
            outputArray.Push("")

        outputArray := this.kolonneNavnogNummerUdfyldArray(outputArray)

        this.kolonneNavnogNummerArray := outputArray

        return outputArray
    }

    kolonneNummerArrayFindHøjestePlads() {

        kolonneNummerHojest := 0
        KolonneNavnOgNummer := this.TestKolonneNavnOgNummer


        for kolonneNavn, kolonneObj in KolonneNavnOgNummer.OwnProps()
        {
            if type(kolonneObj.KolonneNummer) = "Integer"
                if kolonneObj.KolonneNummer >= kolonneNummerHojest
                    kolonneNummerHojest := kolonneObj.KolonneNummer
            if Type(kolonneObj.KolonneNummer) = "Array"
                for kolonneNummerArray in kolonneObj.KolonneNummer
                    if kolonneNummerArray >= kolonneNummerHojest
                        kolonneNummerHojest := kolonneNummerArray
        }
        return kolonneNummerHojest

    }

    kolonneNavnogNummerUdfyldArray(pInputArray) {

        KolonneNavnOgNummer := this.TestKolonneNavnOgNummer


        for kolonneNavn, kolonneObj in KolonneNavnOgNummer.OwnProps()
        {
            if Type(kolonneObj.kolonneNummer) = "Integer"
                pInputArray[kolonneObj.kolonneNummer] := kolonneNavn
            if Type(kolonneObj.kolonneNummer) = "Array"
                for kolonneNummerArray in kolonneObj.kolonneNummer
                    pInputArray[kolonneNummerArray] := kolonneNavn
        }

        return pInputArray
    }

    lavTestKolonnerExcel() {

        kolonneRække := 1
        for kolonneNavn in this.TestKolonneNavnOgNummerArray
        {
            aktivKolonne := A_Index
            aktivKolonneObj := this.TestKolonneNavnOgNummer.%kolonneNavn%
            aktivCelle := this.aktivSheet.cells(kolonneRække, aktivKolonne)
            aktivCelle.value := aktivKolonneObj.kolonneNavn
            this.aktivSheet.columns().AutoFit
            if aktivKolonneObj.kolonneKommentar
            {
                ; aktivCelle.addcomment()
                ; aktivCelle.comment.text(aktivKolonneObj.kolonneKommentar)
                aktivCelle.addcommentthreaded(aktivKolonneObj.kolonneKommentar)
            }

        }

    }

    udfyldVognløbRækker(pVl, pRækkenummer) {

        pVl.parametre.sorterUndtagneTransporttyperEksisterende()
        pVl.parametre.sorterUndtagneTransporttyperForventet()
        pVl.parametre.sorterKørerIkkeTransporttyperEksisterende()
        pVl.parametre.sorterKørerIkkeTransporttyperForventet()

        for parameterNavn, parameterObj in pVl.parametre.OwnProps()
        {
            if !parameterObj.eksisterendeIndhold
                continue
            parameterObj := pVl.parametre.%parameterNavn%
            if parameterObj.eksisterendeIndhold
                if type(parameterObj.eksisterendeIndhold) != "Array"
                {
                    parmameterKolonne := parameterobj.navn
                    kolonneNummer := this.TestKolonneNavnOgNummer.%parmameterKolonne%.kolonneNummer

                    aktivCelle := this.aktivSheet.cells(pRækkenummer, kolonneNummer)
                    aktivCelle.value := parameterObj.eksisterendeIndhold
                }
            if type(parameterObj.eksisterendeIndhold) = "Array"
            {
                kolonneNummer := this.testKolonneNavnOgNummer.%parameterobj.navn%.kolonneNummer[1]
                for celleIndhold in parameterObj.eksisterendeIndhold
                {
                    aktivCelleIndhold := celleIndhold
                    parmameterKolonne := parameterobj.navn

                    aktivCelle := this.aktivSheet.cells(pRækkenummer, kolonneNummer)
                    aktivCelle.value := aktivCelleIndhold
                    kolonneNummer += 1
                }
            }
            this.aktivSheet.columns().AutoFit
        }
        ; kolonneNavn := pvl.parametre.vognløbsnummer.navn
        ; parameterObj := pvl.parametre.%kolonneNavn%


    }

    lavExcelTemplate(pPath) {

        if FileExist(pPath)
            FileDelete(pPath)
        excelNyWorkbook := excelLavNyWorkbook(pPath)
        excelNyWorkbook.quit()
        this.åbenWorkbookReadWrite(pPath)
        this.setAktivSheet(1)
        this.navngivSheet(1, "Alle Gyldige Kolonner")
        this.kolonneNavnogNummerTilArray()
        this.lavTestKolonnerExcel()
        this.udfyldVognløbRækker(this.testVl, 2)
        this.gemWorkbook()

    }
}

class mockExcelP6Data extends excel {

    __New() {

        gyldigeKolonner := Map(
            "Budnummer", 1,
            "Vognløbsnummer", 1,
            "Kørselsaftale", 1,
            "Styresystem", 1,
            "Startzone", 1,
            "Slutzone", 1,
            "Hjemzone", 1,
            "MobilnrChf", 1,
            "Vognløbskategori", 1,
            "Planskema", 1,
            "Økonomiskema", 1,
            "Statistikgruppe", 1,
            "Vognløbsnotering", 1,
            "Starttid", 1,
            "Sluttid", 1,
            "Sluttid", 1,
            "Undtagne transporttyper", 1,
            "Ugedage", 1
        )
        this.kolonneNavnOgNummer := Map("Budnummer", 1, "Vognløbsnummer", 2)

        this.aktivWorksheet.SheetArray := Array()

        this.aktivWorksheet.SheetArray.Push(Map(
            "Budnummer", "24-2267",
            "Vognløbsnummer", "31400",
            "Kørselsaftale", "3400",
            "Styresystem", "1",
            "Startzone", "ÅRH142",
            "Slutzone", "ÅRH142",
            "Hjemzone", "ÅRH142",
            "MobilnrChf", "701122010",
            "Vognløbskategori", "FG9",
            "Planskema", "31400",
            "Økonomiskema", "31400",
            "Statistikgruppe", "2GVEL",
            "Vognløbsnotering", "Notering1",
            "Starttid", "09",
            "Sluttid", "23",
            "Sluttid", "23",
            "UndtagneTransporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Vognløbsdato", "",
            "Ugedage", Array("ma", "ma", "ma")
        ))
        this.aktivWorksheet.SheetArray.Push(Map(
            "Budnummer", "24-2266",
            "Vognløbsnummer", "31400",
            "Kørselsaftale", "3400",
            "Styresystem", "1",
            "Startzone", "ÅRH143",
            "Slutzone", "ÅRH143",
            "Hjemzone", "ÅRH143",
            "MobilnrChf", "701122011",
            "Vognløbskategori", "FG9",
            "Planskema", "31401",
            "Økonomiskema", "31401",
            "Statistikgruppe", "2GVEL",
            "Vognløbsnotering", "Notering2",
            "Starttid", "10",
            "Sluttid", "22",
            "Sluttid", "22",
            "UndtagneTransporttyper", Array("LAV", "NJA", "TRANSPORT", "TMHJUL"),
            "Vognløbsdato", "",
            "Ugedage", Array("ma")
        ))

        this.færdigbehandletData := { kolonneNavnOgNummer: this.kolonneNavnOgNummer, rækkerSomMapIArray: this.aktivWorksheet.SheetArray }

    }

    getKolonneNavnOgNummer()
    {
        return this.færdigbehandletData.kolonneNavnOgNummer
    }

    getRækkeData()
    {
        return this.færdigbehandletData.rækkerSomMapIArray
    }

    get() {

        return this.færdigbehandletData.rækkerSomMapIArray
    }
}