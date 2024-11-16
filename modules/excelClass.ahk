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

class excelIndlæsVlData extends Class {
    __New(pExcelFil, pArkNavnEllerNummer) {
        this.app := ComObject("Excel.Application")

        this.helperIndlæsAlt(pExcelFil, pArkNavnEllerNummer)

        this.quit()
    }

    aktivWorkbook := Object()
    aktivWorksheet := Object()
    aktivWorksheet.SheetArray := Array()
    aktivWorksheet.KolonneNavnOgNummer := Map()

    gyldigeKolonneNavnOgNummer := {
        Budnummer: { kolonneNavn: "Budnummer", kolonneNummer: 0, iBrug: 0 },
        Vognløbsnummer: { kolonneNavn: "Vognløbsnummer", kolonneNummer: 0, iBrug: 0 },
        Vognløbsdato: { kolonneNavn: "Vognløbsdato", kolonneNummer: 0, iBrug: 0 },
        Kørselsaftale: { kolonneNavn: "Kørselsaftale", kolonneNummer: 0, iBrug: 0 },
        Styresystem: { kolonneNavn: "Styresystem", kolonneNummer: 0, iBrug: 0 },
        Starttid: { kolonneNavn: "Starttid", kolonneNummer: 0, iBrug: 0 },
        Sluttid: { kolonneNavn: "Sluttid", kolonneNummer: 0, iBrug: 0, overMidnat: 0 },
        Startzone: { kolonneNavn: "Startzone", kolonneNummer: 0, iBrug: 0 },
        Slutzone: { kolonneNavn: "Slutzone", kolonneNummer: 0, iBrug: 0 },
        Hjemzone: { kolonneNavn: "Hjemzone", kolonneNummer: 0, iBrug: 0 },
        MobilnrChf: { kolonneNavn: "MobilnrChf", kolonneNummer: 0, iBrug: 0 },
        Vognløbskategori: { kolonneNavn: "Vognløbskategori", kolonneNummer: 0, iBrug: 0 },
        Planskema: { kolonneNavn: "Planskema", kolonneNummer: 0, iBrug: 0 },
        Økonomiskema: { kolonneNavn: "Økonomiskema", kolonneNummer: 0, iBrug: 0 },
        Statistikgruppe: { kolonneNavn: "Statistikgruppe", kolonneNummer: 0, iBrug: 0 },
        Vognløbsnotering: { kolonneNavn: "Vognløbsnotering", kolonneNummer: 0, iBrug: 0 },
        VognmandLinie1: { kolonneNavn: "VognmandLinie1", kolonneNummer: 0, iBrug: 0 },
        VognmandLinie2: { kolonneNavn: "VognmandLinie2", kolonneNummer: 0, iBrug: 0 },
        VognmandLinie3: { kolonneNavn: "VognmandLinie3", kolonneNummer: 0, iBrug: 0 },
        VognmandLinie4: { kolonneNavn: "VognmandLinie4", kolonneNummer: 0, iBrug: 0 },
        VognmandTelefon: { kolonneNavn: "vognmandTelefon", kolonneNummer: 0, iBrug: 0 },
        ObligatoriskVognmand: { kolonneNavn: "ObligatoriskVognmand", kolonneNummer: 0, iBrug: 0 },
        KørselsaftaleVognmand: { kolonneNavn: "KørselsaftaleVognmand", kolonneNummer: 0, iBrug: 0 },
        Ugedage: { kolonneNavn: "Ugedage", kolonneNummer: Array(), iBrug: 0 },
        UndtagneTransporttyper: { kolonneNavn: "UndtagneTransporttyper", kolonneNummer: Array(), iBrug: 0 },
        KørerIkkeTransportyper: { kolonneNavn: "KørerIkkeTransportyper", kolonneNummer: Array(), iBrug: 0 },
    }

    ugyldigeKolonneNavne := {}


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
                this.ugyldigeKolonneNavne.%nuværendeKolonneNavn% := {kolonneNavn: nuværendeKolonneNavn, kolonneNummer: kolonneNummer}
        }
    }


    dataIndlæsRækkeArrayMinusKolonneNavne() {

        this.vlArray := Array()
        loop this.aktivWorksheet.RækkerEnd {
            rækkenummer := A_Index
            kolonneNavnRække := 1
            this.vlArray.Push(Map())
            loop this.aktivWorksheet.KolonnerEnd {
                kolonneNummer := A_Index
                kolonneNavn := this.aktivWorksheet.SheetArray[kolonneNavnRække, kolonneNummer]
                celleIndhold := this.aktivWorksheet.SheetArray[rækkenummer, kolonneNummer]
                if this.gyldigeKolonneNavnOgNummer.HasProp(kolonneNavn)
                    {
                if Type(celleIndhold) = "Float"
                    celleIndhold := String(Floor(celleIndhold))
                if (this.vlArray[rækkenummer].Has(kolonneNavn)) {
                    if (type(this.vlArray[rækkenummer][kolonneNavn]) != "Array")
                        this.vlArray[rækkenummer][kolonneNavn] := Array(this.vlArray[
                            rækkenummer][kolonneNavn])
                    this.vlArray[rækkenummer][kolonneNavn].push(celleIndhold)
                }
                else
                    this.vlArray[rækkenummer][kolonneNavn] := celleIndhold
                    }
            }
        }
        this.vlArray.RemoveAt(1)
        return
    }

    helperIndlæsAlt(pExcelFil, pArkNavnEllerNummer) {
        this.åbenWorkbookReadonly(pExcelFil)
        this.setAktivWorksheet(pArkNavnEllerNummer)
        this.dataFindBrugtExcelRangeIAktivWorksheet()
        this.dataIndlæsAktivRangetilArray()
        this.dataIndlæsKolonneNavnogNummerTilMap()
        this.dataIndlæsRækkeArrayMinusKolonneNavne()
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


class excelLavNyWorkbook {

    LavNyWorkbook(pExcelFil) {

        this.aktivWorkbookComObj := this.app.Workbooks.add()

        this.excelFilNavnLong := pExcelFil
        SplitPath(this.excelFilNavnLong, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.excelFilNavn := varFilNavn
        this.excelFilDir := varFilDir
        this.excelFilNavnUdenExtension := varFilNavnUdenExtension

        this.aktivWorkbookComObj.Saveas(pExcelFil)
        return
    }

}

class excelBehandlWorkbook {

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

    gemWorkbook() {

        this.aktivWorkbookComObj.Save()
    }

    kolonneNavnogNummerTilArray(pKolonneNavnOgNummer) {

        outputArray := Array()

        kolonneNummerHojest := this.kolonneNummerArrayFindHøjestePlads(pKolonneNavnOgNummer)

        loop kolonneNummerHojest
            outputArray.Push("")

        outputArray := this.kolonneNavnogNummerUdfyldArray(pKolonneNavnOgNummer, outputArray)

        this.kolonneNavnogNummerArray := outputArray

        return outputArray
    }

    kolonneNummerArrayFindHøjestePlads(pKolonneNavnOgNummer) {

        kolonneNummerHojest := 0
        for kolonneNavn, kolonneNummer in pKolonneNavnOgNummer
        {
            if type(kolonneNummer) = "Integer"
                if kolonneNummer >= kolonneNummerHojest
                    kolonneNummerHojest := kolonneNummer
            if Type(kolonneNummer) = "Array"
                for kolonneNummerArray in kolonneNummer
                    if kolonneNummerArray >= kolonneNummerHojest
                        kolonneNummerHojest := kolonneNummerArray
        }
        return kolonneNummerHojest

    }

    kolonneNavnogNummerUdfyldArray(pKolonneNavnOgNummer, pInputArray) {

        for kolonneNavn, kolonneNummer in pKolonneNavnOgNummer
        {
            if Type(kolonneNummer) = "Integer"
                pInputArray[kolonneNummer] := kolonneNavn
            if Type(kolonneNummer) = "Array"
                for kolonneNummerArray in kolonneNummer
                    pInputArray[kolonneNummerArray] := kolonneNavn
        }

        return pInputArray
    }
}

class mockExcelP6Data extends Class {

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