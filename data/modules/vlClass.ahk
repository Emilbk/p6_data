#Requires AutoHotkey v2.0
#Include ../deepCopy.ahk

class VognløbConstructor {
    __New(pExcelInput, pGyldigeKolonner) {

        this.gyldigeKolonner := pGyldigeKolonner
        this.behandlVognløsbsdata(pExcelInput)
    }

    vognløbInput := ""
    vognløbOutput := { masterVognløb: VognløbObj(), vognløbsListe: Array() }
    ; vognløbOutput := { masterVognløb: VognløbObj(), vognløbsListe: Array() }


    /**
     *  Main interface
     * @param pVognløbsdata 
     * @returns {Array} 
     */
    behandlVognløsbsdata(pVognløbsdata) {
        this.setVognløbsdataTilBehandling(pVognløbsdata)
        this.danVognløbOutput()

        return this.vognløbOutput
    }

    setVognløbsdataTilBehandling(pVognløbsData) {

        this.vognløbInput := pVognløbsData
        return
    }
    /**
     * Array med hvert vognløb, underinddelt i array vlObj med konkrete vognløbsdage
     * @returns {Array} 
     */
    getBehandletVognløbsArray()
    {
        return this.vognløbOutput
    }


    danVognløbOutput() {

        this.danVognløbsArray()
        this.danVognløbIVognløbsArray()
    }


    danVognløbsArray() {
        vognløbsarray := this.vognløbInput
        vognløbOutput := this.vognløbOutput

        for vognløb in vognløbsarray
            vognløbOutput.vognløbsListe.push(Array())

        this.vognløbOutput := vognløbOutput
        return
    }

    ; omskriv
    danVognløbIVognløbsArray() {
        vognløbsarray := this.vognløbInput
        vognløbOutput := this.vognløbOutput

        for enkeltVognløbInput in vognløbsarray
        {
            ugedagArrayCount := 0
            
            outerIndex := A_Index
            masterVL := vognløbOutput.masterVognløb := VognløbObj()
            masterVL.setVognløbsDataTilIndlæsning(enkeltVognløbInput, this.gyldigeKolonner)
            masterVl.udfyldundtagneTransportTyperArray()
            masterVl.udfyldKørerIkkeTransporttyperArray()
            masterVl.setGyldigeKolonner(this.gyldigeKolonner)
            for ugedag in enkeltVognløbInput["Ugedage"]
            {
                if ugedag = ""
                    continue
                ugedagArrayCount += 1
                ugedag := Format("{:U}", ugedag)

                vognløbOutput.vognløbsListe[outerIndex].push(Array)
                vognløbOutput.vognløbsListe[outerIndex][ugedagArrayCount] := VognløbObj()
                aktueltVognløb := vognløbOutput.vognløbsListe[outerIndex][ugedagArrayCount]
                aktueltVognløb.setVognløbsDataTilIndlæsning(enkeltVognløbInput, this.gyldigeKolonner)
                aktueltVognløb.setVognløbsdato(ugedag)
                aktueltVognløb.tjekSlutTidOverMidnat()
                aktueltVognløb.udfyldundtagneTransportTyperArray()
                aktueltVognløb.udfyldKørerIkkeTransporttyperArray()
                aktueltVognløb.setGyldigeKolonner(this.gyldigeKolonner)

                ; vognløbOutput[outerIndex][ugedagArrayCount].setfejlLog(enkeltVognløbInput)
                ; vognløbOutput[outerIndex][ugedagArrayCount].setfejlLog(enkeltVognløbInput)
                aktueltVognløb.parametre.Vognløbsdato.forventetIndhold := ugedag

            }
        }

        return vognløbOutput
    }

}

/** Repræsenterer et specifikt vognløb på en specifik dato (eller fast ugedag) */
class VognløbObj
{

    parametre := parameterClass()
    fejlLog := fejlLogObj()
    gyldigeKolonner := {}

    setGyldigeKolonner(pGyldigeKolonner) {

        this.gyldigeKolonner := pGyldigeKolonner

    }
    setFejlLog(pFejlObj) {

        this.fejlLog := pFejlObj
        this.fejlLog.fundetFejl := 1

    }
    udfyldUndtagneTransportTyperArray() {

        antalTransportTyper := this.parametre.UndtagneTransporttyper.forventetIndhold.Length
        ønsketAntalTransportTyper := 20

        if antalTransportTyper > ønsketAntalTransportTyper
            throw Error(Format("Fejl i parameter:`n{1}.`nKan maks. indeholde {2} kolonner. Er {3} kolonner.", "UndtagneTransportTyper", 2, antalTransportTyper))

        while this.parametre.UndtagneTransporttyper.forventetIndhold.Length != ønsketAntalTransportTyper
            this.parametre.UndtagneTransporttyper.forventetIndhold.push(A_Space)

        for index, transportType in this.parametre.UndtagneTransporttyper.forventetIndhold
            if transportType = ""
                this.parametre.UndtagneTransporttyper.forventetIndhold[index] := A_Space
    }

    udfyldKørerIkkeTransporttyperArray() {
        antalTransportTyper := this.parametre.KørerIkkeTransporttyper.forventetIndhold.Length
        ønsketAntalTransportTyper := 10
        if antalTransportTyper > ønsketAntalTransportTyper
            throw Error(Format("Fejl i parameter:`n{1}.`nKan maks. indeholde {2} kolonner. Er {3} kolonner.", "KørerIkkeTransportTyper", 10, antalTransportTyper))


        while this.parametre.KørerIkkeTransporttyper.forventetIndhold.Length != ønsketAntalTransportTyper
            this.parametre.KørerIkkeTransporttyper.forventetIndhold.push(A_Space)

        for index, transportType in this.parametre.KørerIkkeTransporttyper.forventetIndhold
            if transportType = ""
                this.parametre.KørerIkkeTransporttyper.forventetIndhold[index] := A_Space
    }

    setVognløbsdato(pVognløbsdato) {
        for parameterNavn, parameterObj in this.parametre.OwnProps()
            if (parameterObj.kolonneNavn = "Ugedage")
                parameterObj.forventetIndhold := pVognløbsdato
    }

    tjekSlutTidOverMidnat() {


        ; Hvorfor?
        if !this.parametre.Starttid.forventetIndhold or !this.parametre.Sluttid.forventetIndhold or !this.parametre.Vognløbsdato.forventetIndhold
            return

        slutTid := this.parametre.Sluttid.forventetIndhold
        fasteDageArray := ["MA", "TI", "ON", "TO", "FR", "LØ", "SØ"]
        arrayPos := 0
        for index, fastdag in fasteDageArray
            if fastdag = this.parametre.Vognløbsdato.forventetIndhold
                arrayPos := index

        VlDato := this.parametre.Vognløbsdato.forventetIndhold
        if IsInteger(SubStr(VlDato, 1, 2))
            if InStr(slutTid, "*")
            {
                år := SubStr(vlDato, -4, 4)
                måned := SubStr(vlDato, 4, 2)
                dag := SubStr(vlDato, 1, 2)

                dateStr := år måned dag
                nyDato := DateAdd(datestr, 1, "Days")

                slutDato := FormatTime(nyDato, "dd-MM-yyyy")

                for parameterNavn, parameterObj in this.parametre.OwnProps()
                {
                    if (parameterObj.kolonneNavn = "Ugedage")
                        if (parameterObj.parameterNavn != "Vognløbsdato" and parameterObj.parameterNavn != "VognløbsdatoStart")
                            parameterObj.forventetIndhold := slutdato
                }
                return

            }
            else
            {

                for parameterNavn, parameterObj in this.parametre.OwnProps()
                {
                    if (parameterObj.kolonneNavn = "Ugedage")
                        if (parameterObj.parameterNavn != "Vognløbsdato" and parameterObj.parameterNavn != "VognløbsdatoStart")
                            parameterObj.forventetIndhold := vlDato
                }
                return
            }

        if InStr(slutTid, "*")
        {
            slutTid := SubStr(slutTid, 1, 5)
            if arrayPos < 7
            {
                for parameterNavn, parameterObj in this.parametre.OwnProps()
                {
                    if (parameterObj.kolonneNavn = "Ugedage")
                        if (parameterObj.parameterNavn != "Vognløbsdato" and parameterObj.parameterNavn != "VognløbsdatoStart")
                            parameterObj.forventetIndhold := fastedagearray[arrayPos + 1]
                }
            }
            else
                for parameterNavn, parameterObj in this.parametre.OwnProps()
                {
                    if (parameterObj.kolonneNavn = "Ugedage")
                        if (parameterObj.parameterNavn != "Vognløbsdato" and parameterObj.parameterNavn != "VognløbsdatoStart")
                            parameterObj.forventetIndhold := fastedagearray[1]
                }


        }
    }
    setVognløbsDataTilIndlæsning(pVLParameter, pGyldigeKolonner) {

        for kolonneNavnExcel, parameterIndholdExcel in pVLParameter
        {
            for parameterNavn, parameterObj in this.parametre.OwnProps()
                if parameterObj.kolonneNavn = kolonneNavnExcel
                {
                    parameterObj.forventetIndhold := parameterIndholdExcel
                    parameterObj.iBrug := 1
                    parameterObj.kolonneNummer := pGyldigeKolonner.%kolonneNavnExcel%.KolonneNummer
                }

        }
    }

    setTjekketVognløb(pTjekketVognløb) {

        this.TjekketVognløb := pTjekketVognløb
    }

    getTjekketVognløb() {

        return this.TjekketVognløb
    }


    tjekForbudtVognløbsDato() {

        forbudteDatoer := ["23-12", "24-12", "31-12", "01-01"]

        for forbudtDato in forbudteDatoer
            if this.parametre.Vognløbsdato = forbudtDato
                Throw Error("Forbudt dato - " this.parametre.Vognløbsdato)

    }
}