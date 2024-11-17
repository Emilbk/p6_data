#Requires AutoHotkey v2.0
#Include ../Lib/deepCopy.ahk

class VognløbConstructor {
    __New(pExcelInput) {

        this.behandlVognløsbsdata(pExcelInput)
    }

    vognløbInput := ""
    vognløbOutput := Array()

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
            vognløbOutput.push(Array())

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
            for ugedag in enkeltVognløbInput["Ugedage"]
            {
                if ugedag = ""
                    continue
                ugedagArrayCount += 1
                ugedag := Format("{:U}", ugedag)
                enkeltVognløbInput["Vognløbsdato"] := ugedag
                vognløbOutput[outerIndex].push(Array)
                vognløbOutput[outerIndex][ugedagArrayCount] := VognløbObj()
                aktueltVognløb := vognløbOutput[outerIndex][ugedagArrayCount]

                aktueltVognløb.setVognløbsDataTilIndlæsning(enkeltVognløbInput)
                aktueltVognløb.udfyldundtagneTransportTyperArray()
                aktueltVognløb.udfyldKørerIkkeTransporttyperArray()

                ; vognløbOutput[outerIndex][ugedagArrayCount].setfejlLog(enkeltVognløbInput)
                ; vognløbOutput[outerIndex][ugedagArrayCount].setfejlLog(enkeltVognløbInput)
                aktueltVognløb.parametre.Vognløbsdato.forventetIndhold := ugedag
                aktueltVognløb.tjekSlutTidOverMidnat()

                aktueltVognløb.parametre.Ugedage := ""
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

    setFejlLog(pFejlObj) {

        this.fejlLog := pFejlObj
        this.fejlLog.fundetFejl := 1

    }
    udfyldUndtagneTransportTyperArray() {

        antalTransportTyper := this.parametre.UndtagneTransporttyper.forventetIndhold.Length
        ønsketAntalTransportTyper := 20

        while this.parametre.UndtagneTransporttyper.forventetIndhold.Length != ønsketAntalTransportTyper
            this.parametre.UndtagneTransporttyper.forventetIndhold.push(A_Space)

        for index, transportType in this.parametre.UndtagneTransporttyper.forventetIndhold
            if transportType = ""
                this.parametre.UndtagneTransporttyper.forventetIndhold[index] := A_Space
    }

    udfyldKørerIkkeTransporttyperArray() {

        antalTransportTyper := this.parametre.KørerIkkeTransporttyper.forventetIndhold.Length
        ønsketAntalTransportTyper := 10

        while this.parametre.KørerIkkeTransporttyper.forventetIndhold.Length != ønsketAntalTransportTyper
            this.parametre.KørerIkkeTransporttyper.forventetIndhold.push(A_Space)

        for index, transportType in this.parametre.KørerIkkeTransporttyper.forventetIndhold
            if transportType = ""
                this.parametre.KørerIkkeTransporttyper.forventetIndhold[index] := A_Space
    }

    tjekSlutTidOverMidnat() {

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

                this.parametre.VognløbsdatoSlut.forventetIndhold := slutDato

                return

            }


        if InStr(slutTid, "*")
        {
            slutTid := SubStr(slutTid, 1, 5)
            if arrayPos < 7
                this.parametre.VognløbsdatoSlut.forventetIndhold := fasteDageArray[arrayPos + 1]
            else
                this.parametre.VognløbsdatoSlut.forventetIndhold := fasteDageArray[1]
        }
        else
            this.parametre.VognløbsdatoSlut.forventetIndhold := fasteDageArray[arrayPos]
    }

    setVognløbsDataTilIndlæsning(pVLParameter) {

        for vlKey, vlIndhold in pVLParameter
        {
            this.parametre.%vlKey%.forventetIndhold := vlIndhold
            this.parametre.%vlKey%.iBrug := 1
        }
    }

    setTjekketVognløb(pTjekketVognløb) {

        this.TjekketVognløb := pTjekketVognløb
    }

    getTjekketVognløb() {

        return this.TjekketVognløb
    }

    ; setFejlLog(pVlData)
    ; {
    ;     this.fejlLog := fejlLogObj()
    ;     this.fejlLog.setVognløbsnummerOgDato(pVlData)
    ; }
    test() {
        MsgBox this.tilIndlæsning.vognløbsnummer " - " this.tilIndlæsning.Vognløbsdato

        return
    }

}