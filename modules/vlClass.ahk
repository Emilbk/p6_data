#Requires AutoHotkey v2.0
#Include ../Lib/deepCopy.ahk

class VognløbConstructor {

    vognløbInput := ""
    vognløbOutput := Array()

    setVognløbsdata(pVognløbsData) {

        this.vognløbInput := pVognløbsData
        this.vognløbArrayPrVognløb()
        this.vognløbArrayPrUgedag()
        return
    }

    getVognløbsdata()
    {
        return this.vognløbOutput
    }


    vognløbArrayPrVognløb() {
        vognløbsarray := this.vognløbInput
        vognløbOutput := this.vognløbOutput

        for vognløb in vognløbsarray
            vognløbOutput.push(Array())

        this.vognløbOutput := vognløbOutput
        return
    }

    vognløbArrayPrUgedag() {
        vognløbsarray := this.vognløbInput
        vognløbOutput := this.vognløbOutput

        for vognløb in vognløbsarray
        {
            ugedagArrayCount := 0
            outerIndex := A_Index
            for ugedag in vognløb["Ugedage"]
            {
                if ugedag = ""
                    continue
                ugedagArrayCount += 1
                ugedag := Format("{:U}", ugedag)
                vognløb["Vognløbsdato"] := ugedag
                vognløbOutput[outerIndex].push(Array)
                vognløbOutput[outerIndex][ugedagArrayCount] := VognløbObj()
                vognløbOutput[outerIndex][ugedagArrayCount].setVognløb(vognløb)
                vognløbOutput[outerIndex][ugedagArrayCount].tilIndlæsning.Vognløbsdato := ugedag
                vognløbOutput[outerIndex][ugedagArrayCount].tilIndlæsning.Ugedage := ""
            }
        }

        return vognløbOutput
    }
    ; static setKolonneNavnOgNummer(pKolonneNavnOgNummer){

    ;     VognløbConstructor.kolonneNavnOgNummer := pKolonneNavnOgNummer
    ; }
}

/** Repræsenterer et specifikt vognløb på en specifik dato (eller fast ugedag) */
class VognløbObj
{
    ; /**
    ;  * @param {Object} excelObjP6Data obj fra class excelobjP6Data
    ;  * @property {Array} this.vlDataIndlæst asdsad
    ;  */
    ; __New(excelObjP6Data) {

    ;     this.excelDataTilIndlæsning := excelObjP6Data.getData()
    ;     this.VlDataIndlæst := Array(Map())

    ;     /** @type {Array}  */
    ;     this.vlDataTilIndlæsningArray := Array()

    ;     for vl in this.excelDataTilIndlæsning
    ;         this.vlDataTilIndlæsningArray.push(vl)
    ; }
    
    tilIndlæsning := Object()

    tilIndlæsning.Budnummer := ""
    tilIndlæsning.Vognløbsnummer := ""
    tilIndlæsning.Kørselsaftale := ""
    tilIndlæsning.Styresystem := ""
    tilIndlæsning.Startzone := ""
    tilIndlæsning.Slutzone := ""
    tilIndlæsning.Hjemzone := ""
    tilIndlæsning.MobilnrChf := ""
    tilIndlæsning.Vognløbskategori := ""
    tilIndlæsning.Planskema := ""
    tilIndlæsning.Økonomiskema := ""
    tilIndlæsning.Statistikgruppe := ""
    tilIndlæsning.Vognløbsnotering := ""
    tilIndlæsning.Starttid := ""
    tilIndlæsning.Sluttid := ""
    tilIndlæsning.UndtagneTransporttyper := ""
    tilIndlæsning.Vognløbsdato := ""
    tilIndlæsning.Ugedage := ""

    setVognløb(vlData) {
        for vlKey, vlIndhold in vlData
        {
            this.tilIndlæsning.%vlkey% := vlIndhold
        }
    }
    test(){
        MsgBox this.tilIndlæsning.vognløbsnummer " - " this.tilIndlæsning.Vognløbsdato

        return
    }
    
}