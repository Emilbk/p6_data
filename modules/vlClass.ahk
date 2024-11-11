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
    behandlVognløsbsdata(pVognløbsdata){
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


    danVognløbOutput(){

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
                vognløbOutput[outerIndex][ugedagArrayCount].setVognløbsDataTilIndlæsning(enkeltVognløbInput)
                vognløbOutput[outerIndex][ugedagArrayCount].setfejlLog(enkeltVognløbInput)
                vognløbOutput[outerIndex][ugedagArrayCount].tilIndlæsning.Vognløbsdato := ugedag
                vognløbOutput[outerIndex][ugedagArrayCount].tilIndlæsning.Ugedage := ""
            }
        }

        return vognløbOutput
    }

}

/** Repræsenterer et specifikt vognløb på en specifik dato (eller fast ugedag) */
class VognløbObj
{
    
    tjekkedeParametre := p6Parameter()
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
    
    setVognløbsDataTilIndlæsning(pVLParameter) {
        
        for vlKey, vlIndhold in pVLParameter
        {
            this.tilIndlæsning.%vlKey% := vlIndhold
        }
    }
    
    setTjekketVognløb(pTjekketVognløb){

        this.TjekketVognløb := pTjekketVognløb
    }

    getTjekketVognløb(){

        return this.TjekketVognløb
    }

    setFejlLog(pVlData)
    {
        this.fejlLog := fejlLogObj()
        this.fejlLog.setVognløbsnummerOgDato(pVlData)
    }
    test(){
        MsgBox this.tilIndlæsning.vognløbsnummer " - " this.tilIndlæsning.Vognløbsdato

        return
    }
    
}