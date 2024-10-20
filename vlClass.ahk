#Requires AutoHotkey v2.0
#Include deepCopy.ahk

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

    Vognløb := Array()

    indhentVognløbsdata(pVlArray)
    {
        this.vlData := pVlArray

        return
    }

    opretVognløbForHverDato()
    {
        this.Vognløb := array()
        arrayCount := 0
        for ugedag in this.vlData["Ugedage"]
        {
            if ugedag = ""
                continue
            ugedag := Format("{:U}", ugedag)
            dp := DeepCopy(this.vlData)
            midlObj := dp()
            arrayCount += 1
            this.Vognløb.Push(midlObj)
            this.Vognløb[arrayCount]["Vognløbsdato"] := ugedag
            ; this.Vognløb.Set(ugedag, midlObj)
            ; this.Vognløb[ugedag]["Vognløbsdato"] := ugedag

            ; this.Vognløb[ugedag] := midlObj()

            ; this.Vognløb[a_index].Set(ugedag, midlObj)
            ; this.Vognløb[a_index].Set("Vognløbsdato", ugedag)


        }
        return
    }

    ; eksempelDataStruktur() {

    ; }

    ; aeksempelDatastruktur() {

    ;     for vlsomMapKey, vlSomMap in this.Vognløb["Fr"]
    ;         for celleArraySomMapValue in vlSomMap
    ;             MsgBox vlSomMap["Vognløbsnummer"][1] ": " kolonneNavnSomMapKey " - " endeligCelleVærdi
    ;     return
    ; }
}