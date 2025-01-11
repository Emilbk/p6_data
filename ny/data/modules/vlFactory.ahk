#Include includeModules.ahk
; TODO dobbelte parametre skal opreetsse
class vlFactory {

    static udrulVognløb(pDataArray) {
        vlArray := []

        for vlIndex, vl in pDataArray {
            {

                vlArray.Push([])
                vlArray[vlIndex].vognløbsnummer := vl["Vognløbsnummer"].forventet
                ugedageArray := vl["Ugedage"].forventet
                for ugedagIndex, ugedag in ugedageArray {
                    dc := DeepCopy(vl)
                    vlKopi := dc()
                    vlArray[vlIndex].push(vlKopi)
                    vlArray[vlIndex][ugedagIndex]["Vognløbsdato"].forventet := ugedag
                }

            }
        }
        return vlArray
    }   

}

