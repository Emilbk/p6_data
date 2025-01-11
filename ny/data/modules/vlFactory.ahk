#Include includeModules.ahk
; TODO dobbelte parametre skal opreetsse
class vlFactory {

    static udrulVognløb(pDataArray) {
        vlArray := []

        for vlIndex, vl in pDataArray {
            {

                vlArray.Push([])
                ugedageArray := vl["Ugedage"].forventet
                for ugedagIndex, ugedag in ugedageArray {
                    dc := DeepCopy(vl)
                    vlKopi := dc()
                    vlArray[vlIndex].push(vognløb(vlkopi))
                    vlArray[vlIndex][ugedagIndex].parametre["Vognløbsdato"].forventet := ugedag
                    vlArray[vlIndex].master := vognløb(vlKopi)

                }

            }
        }
        return vlArray
    }

}
