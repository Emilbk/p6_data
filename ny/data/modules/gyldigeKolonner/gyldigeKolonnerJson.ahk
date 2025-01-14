#include ../../lib/cJson.ahk
class gyldigKolonneJson {

    static __New() {

        gyldigKolonneJson._parameter := Map()
        gyldigKolonneJson._parameter.default := [false, false, false, false, false]
        gyldigKolonneJson._parameter.CaseSense := 0
        static _jsonInd := json.Load(fileread(
            "c:\Users\nixVM\Documents\ahk\p6_data\ny\data\modules\gyldigeKolonner\gyldigeKolonner.Json"))

        for key, value in _jsonInd
            gyldigKolonneJson._parameter.set(key, value)
    }

    static data {
        get {
            return gyldigKolonneJson._parameter
        }

    }
    static maxParameterLængde(kolonneNavn) {
        return gyldigKolonneJson._parameter[kolonneNavn][1]
    }
    static maxArrayLængde(kolonneNavn) {
        return gyldigKolonneJson._parameter[kolonneNavn][2]
    }
    static erGyldigKolonne(kolonneNavn) {
        return gyldigKolonneJson._parameter[kolonneNavn][3]
    }
    static kolonneID(kolonneNavn) {
        return gyldigKolonneJson._parameter[kolonneNavn][4]
    }
    static erExcelKolonne(kolonneNavn) {
        return gyldigKolonneJson._parameter[kolonneNavn][5]
    }
}
