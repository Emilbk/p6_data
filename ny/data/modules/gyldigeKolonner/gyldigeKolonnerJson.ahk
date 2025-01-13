#include ../../lib/cJson.ahk
class gyldigKolonneJson {
    static _parameter := json.Load(fileread("c:\Users\nixVM\Documents\ahk\p6_data\ny\data\modules\gyldigeKolonner\gyldigeKolonner.Json"))
    static data {
        get {
            return gyldigKolonneJson._parameter
        }

    }
    static maxArrayLængde(kolonneNavn){
        return gyldigKolonneJson._parameter[kolonneNavn][2]
    }
    static maxParameterLængde(kolonneNavn){
        return gyldigKolonneJson._parameter[kolonneNavn][1]
    }
    
    static erGyldigKolonne(kolonneNavn){
        return gyldigKolonneJson._parameter[kolonneNavn][3]
    }
}



