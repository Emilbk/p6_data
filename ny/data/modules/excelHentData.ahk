#Include includeModules.ahk

; TODO
; Hvordan dobbelte parametre?

class _excelHentData {

    __New(pExcelFil, pArkNavnEllerNummer, excelApp := "") {
        this.app := excelApp ? excelApp : ComObject("Excel.Application")
        this.fil := { path: pExcelFil }
        this.sheet := { navn: pArkNavnEllerNummer, data: 0 }

        this._setFilVariabler()
    }

    _setFilVariabler() {

        SplitPath(this.fil.path, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.fil.navn := varFilNavn
        this.fil.navnIngenExt := varFilNavnUdenExtension
        this.fil.dir := varFilDir
    }
    _åbenWorkbookReadonly() {
        this.sheet.aktivWorkbook := this.app.Workbooks.open(this.fil.path, "ReadOnly" = true)
        this.sheet.aktivWorksheet := this.sheet.aktivWorkbook.Sheets(this.sheet.navn)
    }

    _indlæsAktivRangeTilArray() {
        this.sheet.SafeArray := this.sheet.aktivWorksheet.usedrange.value
    }

    _quit() {
        this.app.quit()
    }

    getDataArray {
        get {

            if !this.sheet.data {
                this._åbenWorkbookReadonly()
                this._indlæsAktivRangeTilArray()
            }

            if this.app
                this._quit()

            return this._konverterTil2dArray()

        }
    }
    _konverterTil2DArray() {

        outputArray := []
        safeArray := this.sheet.SafeArray
        maxRækker := safeArray.MaxIndex(1)
        maxKolonner := safeArray.MaxIndex(2)

        loop maxRækker {
            rækkeIndex := A_Index
            outputArray.Push(Array())
            loop maxKolonner {
                kolonneIndex := A_Index
                aktivCelle := safeArray[rækkeIndex, kolonneIndex]
                if IsFloat(aktivCelle)
                    aktivCelle := String(Floor(aktivCelle))
                outputArray[rækkeIndex].Push(aktivCelle)
            }
        }

        return outputArray
    }
}
class _excelStrukturerData {

    __New(excelArray, parameterFactory) {
        this.excelArray := excelArray
        this.parameterFactory := parameterFactory

    }

    danKolonneNavneOgNummer() {
        excelArray := this.excelArray
        kolonneNavne := { gyldigeKolonner: map(), ugyldigeKolonner: map() }

        dataVerificering := _excelVerificerData
        for rækkeIndex, kolonne in excelArray
            for kolonneIndex, kolonneNavn in kolonne {
                if (rækkeIndex = 1 and dataVerificering.erGyldigKolonne(kolonneNavn))
                    kolonneNavne.gyldigeKolonner.Set(kolonneNavn, kolonneIndex)
                if (rækkeIndex = 1 and !dataVerificering.erGyldigKolonne(kolonneNavn))
                    kolonneNavne.ugyldigeKolonner.Set(kolonneNavn, kolonneIndex)
            }

        kolonneNavne.gyldigeKolonner["Ugedage"] := Array()
        kolonneNavne.gyldigeKolonner["UndtagneTransporttyper"] := Array()
        kolonneNavne.gyldigeKolonner["KørerIkkeTransporttyper"] := Array()

        for kolonne in excelArray[1] {
            if kolonne = "Ugedage"
                kolonneNavne.gyldigeKolonner["Ugedage"].Push(A_Index)
            if kolonne = "UndtagneTransporttyper"
                kolonneNavne.gyldigeKolonner["UndtagneTransporttyper"].Push(A_Index)
            if kolonne = "KørerIkkeTransporttyper"
                kolonneNavne.gyldigeKolonner["KørerIkkeTransporttyper"].Push(A_Index)
        }
        return kolonneNavne
    }

    ; lav factory
    danRækkeArray() {
        excelArray := this.excelArray
        raekkeArray := Array()
        dataVerificering := _excelVerificerData
        outputArray := Array()
        gyldigeParametre := parameterGyld.data

        for rækkeindex, raekke in excelarray {
            raekkearray.push(map())
            raekkeArray[rækkeindex].set("Vognløbsdato", this.parameterFactory.forKolonneNavn(excelParameter({ parameterNavn: "Vognløbsdato"})))
            for kolonneindex, celle in raekke {
                parameterNavn := excelarray[1][kolonneindex]
                excelPar := excelParameter({
                    parameterNavn: parameternavn,
                    parameterIndhold: celle,
                    kolonneIndex: kolonneindex,
                    kolonneNavn: parameterNavn,
                    rækkeIndex: rækkeindex
                })
                if dataVerificering.erGyldigKolonne(parameterNavn) {
                    raekkearray[rækkeindex].set(parameterNavn, this.parameterFactory.forKolonneNavn(excelPar))
                }
            }
        }
        raekkeArray.RemoveAt(1)

        parameterFactory.reset()
        return raekkeArray

    }

}

class _excelVerificerData {
    static _gyldigeKolonner := gyldigeKolonner.data
    static _ugyldigeKolonner := Map()

    static _verificerKolonner(pKolonner) {
        for kolonne in pKolonner
            if !_excelVerificerData._gyldigeKolonner.has(kolonne)
                _excelVerificerData._ugyldigeKolonner.Set(kolonne, A_Index)
            else
                _excelVerificerData._gyldigeKolonner[kolonne] := true
    }

    static ugyldigeKolonner[pKolonner] {
        get {
            _excelVerificerData._verificerKolonner(pKolonner)
            return _excelVerificerData._ugyldigeKolonner
        }
    }

    static gyldigeKolonner[pKolonner] {
        get {
            _excelVerificerData._verificerKolonner(pKolonner)
            return _excelVerificerData._gyldigeKolonner
        }
    }

    static erGyldigKolonne(kolonneNavn) {

        if _excelVerificerData._gyldigeKolonner.Has(kolonneNavn)
            return true

    }
    ;; TODO
    static erGyldigArrayLængde(pParameterObj) {

        gyldigeParametre := parameterGyld.data
        testParameter := pParameterObj

        testParameterNavn := testParameter.data["parameterNavn"]
        testParameterIndholdArray := testParameter.data["forventetIndholdArray"]

        if testParameterIndholdArray
            if testParameterIndholdArray.length > gyldigeParametre[testParameterNavn]["maxArray"] {
                testParameter.data["fejl"] := 1
                testParameter.data["fejlBesked"] := "for mange kolonner i kategori"
                testParameter.data["fejlParameterArray"] := testParameter.data["kolonneNavn"]
                testParameter.data["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxArray"]
            }
    }

}

class excelDataBehandler {

    __New(pInputArray, parameterObj) {
        this.dataArray := pInputArray
        this.parameterObj := parameterObj

        this.rækkeArray := _excelStrukturerData(this.dataArray, this.parameterObj).danRækkeArray()
        this.kolonner := _excelStrukturerData(this.dataArray, this.parameterObj).danKolonneNavneOgNummer()

    }

    test(){

    }
    behandledeRækker {
        get {
            return this.rækkeArray
        }
    }
    gyldigeKolonner {
        get {
            return this.kolonner
        }
    }

}

class parameterFactory {
    static ugedageRække := Map()
    static UgedageInstance := ""
    static KørerIkkeTransporttyperRække := Map()
    static KørerIkkeTransporttyperInstance := ""
    static UndtagneTransporttyperRække := Map()
    static UndtagneTransporttyperInstance := ""

    static reset(){
        parameterFactory.ugedageRække := map()
        parameterFactory.UndtagneTransporttyperRække := map()
        parameterFactory.KørerIkkeTransporttyperRække := map()
    }
    static forKolonneNavn(excelParametre) {


        switch excelParametre.kolonneNavn {
            case "Ugedage":
            {
                if !parameterFactory.ugedageRække.Has(excelParametre.rækkeIndex) {
                    parameterFactory.ugedageRække.Set(excelParametre.rækkeIndex, 1)
                    parameterFactory.UgedageInstance := parameterUgedage(excelParametre)
                    return parameterFactory.UgedageInstance
                }
                else {
                    parameterFactory.UgedageInstance.tilføjParametreTilArray(parameterFactory.UgedageInstance, excelParametre)
                    return parameterFactory.UgedageInstance
                }
            }
            case "KørerIkkeTransporttyper":
                if !parameterFactory.KørerIkkeTransporttyperRække.Has(excelParametre.rækkeIndex) {
                    parameterFactory.KørerIkkeTransporttyperRække.Set(excelParametre.rækkeIndex, 1)
                    parameterFactory.KørerIkkeTransporttyperInstance := parameterTransportType(excelParametre)
                    return parameterFactory.KørerIkkeTransporttyperInstance
                }
                else {
                    parameterFactory.KørerIkkeTransporttyperInstance.tilføjParametreTilArray(parameterFactory.KørerIkkeTransporttyperInstance,
                        excelParametre)
                    return parameterFactory.KørerIkkeTransporttyperInstance
                }
            case "UndtagneTransporttyper":
                if !parameterFactory.UndtagneTransporttyperRække.Has(excelParametre.rækkeIndex) {
                    parameterFactory.UndtagneTransporttyperRække.Set(excelParametre.rækkeIndex, 1)
                    parameterFactory.UndtagneTransporttyperInstance := parameterTransportType(excelParametre)
                    return parameterFactory.UndtagneTransporttyperInstance
                }
                else {
                    parameterFactory.UndtagneTransporttyperInstance.tilføjParametreTilArray(parameterFactory.UndtagneTransporttyperInstance,
                        excelParametre)
                    return parameterFactory.UndtagneTransporttyperInstance
                }
                return parameterTransportType(excelParametre)
            case "Starttid":
                return parameterKlokkeslæt(excelParametre)
            case "Sluttid":
                return parameterKlokkeslæt(excelParametre)
            default:
                return parameterAlm(excelParametre)
        }
    }
}

;???????
class parameterInterface {
    __New(excelParametre) {

        ; throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }

    tilføjParametreTilArray(parameterOjb, excelParametre) {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }
    _forMangeTegnIParameter() {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
        }
    _ulovligtTegnIParameter() {

    }
    _danFejl(pFejlbesked) {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }
    _hvisArray(){
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
         
    }
    tjekGyldighed() {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }

    udfyldParameter(excelData) {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }

    forventet {
        set{

            throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
        }
        get{

            throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
        }
    }
}
class parameterAlm extends parameterInterface {
    __New(excelParametre) {
        this.data := parameterSkabelon().tomtParameterSæt
        this._hvisArray()
        this.udfyldParameter(excelParametre)
        this.tjekGyldighed()

    }

    _forMangeTegnIParameter() {

        if StrLen(this.data["forventetIndhold"]) > this.data["maxParameterLængde"] {
            this._danfejl(Format("For mange tegn i parameter `"{1}`". Nuværende {2}, maks {3}.", this.data[
                "forventetIndhold"], StrLen(this.data["forventetIndhold"]), this.data["maxParameterLængde"]))
            return
        }
    }
    _ulovligtTegnIParameter() {
        if RegExMatch(this.data["forventetIndhold"], "[\!\*\@]", &matchObj) {
            this._danFejl(Format("Ulovligt tegn (`"{1}`") i parameter.", matchObj[0]))
            return

        }

    }
    _danFejl(pFejlbesked) {
        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked

    }
    _hvisArray(){
         
    }
    tjekGyldighed() {
        this._forMangeTegnIParameter()
        this._ulovligtTegnIParameter()
    }

    udfyldParameter(excelData) {
        gyldigeParametre := parameterGyld.data

        this.data["kolonneNavn"] := exceldata.parameterNavn
        this.data["parameterNavn"] := exceldata.parameterNavn
        this.data["forventetIndhold"] := StrUpper(exceldata.parameterIndhold)
        this.data["kolonneNummer"] := exceldata.kolonneIndex
        this.data["maxParameterLængde"] := gyldigeParametre[exceldata.parameterNavn]["maxLængde"]
        this.data["maxArrayLængde"] := gyldigeParametre[exceldata.parameterNavn]["maxArray"]
    }

    forventet {
        get {
            return this.data["forventetIndhold"]
        }
        set {
            this.data["forventetIndhold"] := Value
        }
    }
}
class parameterArray extends parameterAlm {
    tilføjParametreTilArray(parameterOjb, excelParametre) {
        parameterOjb.data["forventetIndholdArray"].Push(StrUpper(excelParametre.parameterIndhold))
        parameterOjb.data["kolonneNummerArray"].Push(excelParametre.kolonneIndex)
    }
    _erOverMaxArray() {

        if this.data["forventetIndholdArray"].length > this.data["maxArrayLængde"] {
            this._danfejl(Format("For mange mange kolonner i kategori. Maks {1}, nuværende {2}", this.data[
                "maxArrayLængde"], this.data["forventetIndholdArray"].length))
            return true
        }

    }
    _forMangeTegnIParameter() {

        for tjekParameter in this.data["forventetIndholdArray"]
            if StrLen(tjekParameter) > this.data["maxLængde"] {
                this._danfejl(Format("For mange tegn i parameter {1} Nuværende {2}, maks {3}", tjekParameter, StrLen(
                    tjekParameter), this.data["maxLængde"]))
                return true
            }

    }
    _hvisArray(){
        this.data["forventetIndholdArray"] := Array()
        this.data["kolonneNummerArray"] := Array()
    }
    udfyldParameter(excelData) {
        gyldigeParametre := parameterGyld.data

        this.data["kolonneNavn"] := exceldata.parameterNavn
        this.data["parameterNavn"] := exceldata.parameterNavn
        this.data["forventetIndholdArray"].push(StrUpper(excelData.parameterIndhold))
        this.data["kolonneNummerArray"].push(excelData.kolonneIndex)
        
        this.data["maxParameterLængde"] := gyldigeParametre[exceldata.parameterNavn]["maxLængde"]
        this.data["maxArrayLængde"] := gyldigeParametre[exceldata.parameterNavn]["maxArray"]


    }

    tjekGyldighed() {
        throw Error("tjekGyldighed skal implementeres")
    }
    forventet {
        get {

            return this.data["forventetIndholdArray"]
        }
    }
    faktisk {

        set {

            this.data["faktiskIndholdArray"] := Value
        }
    }
}
class parameterUgedage extends parameterArray {
    _erKalenderdag(ugedag) {
        if RegExMatch(ugedag, "\w*\d\w*")
            return true
    }
    _erGyldigFastDag(ugedag) {

        gyldigeUgedage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]

        for gyldigUgedag in gyldigeUgedage
            if ugedag = gyldigeUgedage
                return true
    }
    _harSkråstreg(ugedag) {
        if InStr(ugedag, "/")
            return true
    }
    _erGyldigDato(ugedag) {

        if !this._harSkråstreg(ugedag)
            return false

        datoArr := StrSplit(ugedag, "/")
        if datoArr.Length != 3
            return false

        dag := datoArr[1]
        måned := datoArr[2]
        år := datoArr[3]

        testStamp := år . måned . dag
        return isTime(testStamp)

    }
    tjekGyldighed() {
        ugedage := this.data["forventetIndholdArray"]

        for ugedag in ugedage
            if !this._erKalenderdag(ugedag) {
                if !this._erGyldigFastDag(ugedag) {
                    this._danFejl(Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", ugedag))
                    return
                }
            }
            else if !this._erGyldigDato(ugedag) {
                this._danFejl(Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", ugedag))
                return
            }
    }
}
class parameterTransportType extends parameterArray {


    tjekGyldighed() {
        if this._erOverMaxArray()
            return
        if this._forMangeTegnIParameter()
            return
    }

}
class parameterKlokkeslæt extends parameterAlm {


    _tjekOgRensAsterisk() {

        if InStr(this.data["forventetIndhold"], "*") {
            this.data["forventetIndhold"] := SubStr(this.data["forventetIndhold"], 1, 5)
            this.data["sluttidspunktErNæsteDag"] := true

        }
    }

    _harKolon() {
        if !InStr(this.data["forventetIndhold"], ":") {
            this._danfejl(Format(
                "Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"
            ))
            return
        }
    }
    _korrektFormat() {

        strArr := StrSplit(this.data["forventetIndhold"], ":")

        if strArr.Length != 2 {
            this._danfejl(Format(
                "Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"
            ))
            return

        }

        time := strArr[1]
        minut := strArr[2]

        if !IsTime("20241212" time minut)
            this._danfejl(Format(
                "Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"
            ))

        return
    }


    tjekGyldighed() {
        this._tjekOgRensAsterisk()
        this._harKolon()
        this._korrektFormat()
    }


}

class parameterSkabelon {
    tomtParameterSæt {
        get {

            this.data := Map()
            this.data.Default := 0

            this.data["kolonneNavn"] := false
            this.data["kolonneNummer"] := false
            this.data["kolonneNummerArray"] := false
            this.data["parameterNavn"] := false
            this.data["forventetIndhold"] := false
            this.data["forventetIndholdArray"] := false
            this.data["faktiskIndhold"] := false
            this.data["faktiskIndholdArray"] := false
            this.data["fejl"] := false
            this.data["fejlBesked"] := false
            this.data["fejlParameterArray"] := false
            this.data["maxParameterLængde"] := false
            this.data["maxArrayLængde"] := false
            this.data["tidspunktErNæsteDag"] := false

            return this.data
        }
    }
}

/**
 * 
 * @param parameterObj {parameterNavn: string, parameterIndhold: string, kolonneNavn: stringOpt, kolonneIndex: stringOpt, rækkeIndex: stringOpt}
 */
class excelParameter {

    ; @param parameterObj {parameterNavn: string, parameterIndhold: string, kolonneNavn: stringOpt, kolonneIndex: stringOpt, rækkeIndex: stringOpt}
    __New(parameterObj) {
        this._parameterObj := parameterObj
        if !parameterObj.parameterNavn
            throw UnsetError('parameterNavn er unset')
    }

    kolonneIndex {
        get {

            return this._parameterObj.HasOwnProp("kolonneIndex") ? this._parameterObj.kolonneIndex : false
        }
    }
    kolonneNavn {
        get {

            return this._parameterObj.HasOwnProp("kolonneNavn") ? this._parameterObj.kolonneNavn : false
        }
    }
    rækkeIndex {
        get {

            return this._parameterObj.HasOwnProp("rækkeIndex") ? this._parameterObj.rækkeIndex : false
        }
    }
    parameterIndhold {
        get {

            return this._parameterObj.HasOwnProp("parameterIndhold") ? this._parameterObj.parameterIndhold : false
        }
    }
    parameternavn {
        get {

            return this._parameterObj.HasOwnProp("parameternavn") ? this._parameterObj.parameternavn : false
        }
    }

}
