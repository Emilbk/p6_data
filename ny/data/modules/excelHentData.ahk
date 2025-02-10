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

    ; TODO lav dynamisk array-inddeling istf. hardcoded på kolonnenavn?
    danKolonneNavneOgNummer() {
        excelArray := this.excelArray
        kolonneNavne := { gyldigeKolonner: Array(), ugyldigeKolonner: Array() }
        r := []

        dataVerificering := gyldigKolonneJson
        for rækkeIndex, kolonne in excelArray {
            r.Push([])
            for kolonneIndex, celle in kolonne {
                if rækkeIndex = 1 {
                    if (rækkeIndex = 1 and dataVerificering.erGyldigKolonne(celle))
                        kolonneNavne.gyldigeKolonner.push([celle, kolonneIndex])
                    else
                        kolonneNavne.ugyldigeKolonner.push([celle, kolonneIndex])
                }
                else
                    r[rækkeIndex].Push({ kolonneIndex: kolonneIndex, ParameterIndhold: celle, kolonneNavn: excelArray[
                        1][kolonneIndex] })

            }
        }
        r.RemoveAt(1)
        kolonnerMedArray := []
        for rIndex, rækker in r {
            kolonnerMedArray.Push([])
            kolonnerMedArray[rIndex].Push({ kolonneIndex: [], parameterIndhold: [], kolonneNavn: "" }, { kolonneIndex: [],
                parameterIndhold: [], kolonneNavn: "" }, { kolonneIndex: [], parameterIndhold: [], kolonneNavn: "" })
            for celle in rækker
                switch celle.kolonneNavn {
                    case "Ugedage":
                    {
                        kolonnerMedArray[rIndex][1].kolonneIndex.Push(celle.kolonneIndex)
                        kolonnerMedArray[rIndex][1].parameterIndhold.Push(celle.ParameterIndhold)
                        kolonnerMedArray[rIndex][1].kolonneNavn := celle.kolonneNavn
                    }
                    case "UndtagneTransportTyper":
                        kolonnerMedArray[rIndex][2].kolonneIndex.Push(celle.kolonneIndex)
                        kolonnerMedArray[rIndex][2].parameterIndhold.Push(celle.ParameterIndhold)
                        kolonnerMedArray[rIndex][2].kolonneNavn := celle.kolonneNavn
                    case "KørerIkkeTransportTyper":
                        kolonnerMedArray[rIndex][3].kolonneIndex.Push(celle.kolonneIndex)
                        kolonnerMedArray[rIndex][3].parameterIndhold.Push(celle.ParameterIndhold)
                        kolonnerMedArray[rIndex][3].kolonneNavn := celle.kolonneNavn
                    default:
                    {
                        kolonnerMedArray[rIndex].Push({ kolonneIndex: celle.kolonneIndex, parameterIndhold: celle.ParameterIndhold,
                            kolonneNavn: celle.kolonneNavn })

                    }

                }
        }

        return kolonnerMedArray
    }

    danRækkeArray() {
        this.kolonneInfo := this.danKolonneNavneOgNummer()
        this.rækkeMap := []

        output := []

        for rIndex, r in this.kolonneInfo {
        {
            output.Push(map())
            output[rIndex].CaseSense := 0
        }
            for kIndex, parameter in r
                output[rIndex].set(parameter.kolonneNavn, parameterFactory.forExcelParameter(excelParameter(parameter)))

            output[rIndex].set("Vognløbsdato", parameterFactory.forExcelParameter(excelParameter({ kolonneNavn: "Vognløbsdato" })))

        }
        return output
    }

}

class excelDataBehandler {

    __New(pInputArray, parameterObj) {
        this.dataArray := pInputArray
        this.parameterObj := parameterObj

        this.rækkeArray := _excelStrukturerData(this.dataArray, this.parameterObj).danRækkeArray()
        this.kolonner := _excelStrukturerData(this.dataArray, this.parameterObj).danKolonneNavneOgNummer()

    }

    test() {

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
    static forExcelParameter(excelParametre) {

        switch excelParametre.kolonneNavn {
            case "Ugedage":
                return parameterUgedage(excelParametre)
            case "KørerIkkeTransportTyper":
                return parameterTransportType(excelParametre)
            case "UndtagneTransportTyper":
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
    _hvisArray() {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)

    }
    _upperCase() {

        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }
    _danFejlObj() {

    }
    tjekGyldighed() {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }

    udfyldParameter(excelData) {
        throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
    }

    forventet {
        set {

            throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
        }
        get {

            throw Error(Format("Funktion {1} skal implementeres", A_ThisFunc), A_ThisFunc)
        }
    }

}
class parameterAlm extends parameterInterface {
    __New(excelParametre := "") {
        this.data := parameterSkabelon().tomtParameterSæt
        this._dataVerificering := gyldigKolonneJson
        this.data.kolonneNavn := excelParametre.kolonneNavn
        this.data.kolonneIndex := excelParametre.KolonneIndex
        this._hvisArray()
        this.setParametre(excelParametre)
    }

    maxParameterLængde {
        get {
            return this._dataVerificering.maxParameterLængde(this.data.kolonnenavn)
        }
    }
    maxArrayLængde {
        get {
            return this._dataVerificering.maxArrayLængde(this.data.kolonnenavn)
        }
    }
    setParametre(excelParametre) {

        this.udfyldParameter(excelParametre)
        this._danFejlObj()
        this._upperCase()
        this.tjekGyldighed()
    }

    _forMangeTegnIParameter() {

        if StrLen(this.data["forventetIndhold"]) > this.data["maxParameterLængde"] {
            this.fejlObj["fejlBesked"] := (Format("For mange tegn i parameter `"{1}`". Nuværende {2}, maks {3}.",
                this.fejlObj["forventetIndhold"], StrLen(this.data["forventetIndhold"]), this.data["maxParameterLængde"]))
                fejlRegister.registrerFejl(this.fejlObj)
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
        this.fejlObj["Status"] := 1
        this.fejlObj["fejlBesked"] := pFejlbesked

    }
    _danFejlObj() {

        this.fejlObj := map(
            "parameterNavn", this.data["parameterNavn"],
            "kolonneNavn", this.data["kolonneNavn"],
            "forventetIndhold", this.data["forventetIndhold"],
            "faktiskIndhold", this.data["faktiskIndhold"],
            "maxParameterLængde", this.data["maxParameterLængde"],
            "fejlBesked", "",
            "kolonneNummer", this.data["kolonneNummer"]
        )
    }
    _hvisArray() {

    }
    _upperCase() {
        for undtagetKolonne in ["Vognløbsnotering", "VognmandLinie1", "VognmandLinie2", "VognmandLinie3",
            "VognmandLinie4"]
            if this.data["kolonneNavn"] = undtagetKolonne
                return

        this.data["forventetIndhold"] := StrUpper(this.data["forventetIndhold"])

    }
    tjekGyldighed() {
        this._forMangeTegnIParameter()
        this._ulovligtTegnIParameter()
    }

    ; parameterAlm
    udfyldParameter(excelData) {
        gyldigeParametre := gyldigKolonneJson
        
        this.data["kolonneNavn"] := exceldata.kolonneNavn
        this.data["maxParameterLængde"] := gyldigeParametre.maxParameterLængde(exceldata.kolonneNavn)
        this.data["maxArrayLængde"] := gyldigeParametre.maxArrayLængde(exceldata.kolonneNavn)
        this.data["parameterNavn"] := exceldata.parameterNavn
        this.data["forventetIndhold"] := exceldata.parameterIndhold
        this.data["kolonneNummer"] := exceldata.kolonneIndex
    }

    forventet {
        get {
            return this.data["forventetIndhold"]
        }
        set {
            this.data["forventetIndhold"] := Value
        }

    }
    faktisk {

        set {

            this.data["faktiskIndhold"] := Value

            if value != this.data["forventetIndhold"]
                throw Error("ikke ens")

        }

        get => this.data["faktiskIndhold"]
    }
}
class parameterArray extends parameterAlm {
    setParametre(excelParametre) {

        this.udfyldParameter(excelParametre)
        this._danFejlObj()
        this.tjekGyldighed()
    }
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
                this._danfejl(Format("For mange tegn i parameter {1} Nuværende {2}, maks {3}", tjekParameter,
                    StrLen(
                        tjekParameter), this.data["maxLængde"]))
                return true
            }

    }
    _hvisArray() {
        this.data["forventetIndholdArray"] := Array()
        this.data["kolonneNummerArray"] := Array()
    }

    _stringUpperArray(par) {
        for ind, str in par
            par[ind] := StrUpper(str)

        return par
    }
    _danFejlObj() {

        this.fejlObj := map(
            "parameterNavn", this.data["parameterNavn"],
            "kolonneNavn", this.data["kolonneNavn"],
            "forventetIndhold", this.data["forventetIndhold"],
            "faktiskIndhold", this.data["faktiskIndhold"],
            "maxParameterLængde", this.data["maxParameterLængde"],
            "fejlBesked", "",
            "kolonneNummer", this.data["kolonneNummer"]
        )
    }
    ; parameterArr
    udfyldParameter(exceldata) {
        gyldigeParametre := gyldigKolonneJson

        this.data["kolonneNavn"] := exceldata.kolonneNavn
        this.data["maxParameterLængde"] := gyldigeParametre.maxParameterLængde(exceldata.kolonneNavn)
        this.data["maxArrayLængde"] := gyldigeParametre.maxArrayLængde(exceldata.kolonneNavn)
        this.data["parameterNavn"] := exceldata.parameterNavn
        this.data["forventetIndholdArray"] := this._stringUpperArray(exceldata.parameterIndhold)
        this.data["kolonneNummerArray"] := exceldata.kolonneIndex

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
                    {
                    }
                    this._danFejl(Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", ugedag))
                    return
                }
            }
            else if !this._erGyldigDato(ugedag) {
                this._danFejl(Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.",
                    ugedag))
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
            this.data.CaseSense := 0

            this.data["kolonneNavn"] := false
            this.data["kolonneNummer"] := false
            this.data["kolonneNummerArray"] := false
            this.data["parameterNavn"] := false
            this.data["forventetIndhold"] := false
            this.data["forventetIndholdArray"] := false
            this.data["faktiskIndhold"] := false
            this.data["faktiskIndholdArray"] := false
            this.data["fejl"] := { status: false, fejlbesked: false, iFunc: false }
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
        if !parameterObj.kolonneNavn
            throw UnsetError('parameterNavn er unset')
    }

    kolonneNavn {
        get {

            return this._parameterObj.HasOwnProp("kolonneNavn") ? this._parameterObj.kolonneNavn : false
        }
    }
    kolonneIndex {
        get {

            return this._parameterObj.HasOwnProp("kolonneIndex") ? this._parameterObj.kolonneIndex : false
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

            return this._parameterObj.HasOwnProp("parameterNavn") ? this._parameterObj.parameternavn : false
        }
    }
    maxArrayLængde {
        get {

            return this._parameterObj.HasOwnProp("maxArrayLængde") ? this._parameterObj.maxArrayLængde : false
        }
    }
    maxParameterLængdeLængde {
        get {

            return this._parameterObj.HasOwnProp("maxParameterLængdeLængde") ? this._parameterObj.maxParameterLængdeLængde :
                false
        }
    }

}
