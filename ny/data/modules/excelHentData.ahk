#Include includeModules.ahk


; TODO
; anden parameterGyldighed objec
; Hvordan dobbelte parametre?
; parameterFactory?


mock := excelMock.excelDataUgyldigFlere
excelPath := "C:\Users\nixVM\Documents\ahk\p6_data\data\test\assets\VLMock.xlsx"
excelArk := 1

class _excelHentData {

    __New(pExcelFil, pArkNavnEllerNummer, excelApp := "") {
        if excelApp
            this.app := excelApp
        else this.app := ComObject("Excel.Application")
        this.fil := {}
        this.excel := {}
        this.fil.path := pExcelFil
        this.excel.arkNavn := pArkNavnEllerNummer

        this._setFilVariabler()
    }

    _setFilVariabler() {

        SplitPath(this.fil.path, &varFilNavn, &varFilDir, , &varFilNavnUdenExtension)
        this.fil.navn := varFilNavn
        this.fil.navnIngenExt := varFilNavnUdenExtension
        this.fil.dir := varFilDir
    }
    _åbenWorkbookReadonly() {
        this.excel.aktivWorkbook := this.app.Workbooks.open(this.fil.path, "ReadOnly" = true)
        this.excel.aktivWorksheet := this.excel.aktivWorkbook.Sheets(this.excel.arkNavn)
    }

    _indlæsAktivRangeTilArray() {
        this.excel.SafeArray := this.excel.aktivWorksheet.usedrange.value
    }

    _quit() {
        this.app.quit()
    }

    getDataArray {
        get {
            this._åbenWorkbookReadonly()
            this._indlæsAktivRangeTilArray()

            outputArray := []
            safeArray := this.excel.SafeArray
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

}

class _excelStrukturerData {

    __New(excelArray, parameterObj) {
        this.excelArray := excelArray
        this.parameterObj := parameterObj

    }

    danKolonneNavneOgNummer() {
        excelArray := this.excelArray
        kolonneNavne := { gyldigeKolonner: map(), ugyldigeKolonner: map() }

        dataVerificering := _excelVerificerData
        for rækkeIndex, kolonne in excelArray
            for kolonneIndex, kolonneNavn in kolonne
            {
                if (rækkeIndex = 1 and dataVerificering.erGyldigKolonne(kolonneNavn))
                    kolonneNavne.gyldigeKolonner.Set(kolonneNavn, kolonneIndex)
                if (rækkeIndex = 1 and !dataVerificering.erGyldigKolonne(kolonneNavn))
                    kolonneNavne.ugyldigeKolonner.Set(kolonneNavn, kolonneIndex)
            }

        kolonneNavne.gyldigeKolonner["Ugedage"] := Array()
        kolonneNavne.gyldigeKolonner["UndtagneTransporttyper"] := Array()
        kolonneNavne.gyldigeKolonner["KørerIkkeTransporttyper"] := Array()

        for kolonne in excelArray[1]
        {
            if kolonne = "Ugedage"
                kolonneNavne.gyldigeKolonner["Ugedage"].Push(A_Index)
            if kolonne = "UndtagneTransporttyper"
                kolonneNavne.gyldigeKolonner["UndtagneTransporttyper"].Push(A_Index)
            if kolonne = "KørerIkkeTransporttyper"
                kolonneNavne.gyldigeKolonner["KørerIkkeTransporttyper"].Push(A_Index)
        }
        return kolonneNavne
    }


    danRækkeArray() {
        excelArray := this.excelArray
        raekkeArray := Array()
        dataVerificering := _excelVerificerData
        outputArray := Array()
        gyldigeParametre := parameterGyld.data

        for rækkeindex, raekke in excelarray
        {
            raekkearray.push(map())
            for kolonneindex, celle in raekke
            {
                parameterNavn := excelarray[1][kolonneindex]
                if dataVerificering.erGyldigKolonne(parameterNavn)
                {
                    raekkearray[rækkeindex].set(parameterNavn, this.parameterObj.forKolonneNavn(parameterNavn))
                    raekkearray[rækkeindex][parameterNavn].data["kolonneNavn"] := parameterNavn
                    raekkearray[rækkeindex][parameterNavn].data["parameterNavn"] := parameterNavn
                    raekkearray[rækkeindex][parameterNavn].data["forventetIndhold"] := celle
                    raekkearray[rækkeindex][parameterNavn].data["kolonneNummer"] := kolonneindex
                    raekkearray[rækkeindex][parameterNavn].data["maxParameterLængde"] := gyldigeParametre[parameterNavn]["maxLængde"]

                }
            }
        }
        ; TODO
        ; kopier dobbelt parametre

        arrayKolonne := ["Ugedage", "UndtagneTransporttyper", "KørerIkkeTransporttyper"]
        for parameterNavn in raekkeArray
            for kolonneNavn in arrayKolonne
            {
                parameterNavn[kolonneNavn] := this.parameterObj.forKolonneNavn(kolonnenavn)
                parameterNavn[kolonneNavn].data["forventetIndholdArray"] := Array()
                parameterNavn[kolonneNavn].data["kolonneNummerArray"] := Array()
            }

        for rækkeIndex, raekke in excelArray
        {
            for kolonneindex, celle in raekke
                for kolonneNavn in arrayKolonne
                {
                    parameterNavn := excelArray[1][kolonneindex]
                    if parameterNavn = kolonneNavn
                    {
                        raekkeArray[rækkeIndex][kolonneNavn].data["forventetIndholdArray"].Push(celle)
                        raekkeArray[rækkeIndex][kolonneNavn].data["kolonneNummerArray"].Push(kolonneindex)
                        raekkeArray[rækkeIndex][kolonneNavn].data["forventetIndhold"] := false
                        raekkeArray[rækkeIndex][kolonneNavn].data["kolonneNavn"] := parameterNavn
                        raekkeArray[rækkeIndex][kolonneNavn].data["parameterNavn"] := parameterNavn
                        raekkeArray[rækkeIndex][kolonneNavn].data["maxArrayLængde"] := gyldigeParametre[kolonneNavn]["maxArray"]
                        raekkearray[rækkeindex][parameterNavn].data["maxParameterLængde"] := gyldigeParametre[parameterNavn]["maxLængde"]

                    }
                }
        }
        raekkeArray.RemoveAt(1)


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
            if testParameterIndholdArray.length > gyldigeParametre[testParameterNavn]["maxArray"]
            {
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

        for arrayRække, arrayIndhold in this.rækkeArray
            for mapKey, mapObj in arrayIndhold
            {
                mapObj.tjekGyldighed()
                ; _excelVerificerData.erGyldigParameterLængde(mapObj)
                ; _excelVerificerData.erGyldigArrayLængde(mapObj)
                ; if map
                ;     _excelVerificerData.erGyldigDato(mapObj)

            }

        ; for arrayRække, arrayIndhold in this.rækkeArray
        ;     for mapKey, mapObj in arrayIndhold


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

class parameterAlm {
    static forKolonneNavn(pKolonnenavn) {

        switch pKolonnenavn
        {
            case "Ugedage":
                return parameterUgedage()
            case "KørerIkkeTransporttyper":
                return parameterTransportType()
            case "UndtagneTransporttyper":
                return parameterTransportType()
            case "Starttid":
                return parameterKlokkeslæt()
            case "Sluttid":
                return parameterKlokkeslæt()
            default:
                return parameterAlm()
        }
    }
    __New() {
        this.data := parameter().parameterSæt
    }

    _forMangeTegnIParameter() {

        if StrLen(this.data["forventetIndhold"]) > this.data["maxParameterLængde"]
        {
            this._danfejl(Format("For mange tegn i parameter `"{1}`". Nuværende {2}, maks {3}.", this.data["forventetIndhold"], StrLen(this.data["forventetIndhold"]), this.data["maxParameterLængde"]))
            return
        }
    }
    _ulovligtTegnIParameter() {
        if RegExMatch(this.data["forventetIndhold"], "[\!\*\@]", &matchObj)
        {
            this._danFejl(Format("Ulovligt tegn (`"{1}`") i parameter.", matchObj[0]))
            return

        }

    }
    _danFejl(pFejlbesked) {
        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked

    }
    tjekGyldighed() {
        this._forMangeTegnIParameter()
        this._ulovligtTegnIParameter()


    }
}
class parameterUgedage {
    __New() {
        this.data := parameter().parameterSæt
    }

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
    _danFejl(pFejlbesked) {

        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked
    }
    tjekGyldighed() {
        ugedage := this.data["forventetIndholdArray"]

        for ugedag in ugedage
            if !this._erKalenderdag(ugedag)
            {
                if !this._erGyldigFastDag(ugedag)
                {
                    this._danFejl(Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", ugedag))
                    return
                }
            }
            else if !this._erGyldigDato(ugedag)
            {
                this._danFejl(Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", ugedag))
                return
            }
    }
}
class parameterTransportType {

    __new() {
        this.data := parameter().parameterSæt
    }
    _erOverMaxArray() {

        if this.data["forventetIndholdArray"].length > this.data["maxArrayLængde"]
        {
            this._danfejl(Format("For mange mange kolonner i kategori. Maks {1}, nuværende {2}", this.data["maxArrayLængde"], this.data["forventetIndholdArray"].length))
            return true
        }

    }

    _forMangeTegnIParameter() {


        for tjekParameter in this.data["forventetIndholdArray"]
            if StrLen(tjekParameter) > this.data["maxLængde"]
            {
                this._danfejl(Format("For mange tegn i parameter {1} Nuværende {2}, maks {3}", tjekParameter, StrLen(tjekParameter), this.data["maxLængde"]))
                return true
            }

    }
    _danfejl(pFejlbesked) {

        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked
    }

    tjekGyldighed() {
        if this._erOverMaxArray()
            return
        if this._forMangeTegnIParameter()
            return
    }

}
class parameterKlokkeslæt {

    __new() {
        this.data := parameter().parameterSæt
    }


    _tjekOgRensAsterisk() {

        if InStr(this.data["forventetIndhold"], "*")
        {
            this.data["forventetIndhold"] := SubStr(this.data["forventetIndhold"], 1, 5)
            this.data["sluttidspunktErNæsteDag"] := true

        }
    }

    _harKolon() {
        if !InStr(this.data["forventetIndhold"], ":")
        {
            this._danfejl(Format("Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"))
            return
        }
    }
    _korrektFormat() {


        strArr := StrSplit(this.data["forventetIndhold"], ":")

        if strArr.Length != 2
        {
            this._danfejl(Format("Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"))
            return

        }


        time := strArr[1]
        minut := strArr[2]

        if !IsTime("20241212" time minut)
            this._danfejl(Format("Fejl i format, skal være gyldigt klokkeslæt i formatet `"TT:MM`", med afsluttende asterisk hvis sluttid over midnat"))

        return
    }

    _danfejl(pFejlbesked) {

        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked
    }

    tjekGyldighed() {
        this._tjekOgRensAsterisk()
        this._harKolon()
        this._korrektFormat()
    }
}
class excelParameterInterface {

    __new() {
        this.data := parameter().parameterSæt
    }

    _danfejl(pFejlbesked) {

        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked
    }

    tjekGyldighed() {

    }

}

class parameter {
    parameterSæt {
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


; test := excelDataBehandler(excelMock.excelDataUgyldigMock, parameterAlm).behandledeRækker

; MsgBox test[1]["Ugedage"].data["fejlBesked"]
; MsgBox test[2]["Ugedage"].data["fejlBesked"]
; return
