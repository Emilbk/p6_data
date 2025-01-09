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
        gyldigeParametre := parameter.data

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
                    raekkearray[rækkeindex][parameterNavn].data["maxLængde"] := gyldigeParametre[parameterNavn]["maxLængde"]

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
    static erGyldigParameterLængde(pParameterObj) {
        gyldigeParametre := parameter.data
        testParameter := pParameterObj

        testParameterNavn := testParameter.data["parameterNavn"]
        testParameterIndholdString := testParameter.data["forventetIndhold"]
        testParameterIndholdArray := testParameter.data["forventetIndholdArray"]


        if testParameterIndholdString and gyldigeParametre[testParameterNavn]["maxLængde"]
            if StrLen(testParameterIndholdString) > gyldigeParametre[testParameterNavn]["maxLængde"]
            {
                testParameter.data["fejl"] := 1
                testParameter.data["fejlBesked"] := "for mange tegn i parameter"
                testParameter.data["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxLængde"]
            }

        if testParameterIndholdArray and gyldigeParametre[testParameterNavn]["maxLængde"]
            for parameterIndhold in testParameterIndholdArray
                if StrLen(parameterIndhold) > gyldigeParametre[testParameterNavn]["maxLængde"]
                {
                    testParameter.data["fejl"] := 1
                    testParameter.data["fejlBesked"] := "for mange tegn i parameter"
                    testParameter.data["fejlParameterArray"] := parameterIndhold
                    testParameter.data["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxLængde"]
                }
    }
    static erGyldigArrayLængde(pParameterObj) {

        gyldigeParametre := parameter.data
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

    static erGyldigDato(pParameterObj) {
        datoArray := pParameterObj["Ugedage"]["forventetIndholdArray"]
        gyldigeUgedage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]

        for dato in datoArray
        {
            if !InStr(dato, "/")
                for gyldigUgedag in gyldigeUgedage
                    if gyldigUgedag = dato
                        break
                    else
                        MsgBox "nje"

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

        if pKolonnenavn = "Ugedage"
            return parameterUgedage()
        else
            return parameterAlm()
    }
    __New() {

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

    }

    _danFejl(pFejlbesked) {
        this.data["fejl"] := 1
        this.data["fejlBesked"] := pFejlbesked

    }
    tjekGyldighed() {

        if StrLen(this.data["forventetIndhold"]) > this.data["maxLængde"]
        {
            this._danFejl(Format("For mange tegn i parameter."))
            return
        }
            
        if RegExMatch(this.data["forventetIndhold"], "[\!\*\@]", &matchObj)
        {
            this._danFejl(Format("Ulovligt tegn (`"{1}`") i parameter.", matchObj[0]))
            return

        }
    }
}
class parameterUgedage {
    __New() {

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

    }

    static _erKalenderdag(pDag) {
        ugedag := pDag
        if RegExMatch(ugedag, "\w*\d\w*")
            return true
    }
    static _erFastDag(pDag) {
        ugedag := pDag

        gyldigeUgedage := ["ma", "ti", "on", "to", "fr", "lø", "sø"]
        erUgedag := false

        for gyldigUgedag in gyldigeUgedage
            if ugedag = gyldigUgedag
                return erUgedag := true
    }
    static _erGyldigDato(pDato) {
        dato := pDato

        if !InStr(dato, "/")
            return false

        datoArry := StrSplit(dato, "/")
        dag := datoArry[1]
        måned := datoArry[2]
        år := datoArry[3]

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
        {
            if parameterUgedage._erKalenderdag(ugedag)
            {
                if !parameterUgedage._erGyldigDato(ugedag)
                {
                    this._danFejl(Format("Fejl i kalenderdato: {1}. Skal være gyldig dato i formatet mm/dd/åååå.", ugedag))
                    return
                }
            }
            else if !parameterUgedage._erFastDag(ugedag)
            {
                this._danFejl(Format("fejl i fast dag: {1}. Skal være i formatet XX, f. eks MA", ugedag))
                return
            }
        }


    }
}
class excelParameterInterface {

    __new() {

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

    }

    _danfejl() {

    }

    tjekGyldighed() {

    }

}


; test := excelDataBehandler(excelMock.excelDataUgyldigMock, parameterAlm).behandledeRækker

; MsgBox test[1]["Ugedage"].data["fejlBesked"]
; MsgBox test[2]["Ugedage"].data["fejlBesked"]
; return
