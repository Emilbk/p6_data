#Include includeModules.ahk

; TODO
; anden parameterGyldighed objec
; Hvordan dobbelte parametre?
; parameterFactory?


mock := excelMock.excelDataUgyldigMock
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

    __New(excelArray) {
        this.excelArray := excelArray

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

        for rækkeindex, raekke in excelarray
        {
            raekkearray.push(map())
            for kolonneindex, celle in raekke
            {
                kolonnenavn := excelarray[1][kolonneindex]
                if dataVerificering.erGyldigKolonne(kolonnenavn)
                {
                    raekkearray[rækkeindex].set(kolonnenavn, _excelParameter.ny)
                    raekkearray[rækkeindex][kolonnenavn]["kolonneNavn"] := kolonnenavn
                    raekkearray[rækkeindex][kolonnenavn]["parameterNavn"] := kolonnenavn
                    raekkearray[rækkeindex][kolonnenavn]["forventetIndhold"] := celle
                    raekkearray[rækkeindex][kolonnenavn]["kolonneNummer"] := kolonneindex

                }
            }
        }
        ; TODO
        ; kopier dobbelt parametre

        arrayKolonne := ["Ugedage", "UndtagneTransporttyper", "KørerIkkeTransporttyper"]
        for kolonneNavn in raekkeArray
            for arr in arrayKolonne
            {
                kolonneNavn[arr] := _excelParameter.ny
                kolonneNavn[arr]["forventetIndholdArray"] := Array()
                kolonneNavn[arr]["kolonneNummerArray"] := Array()
            }

        for rækkeIndex, raekke in excelArray
        {
            for kolonneindex, celle in raekke
                for arr in arrayKolonne
                {
                    kolonnenavn := excelArray[1][kolonneindex]
                    if kolonnenavn = arr
                    {
                        raekkeArray[rækkeIndex][arr]["forventetIndholdArray"].Push(celle)
                        raekkeArray[rækkeIndex][arr]["kolonneNummerArray"].Push(kolonneindex)
                        raekkeArray[rækkeIndex][arr]["forventetIndhold"] := false
                        raekkeArray[rækkeIndex][arr]["kolonneNavn"] := kolonnenavn
                        raekkeArray[rækkeIndex][arr]["parameterNavn"] := kolonnenavn

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

        testParameterNavn := testParameter["parameterNavn"]
        testParameterIndholdString := testParameter["forventetIndhold"]
        testParameterIndholdArray := testParameter["forventetIndholdArray"]


        if testParameterIndholdString and gyldigeParametre[testParameterNavn]["maxLængde"]
            if StrLen(testParameterIndholdString) > gyldigeParametre[testParameterNavn]["maxLængde"]
            {
                testParameter["fejl"] := 1
                testParameter["fejlBesked"] := "for mange tegn i parameter"
                testParameter["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxLængde"]
            }

        if testParameterIndholdArray and gyldigeParametre[testParameterNavn]["maxLængde"]
            for parameterIndhold in testParameterIndholdArray
                if StrLen(parameterIndhold) > gyldigeParametre[testParameterNavn]["maxLængde"]
                {
                    testParameter["fejl"] := 1
                    testParameter["fejlBesked"] := "for mange tegn i parameter"
                    testParameter["fejlParameterArray"] := parameterIndhold
                    testParameter["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxLængde"]
                }
    }
    static erGyldigArrayLængde(pParameterObj) {

        gyldigeParametre := parameter.data
        testParameter := pParameterObj

        testParameterNavn := testParameter["parameterNavn"]
        testParameterIndholdArray := testParameter["forventetIndholdArray"]

        if testParameterIndholdArray
            if testParameterIndholdArray.length > gyldigeParametre[testParameterNavn]["maxArray"]
            {
                testParameter["fejl"] := 1
                testParameter["fejlBesked"] := "for mange kolonner i kategori"
                testParameter["fejlParameterArray"] := testParameter["kolonneNavn"]
                testParameter["maxParameterLængde"] := gyldigeParametre[testParameterNavn]["maxArray"]
            }
    }

}

class _excelParameter {

    static ny {

        get {
            data := Map()
            data.Default := 0

            data["kolonneNavn"] := false
            data["kolonneNummer"] := false
            data["kolonneNummerArray"] := false
            data["parameterNavn"] := false
            data["forventetIndhold"] := false
            data["forventetIndholdArray"] := false
            data["faktiskIndhold"] := false
            data["faktiskIndholdArray"] := false
            data["fejl"] := false
            data["fejlBesked"] := false
            data["fejlParameterArray"] := false
            data["maxParameterLængde"] := false
            data["maxArrayLængde"] := false
            data["vognløbsdato"] := false

            return data
        }


    }
}

class excelDataBehandler {

    __New(pExcelFil, pArkNavnEllerNummer, excelApp := "") {
        this.excelFil := pExcelFil
        this.arkNavnEllerNummer := pArkNavnEllerNummer
        if excelApp
            this.app := excelApp

        this.excel := _excelHentData(pExcelFil, pArkNavnEllerNummer, excelApp)
        this.excelArray := excelMock.excelDataUgyldigMock
        ; this.excelArray := this.excel.getDataArray
        this.rækkeArray := _excelStrukturerData(this.excelArray).danRækkeArray()
        this.kolonner := _excelStrukturerData(this.excelArray).danKolonneNavneOgNummer()

        for arrayRække, arrayIndhold in this.rækkeArray
            for mapKey, mapObj in arrayIndhold
            {
                _excelVerificerData.erGyldigParameterLængde(mapObj)
                _excelVerificerData.erGyldigArrayLængde(mapObj)

            }

        ; for arrayRække, arrayIndhold in this.rækkeArray
        ;     for mapKey, mapObj in arrayIndhold

        this.excel._quit()

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

; excelArray := excelDataBehandler(excelPath, excelArk)
; fArray := excelArray.behandledeRækker
; return