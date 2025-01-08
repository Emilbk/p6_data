#Include includeModules.ahk

mock := excelDataUgyldigMock

class excelHentData {

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

    excelDataArray {
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

class excelStrukturerData {

    __New(excelArray) {
        this.excelArray := excelArray

    }

    danKolonneNavneOgNummer() {
        excelArray := this.excelArray
        kolonneNavne := Map()

        dataVerificering := excelVerificerData
        for rækkeIndex, kolonne in excelArray
            for kolonneIndex, kolonneNavn in kolonne
            {
                if (rækkeIndex = 1 and dataVerificering.erGyldigKolonne(kolonneNavn))
                    kolonneNavne.Set(kolonneNavn, kolonneIndex)
            }

        kolonneNavne["Ugedage"] := Array()
        kolonneNavne["UndtagneTransporttyper"] := Array()
        kolonneNavne["KørerIkkeTransporttyper"] := Array()

        for kolonne in excelArray[1]
        {
            if kolonne = "Ugedage"
                kolonneNavne["Ugedage"].Push(A_Index)
            if kolonne = "UndtagneTransporttyper"
                kolonneNavne["UndtagneTransporttyper"].Push(A_Index)
            if kolonne = "KørerIkkeTransporttyper"
                kolonneNavne["KørerIkkeTransporttyper"].Push(A_Index)
        }
        return kolonneNavne
    }


    danRækkeArray() {
        excelArray := this.excelArray
        raekkeArray := Array()
        dataVerificering := excelVerificerData
        outputArray := Array()

        for rækkeindex, raekke in excelarray
        {
            raekkearray.push(map())
            for kolonneindex, celle in raekke
            {
                kolonnenavn := excelarray[1][kolonneindex]
                if dataVerificering.erGyldigKolonne(kolonnenavn)
                {
                    raekkearray[rækkeindex].set(kolonnenavn, Object())
                    raekkearray[rækkeindex][kolonnenavn].kolonneNavn := kolonnenavn
                    raekkearray[rækkeindex][kolonnenavn].parameterNavn := kolonnenavn
                    raekkearray[rækkeindex][kolonnenavn].forventetIndhold := celle
                    raekkearray[rækkeindex][kolonnenavn].forventetIndholdArray := false

                }
            }
        }


        for kolonneNavn in raekkeArray
        {
            kolonneNavn["Ugedage"] := Object()
            kolonneNavn["Ugedage"].forventetIndholdArray := Array()
            kolonneNavn["UndtagneTransporttyper"] := Object()
            kolonneNavn["UndtagneTransporttyper"].forventetIndholdArray := Array()
            kolonneNavn["KørerIkkeTransporttyper"] := Object()
            kolonneNavn["KørerIkkeTransporttyper"].forventetIndholdArray := Array()
        }

        for rækkeIndex, raekke in excelArray
        {
            for kolonneindex, celle in raekke
            {
                kolonnenavn := excelArray[1][kolonneindex]
                if kolonnenavn = "Ugedage"
                {
                    raekkeArray[rækkeIndex]["Ugedage"].forventetIndholdArray.Push(celle)
                    raekkeArray[rækkeIndex]["Ugedage"].forventetIndhold := false
                    raekkeArray[rækkeIndex]["Ugedage"].kolonneNavn := kolonnenavn
                    raekkeArray[rækkeIndex]["Ugedage"].parameterNavn := kolonnenavn

                }
                if kolonnenavn = "UndtagneTransporttyper"
                {

                    raekkeArray[rækkeIndex]["UndtagneTransporttyper"].forventetIndholdArray.Push(celle)
                    raekkeArray[rækkeIndex]["UndtagneTransporttyper"].forventetIndhold := false
                    raekkeArray[rækkeIndex]["UndtagneTransporttyper"].kolonneNavn := kolonnenavn
                    raekkeArray[rækkeIndex]["UndtagneTransporttyper"].parameterNavn := kolonnenavn
                }
                if kolonnenavn = "KørerIkkeTransporttyper"
                {
                    raekkeArray[rækkeIndex]["KørerIkkeTransporttyper"].forventetIndholdArray.Push(celle)
                    raekkeArray[rækkeIndex]["KørerIkkeTransporttyper"].forventetIndhold := false
                    raekkeArray[rækkeIndex]["KørerIkkeTransporttyper"].kolonneNavn := kolonnenavn
                    raekkeArray[rækkeIndex]["KørerIkkeTransporttyper"].parameterNavn := kolonnenavn

                }
            }
        }
        raekkeArray.RemoveAt(1)


        return raekkeArray

    }

}

class excelVerificerData {

    static _gyldigeKolonner := gyldigeKolonner.data
    static _ugyldigeKolonner := Map()


    static _verificerKolonner(pKolonner) {
        for kolonne in pKolonner
            if !excelVerificerData._gyldigeKolonner.has(kolonne)
                excelVerificerData._ugyldigeKolonner.Set(kolonne, A_Index)
            else
                excelVerificerData._gyldigeKolonner[kolonne] := true
    }

    static ugyldigeKolonner[pKolonner] {
        get {
            excelVerificerData._verificerKolonner(pKolonner)
            return excelVerificerData._ugyldigeKolonner
        }
    }

    static gyldigeKolonner[pKolonner] {
        get {
            excelVerificerData._verificerKolonner(pKolonner)
            return excelVerificerData._gyldigeKolonner
        }
    }

    static erGyldigKolonne(kolonneNavn) {

        if excelVerificerData._gyldigeKolonner.Has(kolonneNavn)
            return true

    }
;; TODO
    static erGyldigParameter(pParameterObj) {
        gyldigeParametre := parameter.data
        testParameter := pParameterObj

        testParameterNavn := testParameter.parameterNavn
        testParameterIndholdString := testParameter.forventetIndhold
        testParameterIndholdArray := testParameter.forventetIndholdArray


        if testParameterIndholdString
            if StrLen(testParameterIndholdString) != gyldigeParametre[testParameterNavn]["maxLængde"]
                MsgBox testParameterIndholdString "passer ike"

        if testParameterIndholdArray
            for parameterIndhold in testParameterIndholdArray
                if StrLen(parameterIndhold) != gyldigeParametre[testParameterNavn]["maxLængde"]
                    MsgBox parameterIndhold "passer ike array"

    }

}

arr := excelStrukturerData(mock).danRækkeArray()
koll := excelStrukturerData(mock).danKolonneNavneOgNummer()
test := excelVerificerData.gyldigeKolonner[koll]

for arrayRække, arrayIndhold in arr
    for mapKey, mapObj in arrayIndhold
        excelVerificerData.erGyldigParameter(mapObj)


return