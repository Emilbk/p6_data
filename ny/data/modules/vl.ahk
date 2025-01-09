#Include includeModules.ahk
test := excelDataBehandler(excelMock.excelDataGyldigMock).behandledeRækker

; TODO dobbelte parametre skal opreetsse
class vognløb {

    static udrulVognløb(pDataRække) {

        outPut := Array()

        for vognløbsdata in pDataRække
        {
            rækkeIndex := A_Index
            outPut.Push(Array())
            outPut[rækkeIndex].masterVl := vognløbsdata
            outPut[rækkeIndex].vlNummer := vognløbsdata["Vognløbsnummer"]["forventetIndhold"]
            for vlDato in vognløbsdata["Ugedage"]["forventetIndholdArray"]
            {
                nyParameterMidl := DeepCopy(vognløbsdata)
                nyParameter := nyParameterMidl()

                nyParameter["Vognløbsdato"] := excelParameter().ny
                nyParameter["Vognløbsdato"]["forventetIndhold"] := vlDato
                nyParameter["VognløbsdatoNæste"]["forventetIndhold"] := vlDato

                outPut[rækkeIndex].Push(vognløb.danVognløbForVognløbsdato(nyParameter))


            }

        }


        return outPut
    }

    static danVognløbForVognløbsdato(pVLData) {


        return pVLData
    }
}

testvl := vognløb.udrulVognløb(test)

testvl[1][1]["Vognløbsdato"]["forventetIndhold"] := "testsetst"
return