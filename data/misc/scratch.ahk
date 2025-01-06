#Include ../test/lib/json.ahk


Budnummer := { parameterNavn: "Budnummer", kolonneNavn: "Budnummer", forventetIndhold: "", eksisterendeIndhold: "", fejl: 0, iBrug: 0, kolonneNummer: 0, maxLængde: "" }

jstr := jsongo.Stringify(Budnummer)
jobj := jsongo.Parse(jstr)


return