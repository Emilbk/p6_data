#Requires AutoHotkey v2.0

class fejlRegister {

    static _register := []

    static registrerFejl(fejlObj) {

        fejlRegister._register.Push(fejlObj)
    }
    static reset() {
        fejlRegister._register := []
    }
}

class parameterFejl extends Error {

    fejlData := {}

}
