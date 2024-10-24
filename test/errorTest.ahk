fejlfunk() {
    throw Error("Dette er en fejl")
}


try {

    fejlfunk()
} catch Error as fejl {
    MsgBox fejl.What
}