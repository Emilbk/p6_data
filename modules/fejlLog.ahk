class P6MsgboxError extends fejlLogObj {
    

    getP6MsgboxFejl(){

        this.fejlLog := fejlLogObj()

        return this.fejllog
    }

}


class p6ForkertDataError extends fejlLogObj {
    
}

class p6DataTjekMismatchError extends fejlLogObj {
    
}

class P6Indtastningsfejl extends fejlLogObj {
    
    construct(pvognløb){

        this.vognløbsnummer := pvognløb.tilIndlæsning.Vognløbsnummer
        this.vognløbsdato := pvognløb.tilIndlæsning.Vognløbsdato

    }
}

class P6ClipboardFejl extends fejlLogObj {
    

    testufnk(){

        for prop, propval in this.OwnProps()
            MsgBox prop " - " propval
    }
}

class fejlLogObj extends Error {
    

    __New(message := "", what := "", extra := "", dataObj := "") {
        super.__New(message, what, extra)
    }
    setVognløbsnummerOgDato(pVognløb)
    {
        this.vognløbsnummer := pVognløb["Vognløbsnummer"]
        this.vognløbsdato := pVognløb["Vognløbsdato"]
    }
    
    importP6MsgboxFejl(P6Msgbox)
    {
        
        this.p6fejlbesked := P6Msgbox.message
        this.p6msgbox := P6Msgbox.extra
        this.p6fejlstack := P6Msgbox.stack
        this.p6fejlwhat := P6Msgbox.what

    }

    importP6DataFejl(p6Datafejl)
    {

    }
}