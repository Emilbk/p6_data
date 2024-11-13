

overordnetFunk(){
    try{
        underordnetFunk1()
        underordnetFunk2()
        underordnetFunk2()

    }
    catch TypeError as tError{
        
    }
    catch ValueError as vError{

    }
   catch IndexError as Funk1{

   }
}

underordnetFunk1(){
   try{
       MsgBox "Funk1"
       throw IndexError("Funk1") 

   }
}
underordnetFunk2(){
    underordnetFunk21()

    MsgBox "funk2"
}

underordnetFunk21(){

    MsgBox "Funk21"
    try{
        throw ValueError("Funk21")

    }catch ValueError as funk21error{

    }
}

overordnetFunk()

return