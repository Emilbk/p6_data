class config {
    
    breakLoop := 0

    setBreakLoop(){
        
        this.breakLoop := 1
    
    }

    removeBreakLoop(){

        this.breakLoop := 0
    }

    getBreakLoopStatus(){

        return this.breakLoop
    }
}