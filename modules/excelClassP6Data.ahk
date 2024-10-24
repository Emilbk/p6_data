/************************************************************************
 * @description Excel-class til brug ved P6-data-makro
 * @author 
 * @date 2024/10/18
 * @version 0.0.1
 * @extends excelclass.ahk
 ***********************************************************************/


#Include excelClass.ahk

/**
 * @parameter gyldigeKolonner Map,
 */
class excelObjP6Data extends excelClass {

    /** @type {Map} */
    gyldigeKolonner := Map(
        "Budnummer", 0,
        "Vognløbsnummer", 0,
        "Kørselsaftale", 0,
        "Styresystem", 0,
        "Startzone", 0,
        "Slutzone", 0,
        "Hjemzone", 0,
        "MobilnrChf", 0,
        "Vognløbskategori", 0,
        "Planskema", 0,
        "Økonomiskema", 0,
        "Statistikgruppe", 0,
        "Vognløbsnotering", 0,
        "Starttid", 0,
        "Sluttid", 0,
        "Sluttid", 0,
        "Undtagne transporttyper", 0,
        "Ugedage", 0
    )

    ugyldigeKolonner := Map(

    )
    
    ; Kolonnenavne opgivet i excel-ark, men ikke defineret i script
    p6DataTjekForGyldigeKolonner(){
        for kolonneNavn in this.aktivWorksheetKolonneNavnOgNummer
            if this.gyldigeKolonner.Has(kolonneNavn)
                this.gyldigeKolonner[kolonneNavn] := 1
            else
                this.ugyldigeKolonner[kolonneNavn] := 0
    
        for kolonnenavn,indhold in this.ugyldigeKolonner
            if indhold = 0
                MsgBox kolonneNavn " er ikke gyldig"
            return
    }


}

