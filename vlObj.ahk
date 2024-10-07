#Requires AutoHotkey v2.0

class vlObj extends Class
{
    vl_data := Map(
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
        "Vognløbskategori", 0,
        "Undtagne transporttyper", [],
        "Ugedage", [0, 0, 0, 0, 0, 0, 0]
    )

    vl_array := []

    IndhentData(p_excel_array)
    {
        this.vl_data := p_excel_array
        return
    }

    IndhentDataArray(p_excel_data)
    {
        for exceldata in p_excel_data
            this.vl_array.Push(exceldata)

        return
    }
}