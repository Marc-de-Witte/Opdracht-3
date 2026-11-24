Attribute VB_Name = "Module1"
Sub ImporteerCSV()
    Dim bestandsPad As Variant
    Dim ws As Worksheet
    Dim laatsteKolom As Long
    Dim LaatsteRij As Long
    Dim eersteKol As Long
    Dim k As Long
    Dim r As Long
    Dim lengteKol As Range
    Dim cel As Range
    Dim kopNamen As Variant
    Dim kopNaam As Variant
    Dim gevondenKol As Range
    Dim ptKol As Range
    Dim delen() As String
    Dim tijdTekst As String
    Dim dDag As Double
    Dim dUur As Double
    Dim dMinuut As Double
    Dim dSeconde As Double
    Dim naamKol As Range
    Dim jaar As Integer
    Dim maand As Integer
    Dim dag As Integer
    Dim uur As Integer
    Dim minuut As Integer
    Dim vergelijkWs As Worksheet
    Dim zoekMatch As Range
    Dim celWaarde As String
    Dim blad As Worksheet
    Dim rekenKoppen As Variant
    Dim rKop As Variant
    Dim naamWaarde As String
    Dim oudeRij As Long
    Dim lpKol As Range
    Dim lpKolOud As Range
    Dim lpColIdx As Long
    Dim lpColIdxOud As Long
    Dim rOud As Long
    Dim gevondenRij As Long
    Dim laatsteRijOud As Long
    Dim bestaatAl As Boolean
    Range("A1:U500").NumberFormat = "0"
    bestandsPad = Application.GetOpenFilename("CSV-bestanden (*.csv), *.csv")
    If bestandsPad = False Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Cells.NumberFormat = "General"
    ws.Name = "Import " & Date & " " & Format(Now, "hhmmss")
    ' Voegt elke query van de CSV in een eigen Excel kolom.
    With ws.QueryTables.Add(Connection:="TEXT;" & bestandsPad, Destination:=ws.Range("A1"))
        .TextFilePlatform = 65001
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileDecimalSeparator = "."
        .TextFileThousandsSeparator = " "
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    ' Verander alle engelse namen naar Nederlandse namen
    ws.UsedRange.NumberFormat = "0"
    ws.UsedRange.Replace What:="OrderData_Name", Replacement:="Naam", LookAt:=xlPart
    ws.UsedRange.Replace What:="OrderData_FabricLengt_Zlayer", Replacement:="Doek lengte Z laag", LookAt:=xlPart
    ws.UsedRange.Replace What:="OrderData_FabricLengt", Replacement:="Doek lengte", LookAt:=xlPart
    ws.UsedRange.Replace What:="OrderData_BlockHight", Replacement:="Blok hoogte", LookAt:=xlPart
    ws.UsedRange.Replace What:="OrderData_BlockWith", Replacement:="Blok breedte", LookAt:=xlPart
    ws.UsedRange.Replace What:="OrderData_BlockType", Replacement:="Blok type", LookAt:=xlPart
    ws.UsedRange.Replace What:="Statistics_TotalProduced", Replacement:="Totaal geproduceerd", LookAt:=xlPart
    ws.UsedRange.Replace What:="Statistics_FailedProducts", Replacement:="Aantal afgebroken", LookAt:=xlPart
    ws.UsedRange.Replace What:="Statistics_TotalFabricUsed", Replacement:="Totaal doek verbruikt", LookAt:=xlPart
    ws.UsedRange.Replace What:="Statistics_TotalTimeActive", Replacement:="Productie tijd (h:m)", LookAt:=xlPart
    ws.UsedRange.Replace What:="AverageCycleTime_FabricTransport", Replacement:="Gemiddelde cyclustijd doek snijden", LookAt:=xlPart
    ws.UsedRange.Replace What:="AverageCycleTime_Glue", Replacement:="Gemiddelde cyclustijd doek lijmen", LookAt:=xlPart
    ws.UsedRange.Replace What:="AverageCycleTime_WrappingZlayer", Replacement:="Gemiddelde cyclustijd Z laag wikkelen", LookAt:=xlPart
    ws.UsedRange.Replace What:="AverageCycleTime_Wrapping", Replacement:="Gemiddelde cyclustijd wikkelen", LookAt:=xlPart
    ws.UsedRange.Replace What:="AverageCycleTime_Total", Replacement:="Totale Gemiddelde cyclustijd", LookAt:=xlPart
    laatsteKolom = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    LaatsteRij = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row