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

    ' Gaat kijken wat de eerste kolom is van Laatst geproduceerd
    eersteKol = 0
    For k = 1 To laatsteKolom
        If InStr(1, ws.Cells(1, k).Value, "LastProduced", vbTextCompare) > 0 Then
            eersteKol = k
            Exit For
        End If
    Next k
    
    If eersteKol > 0 Then
        ws.Columns(eersteKol).Insert Shift:=xlToRight
        ws.Cells(1, eersteKol).Value = "Laatst geproduceerd"
        laatsteKolom = laatsteKolom + 1
        
        
        For r = 2 To LaatsteRij
            jaar = 2000
            maand = 1
            dag = 1
            uur = 0
            minuut = 0
            ' Voegt alle laatst geproduceerde datums bij elkaar
            For k = eersteKol + 1 To laatsteKolom
                If InStr(1, ws.Cells(1, k).Value, "LastProduced_YEAR", vbTextCompare) > 0 Then
                    jaar = CInt(ws.Cells(r, k).Value)
                ElseIf InStr(1, ws.Cells(1, k).Value, "LastProduced_MONTH", vbTextCompare) > 0 Then
                    maand = CInt(ws.Cells(r, k).Value)
                ElseIf InStr(1, ws.Cells(1, k).Value, "LastProduced_DAY", vbTextCompare) > 0 Then
                    dag = CInt(ws.Cells(r, k).Value)
                ElseIf InStr(1, ws.Cells(1, k).Value, "LastProduced_HOUR", vbTextCompare) > 0 Then
                    uur = CInt(ws.Cells(r, k).Value)
                ElseIf InStr(1, ws.Cells(1, k).Value, "LastProduced_MINUTE", vbTextCompare) > 0 Then
                    minuut = CInt(ws.Cells(r, k).Value)
                End If
            Next k
            ' Geeft de formaat toe aan het de cellen van laatst geproduceerd
            ws.Cells(r, eersteKol).Value = DateSerial(jaar, maand, dag) + TimeSerial(uur, minuut, 0)
            ws.Cells(r, eersteKol).NumberFormat = "dd-mm-yyyy hh:mm"
        Next r
        ' verwijder alle andere laatst geproduceerde kolomen
        For k = laatsteKolom To eersteKol + 1 Step -1
            If InStr(1, ws.Cells(1, k).Value, "LastProduced", vbTextCompare) > 0 Then
                ws.Columns(k).Delete
            End If
        Next k
    End If

    ' voegt aan elke kolomen onder die kop een eigen unit toe
    kopNamen = Array("Doek lengte", "Blok hoogte", "Doek lengte Z laag", "Blok breedte", "Totaal geproduceerd", "Aantal afgebroken", "Totaal doek verbruikt", "Gemiddelde cyclustijd doek snijden", "Gemiddelde cyclustijd doek lijmen", "Gemiddelde cyclustijd wikkelen", "Gemiddelde cyclustijd Z laag wikkelen", "Totale Gemiddelde cyclustijd")
    For Each kopNaam In kopNamen
        Set gevondenKol = ws.Rows(1).Find(What:=kopNaam, LookAt:=xlWhole, MatchCase:=False)
        
        If Not gevondenKol Is Nothing Then
            LaatsteRij = ws.Cells(ws.Rows.Count, gevondenKol.Column).End(xlUp).Row
            
            For r = 2 To LaatsteRij
                Set cel = ws.Cells(r, gevondenKol.Column)
                
                If IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Doek lengte" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Blok hoogte" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Doek lengte Z laag" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Blok breedte" Then
                    cel.Value = CDbl(cel.Value)
                    cel.NumberFormat = "0"" mm"""
                ElseIf IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Totaal geproduceerd" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Aantal afgebroken" Then
                    cel.Value = CDbl(cel.Value)
                    cel.NumberFormat = "0"" Stuks"""
                ElseIf IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Totaal doek verbruikt" Then
                    cel.Value = CDbl(cel.Value)
                    cel.NumberFormat = "0.00"" m"""
                ElseIf IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Gemiddelde cyclustijd doek snijden" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Gemiddelde cyclustijd doek lijmen" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Gemiddelde cyclustijd wikkelen" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Gemiddelde cyclustijd Z laag wikkelen" Or IsNumeric(cel.Value) And cel.Value <> "" And kopNaam = "Totale Gemiddelde cyclustijd" Then
                    cel.Value = CDbl(cel.Value)
                    cel.NumberFormat = "0.0"" sec"""
                End If
            Next r
        End If
    Next kopNaam

    Set gevondenKol = ws.Rows(1).Find(What:="Blok type", LookAt:=xlWhole, MatchCase:=False)
    ' geeft de cellen de juiste blok type toe
    If Not gevondenKol Is Nothing Then
        LaatsteRij = ws.Cells(ws.Rows.Count, gevondenKol.Column).End(xlUp).Row
        
        For r = 2 To LaatsteRij
            Set cel = ws.Cells(r, gevondenKol.Column)
            If cel.Value = 0 Then
                cel.Value = "Rechthoek"
            Else
                cel.Value = "Schuin"
            End If
        Next r
    End If

    Set ptKol = ws.Rows(1).Find(What:="productie tijd (h:m)", LookAt:=xlPart, MatchCase:=False)
    'verandert de productie tijd naar uren en minuten
    If Not ptKol Is Nothing Then
        LaatsteRij = ws.Cells(ws.Rows.Count, ptKol.Column).End(xlUp).Row
        
        For r = 2 To LaatsteRij
            tijdTekst = Trim(ws.Cells(r, ptKol.Column).Value)
            
            If InStr(tijdTekst, ":") > 0 Then
                If InStr(tijdTekst, ".") > 0 Then
                    tijdTekst = Split(tijdTekst, ".")(0)
                End If
                
                delen = Split(tijdTekst, ":")
                
                If UBound(delen) >= 2 Then
                    dUur = CDbl(delen(0))
                    dMinuut = CDbl(delen(1))
                    dSeconde = CDbl(delen(2))
                    ws.Cells(r, ptKol.Column).Value = (dUur / 24) + (dMinuut / 1440) + (dSeconde / 86400)
                    ws.Cells(r, ptKol.Column).NumberFormatLocal = "[uu]:mm"
                End If
            End If
        Next r
    End If