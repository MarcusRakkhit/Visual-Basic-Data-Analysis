Function CreateAveragePriceChart(routeGroups As Collection)
    Dim ws As Worksheet
    Dim routeName As String
    Dim flights As Collection
    Dim i As Long, r As Long
    Dim price As Variant
    Dim avg As Double, total As Double, count As Long
    Dim chartObj As ChartObject
    Dim chartRange As Range
    Dim tbl As ListObject

    '-----------------------------------------
    ' Check if the sheet "Route_Averages" exists
    '-----------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Route_Averages")
    On Error GoTo 0
    
    '-----------------------------------------
    ' Create / clear sheet
    '-----------------------------------------
    If ws Is Nothing Then
        ' Sheet does not exist, so create it
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Route_Averages"
        MsgBox "Route Averages added"
    Else
        ' Sheet exists, delete and recreate
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        
        ' Add a fresh sheet
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Route_Averages"
        MsgBox "Route Averages overridden"
    End If
    
    
    '-----------------------------------------
    ' Write headers
    '-----------------------------------------
    ws.Range("A1").Value = "Route"
    ws.Range("B1").Value = "Average Price (AUD)"

    '-----------------------------------------
    ' Compute averages
    '-----------------------------------------
    r = 2
    For i = 1 To routeGroups.count
        routeName = routeGroups(i)(1)
        Set flights = routeGroups(i)(2)

        total = 0
        count = 0
    
        Dim f As Long
        
        'Adds all route prices (and gets count)
        For f = 1 To flights.count
            price = flights(f)(4)
            If IsNumeric(price) And price > 0 Then
                total = total + CDbl(price)
                count = count + 1
            End If
        Next f
        ' Divides total with count to get the mean/average
        If count > 0 Then avg = total / count Else avg = 0
        
        ' Prints the route and average
        ws.Cells(r, 1).Value = routeName
        ws.Cells(r, 2).Value = avg
        r = r + 1
    Next i

    '-----------------------------------------
    ' Format as table with orange style
    '-----------------------------------------
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B" & r - 1), , xlYes)
    tbl.TableStyle = "TableStyleMedium10"   ' Orange style
    tbl.Range.Columns.AutoFit
    ws.Columns("B").NumberFormat = "$#,##0.00"

    '-----------------------------------------
    ' Create stacked bar chart
    '-----------------------------------------
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=20, Width:=600, Height:=400)
    chartObj.Chart.ChartType = xlBarStacked
    Set chartRange = ws.Range("A1:B" & r - 1)

    With chartObj.Chart
        .SetSourceData chartRange
        .HasTitle = True
        .ChartTitle.Text = "Average Price per Route"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Characters.Text = "Routes"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Characters.Text = "Average Price (AUD)"
    End With
End Function

Function CreateTotalPriceChart(routeGroups As Collection)
    Dim ws As Worksheet
    Dim r As Long, f As Long, i As Long
    Dim routeName As String
    Dim flights As Collection
    Dim price As Double
    Dim maxFlights As Long
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    Dim seriesRange As Range
    Dim categoriesRange As Range
    Dim s As Long
    
    '-----------------------------------------
    ' Check if the sheet "Route_Totals" exists
    '-----------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Route_Totals")
    On Error GoTo 0
    
    '-----------------------------------------
    ' Create / clear sheet
    '-----------------------------------------
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Route_Totals"
        MsgBox "Route Totals added"
    Else
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Route_Totals"
        MsgBox "Route Totals overridden"
    End If
    
    '-----------------------------------------
    ' Determine max flights
    '-----------------------------------------
    maxFlights = 0
    For i = 1 To routeGroups.count
        If routeGroups(i)(2).count > maxFlights Then maxFlights = routeGroups(i)(2).count
    Next i
    
    '-----------------------------------------
    ' Write headers
    '-----------------------------------------
    ws.Cells(1, 1).Value = "Route"
    For f = 1 To maxFlights
        ws.Cells(1, f + 1).Value = "Flight " & f
    Next f
    
    '-----------------------------------------
    ' Fill flight prices
    '-----------------------------------------
    For r = 1 To routeGroups.count
        routeName = routeGroups(r)(1)
        Set flights = routeGroups(r)(2)
        
        ws.Cells(r + 1, 1).Value = routeName
        For f = 1 To flights.count
            price = flights(f)(4)
            If IsNumeric(price) And price > 0 Then
                ws.Cells(r + 1, f + 1).Value = CDbl(price)
            Else
                ws.Cells(r + 1, f + 1).Value = 0
            End If
        Next f
    Next r
    
    '-----------------------------------------
    ' Format as table with orange style
    '-----------------------------------------
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(routeGroups.count + 1, maxFlights + 1)), , xlYes)
    tbl.TableStyle = "TableStyleMedium10"
    tbl.Range.Columns.AutoFit
    ws.Range(ws.Columns(2), ws.Columns(maxFlights + 1)).NumberFormat = "$#,##0.00"
    
    '-----------------------------------------
    ' Create stacked bar chart
    '-----------------------------------------
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=20, Width:=700, Height:=400)
    chartObj.Chart.ChartType = xlBarStacked
    
    ' Series = numeric data (exclude Route column)
    Set seriesRange = ws.Range(ws.Cells(2, 2), ws.Cells(routeGroups.count + 1, maxFlights + 1))
    ' Categories = route names
    Set categoriesRange = ws.Range(ws.Cells(2, 1), ws.Cells(routeGroups.count + 1, 1))
    
    With chartObj.Chart
        .SetSourceData seriesRange
        .HasTitle = True
        .ChartTitle.Text = "Total Price per Route (All Flights)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Routes"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Price (AUD)"
        .Legend.Position = xlLegendPositionRight
        
        '-----------------------------------------
        ' Rename series to Flight 1, Flight 2, ...
        '-----------------------------------------
        For s = 1 To .SeriesCollection.count
            .SeriesCollection(s).Name = "Flight " & s
        Next s
    End With
End Function
Function FindAverageAndTotal(flights As Collection) As Collection
    Dim routes As New Collection        ' unique route names
    Dim seen As New Collection          ' helper to track duplicates
    Dim result As New Collection        ' final result

    Dim route As String
    Dim flight As Collection
    Dim group As Collection
    Dim pair As Collection
    Dim i As Long, r As Long

    '-----------------------------------------
    ' STEP 1 — Build unique route list
    '-----------------------------------------
    On Error Resume Next

    For i = 2 To flights.count
        Set flight = flights(i)
        route = CStr(flight(2))  ' route is element 2 of flight

        seen.Add True, route
        If Err.Number = 0 Then
            routes.Add route
        Else
            Err.Clear
        End If
    Next i

    On Error GoTo 0

    '-----------------------------------------
    ' STEP 2 — For each route, collect flights
    '-----------------------------------------

    For r = 1 To routes.count
        route = routes(r)

        ' new group of flights for this route
        Set group = New Collection

        ' scan all flights again to find matches
        For i = 2 To flights.count
            Set flight = flights(i)
            ' Checks if the route matches the route within the 'routes' list
            If CStr(flight(2)) = route Then
                group.Add flight
            End If
        Next i

        ' create {route, group}
        Set pair = New Collection
        pair.Add route          ' item(1) = route
        pair.Add group          ' item(2) = collection of flights

        result.Add pair
    Next r

    '-----------------------------------------
    ' STEP 3 — Call chart creation
    '-----------------------------------------
    Call CreateAveragePriceChart(result)
    Call CreateTotalPriceChart(result)

    '-----------------------------------------
    ' Return grouped collection
    '-----------------------------------------
    Set FindAverageAndTotal = result
End Function


' Creates a table from excel data
Function MakeTable(flights) As Collection
    
    ' Opens the Flights_Cleaned excel worksheet
    Dim flightSheet As Worksheet
    Set flightSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    
    ' Defines table we will store excel data
    Set flights = New Collection
    
    'Defines row and column size from excel
    Dim lastRow As Long, LastCol As Long
    lastRow = flightSheet.Cells(flightSheet.Rows.count, 1).End(xlUp).Row
    LastCol = flightSheet.Cells(1, flightSheet.Columns.count).End(xlToLeft).Column
    
    Dim r As Long, c As Long
    
    ' Stores every value that will be taken into "Flights" table
    Dim flight As Collection
    
    ' Loop through each row in the sheet
    For r = 1 To lastRow
        Set flight = New Collection
        ' Adds cell value into the flight row
        For c = 1 To LastCol
            flight.Add flightSheet.Cells(r, c).Value
        Next c
        ' Adds the row into the flight table
        flights.Add flight
    Next r
    
    ' Returning table
    Set MakeTable = flights
End Function

' Adds Calculations sheet
Function AddSheet(sheetName As String)
    ' Stores new worksheet
    Dim newSheet As Worksheet
    
    ' Check if sheet already exists
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If newSheet Is Nothing Then
        ' Add new sheet
        ThisWorkbook.Sheets.Add.Name = sheetName
        MsgBox "Average added"
    Else
        ' Disable confirmation dialogs (don't really need alerts when the sheets are deleted before overriding)
        Application.DisplayAlerts = False
        newSheet.Delete
        ' Delete existing sheet
        Application.DisplayAlerts = True
        
        ' Adds a fresh sheet
        MsgBox "Average overriden"
        ThisWorkbook.Sheets.Add.Name = sheetName
    End If
    

End Function

Sub FindMean()
    ' Stores new worksheet
    Dim checkSheet As Worksheet
    
    ' this button can only be played if the person has created the "Flights_Cleaned" table
    On Error Resume Next
    Set checkSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    On Error GoTo 0
    
    ' creates a new table
    Dim flights As Collection
    Set flights = MakeTable(flights)
    
    ' Finds the average
    Dim average As Collection
    Set average = FindAverageAndTotal(flights)
        
End Sub