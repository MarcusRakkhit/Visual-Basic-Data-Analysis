' Writes the total price of flights in the new worksheet
Function Calculateitems(Flights) As Collection
    
    Dim i As Long, j As Long
    
    ' Stores a flight row temporily to see if there are errors
    Dim rowData As Collection
    
    ' Gets the sheet we created in the "AddSheet" function
    Dim getSheet As Worksheet
    Set getSheet = ThisWorkbook.Sheets("Calculations")
    
    Dim total As Long
    
    ' The total starts at 0 by default
    total = 0
    
    ' Adds all the data entry prices
    For i = Flights.Count To 2 Step -1
        total = total + Flights(i)(4)
    Next i
    
    ' Prints all cells into the new sheet
    getSheet.Cells(1, 1).value = "TOTAL PRICE:"
    getSheet.Cells(2, 1).value = total
End Function

' Creates a table from excel data
Function MakeTable(Flights) As Collection
    
    ' Opens the Flights_Cleaned excel worksheet
    Dim flightSheet As Worksheet
    Set flightSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    
    ' Defines table we will store excel data
    Set Flights = New Collection
    
    'Defines row and column size from excel
    Dim LastRow As Long, LastCol As Long
    LastRow = flightSheet.Cells(flightSheet.Rows.Count, 1).End(xlUp).Row
    LastCol = flightSheet.Cells(1, flightSheet.Columns.Count).End(xlToLeft).Column
    
    Dim r As Long, c As Long
    
    ' Stores every value that will be taken into "Flights" table
    Dim Flight As Collection
    
    ' Loop through each row in the sheet
    For r = 1 To LastRow
        Set Flight = New Collection
        ' Adds cell value into the flight row
        For c = 1 To LastCol
            Flight.Add flightSheet.Cells(r, c).value
        Next c
        ' Adds the row into the flight table
        Flights.Add Flight
    Next r
    
    ' Returning table
    Set MakeTable = Flights
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
        MsgBox "Calculations added"
    Else
        ' Disable confirmation dialogs
        Application.DisplayAlerts = False
        newSheet.Delete
        ' Delete existing sheet
        Application.DisplayAlerts = True
        
        ' Adds a fresh sheet
        MsgBox "Calculations overriden"
        ThisWorkbook.Sheets.Add.Name = sheetName
    End If
    

End Function

' Once "Calculate Data" Button has been clicked, this function will open
Sub CalculateData()
    
    ' Stores new worksheet
    Dim checkSheet As Worksheet
    
    ' this button can only be played if the person has created the "Flights_Cleaned" table
    On Error Resume Next
    Set checkSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    On Error GoTo 0
    
    If checkSheet Is Nothing Then
        ' error prompt
        MsgBox "You need to clean your data first"
    Else
        AddSheet ("Calculations")
    
        Dim Flights As Collection
        Set Flights = MakeTable(Flights)
        
        Dim calculate As Collection
        Set calculate = Calculateitems(Flights)
    End If
    
    
End Sub
