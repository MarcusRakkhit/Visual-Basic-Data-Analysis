' Converts all elements in "Flights_Cleaned" into a proper table
Function CreateExcelTable()

        Dim flightSheet As Worksheet
        Dim dataRange As Range
        Dim DateTable As ListObject
        

        ' Set the worksheet where the table will be created
        Set flightSheet = ThisWorkbook.Sheets("Flights_Cleaned")
        
        ' Gets the row and column size
        Dim LastRow As Long, LastCol As Long
        LastRow = flightSheet.Cells(flightSheet.Rows.Count, 1).End(xlUp).Row
        LastCol = flightSheet.Cells(1, flightSheet.Columns.Count).End(xlToLeft).Column
        
        
        ' Define the range of data for the table
        Set dataRange = flightSheet.Range( _
            flightSheet.Cells(1, 1), _
            flightSheet.Cells(LastRow, LastCol))

        ' Create new table
        Set DateTable = flightSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=dataRange, _
            XlListObjectHasHeaders:=xlYes)
        
        ' Labelling/Styling the table
        DateTable.Name = "CleanFlightSet"
        DateTable.TableStyle = "TableStyleMedium9"

End Function

Function RowHasDuplicates(flightSheet) As Collection
    ' Dict will store rows
    Dim dict As Object
    
    Dim i As Long
    
    ' Stores values without duplicate IDs
    Dim newSheet As New Collection
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To flightSheet.Count
        ' Checks if the duplicated ID exists in dictionary object
        If Not dict.Exists(flightSheet(i)(1)) Then
            ' If not, then the row will be added and kept to the new excel sheet
            dict.Add flightSheet(i)(1), 1
            newSheet.Add flightSheet(i)
        End If
    Next i
    
    Set RowHasDuplicates = newSheet
    
End Function

' Ensures there are no errors in the table
Function cleanTable(flightSheet) As Collection
    Dim i As Long, j As Long
    
    ' Stores a flight row temporily to see if there are errors in a particular row
    Dim rowData As Collection
    
    ' Gets the sheet we created in the "AddSheet" function
    Dim getSheet As Worksheet
    Set getSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    
    ' Checks the function to see if there are duplicate IDs
    Set flightSheet = RowHasDuplicates(flightSheet)
    
    For i = flightSheet.Count To 2 Step -1
        ' sets temporaily row with its current index
        Set rowData = flightSheet(i)
    
        For j = 1 To rowData.Count
            ' Checks if the row contains blank cells, has an ACTIVE status, has a pricing less than $1 and/or contains non-numeric values in pricing
            If IsEmpty(rowData(j)) Or flightSheet(i)(5) <> "ACTIVE" Or flightSheet(i)(4) < 1 Or Not IsNumeric(flightSheet(i)(4)) Then
                ' Removes row from the table
                flightSheet.Remove i
                Exit For
            End If
        Next j
    
    Next i
    
    ' Prints all cells into the new sheet
    For i = 1 To flightSheet.Count
        Set rowData = flightSheet(i)
    
        For j = 1 To rowData.Count
            getSheet.Cells(i, j).value = rowData(j)
        Next j
    
    Next i
    
End Function

' Creates a table from original excel dataset
Function MakeTable(Flights) As Collection
    
    ' Opens the original excel worksheet
    Dim flightSheet As Worksheet
    Set flightSheet = ThisWorkbook.Sheets("Flights")
    
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
        ' Adds cell value into the row
        For c = 1 To LastCol
            Flight.Add flightSheet.Cells(r, c).value
        Next c
        ' Adds the row into the table
        Flights.Add Flight
    Next r
    
    ' Returning table
    Set MakeTable = Flights
End Function

' Adds Flights_Cleaned sheet
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
        MsgBox "Flights_Cleaned added"
    Else
        ' Disable confirmation dialogs
        Application.DisplayAlerts = False
        newSheet.Delete
        ' Delete existing sheet
        Application.DisplayAlerts = True
        
        ' Adds a fresh sheet
        MsgBox "Flights_Cleaned overriden"
        ThisWorkbook.Sheets.Add.Name = sheetName
    End If
    
    
    
End Function

' Once "Clean Data" Button has been clicked, this function will open
Sub CleanData()
    
    AddSheet ("Flights_Cleaned")
    
    Dim Flights As Collection
    Set Flights = MakeTable(Flights)
    
    Dim clean As Collection
    Set clean = cleanTable(Flights)
    
    Dim result As Variant
    result = CreateExcelTable()
    
    
End Sub