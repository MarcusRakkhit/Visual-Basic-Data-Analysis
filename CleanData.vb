
' Converts any price not in AUD
Function ConvertToAUD(RowData) As Boolean
        ConvertToAUD = True
        
        Dim NZDRate As Double
        Dim SGDRate As Double
        Dim IDRRate As Double
        Dim THBRate As Double
        
        'This is manual data for currency difference to AUD and could change
        NZDRate = 0.88
        SGDRate = 1.18
        IDRRate = 0.000097
        THBRate = 0.043
        Dim convertedPrice As Double
        convertedPrice = 0
        
        'Converts money based on its currency
        ' the formula is '[price]*[Currency conversion]'
        If (RowData(5) = "AUD") Then
            Exit Function
        ElseIf (RowData(5) = "NZD") Then
            convertedPrice = RowData(4) * NZDRate
        ElseIf (RowData(5) = "SGD") Then
            convertedPrice = RowData(4) * SGDRate
        ElseIf (RowData(5) = "IDR") Then
            convertedPrice = RowData(4) * IDRRate
        ElseIf (RowData(5) = "THB") Then
            convertedPrice = RowData(4) * THBRate
        Else
            ConvertToAUD = False
            Exit Function
        End If
        
        RowData.Remove 4
        RowData.Add convertedPrice, , 4   ' Insert converted price at the same position
        
        'Ensures all currency is converted to AUD
        RowData.Remove 5
        RowData.Add "AUD", , 5
End Function

Function CreateExcelTable(ParamArray sheetNames())

    Dim ws As Worksheet
    Dim dataRange As Range
    Dim tableObj As ListObject
    Dim lastRow As Long, LastCol As Long, firstCol As Long
    Dim i As Long, c As Long
    
    ' We have multiple sheets to create
    For i = LBound(sheetNames) To UBound(sheetNames)

        Set ws = ThisWorkbook.Sheets(sheetNames(i))

        ' -------- Remove existing tables --------
        Do While ws.ListObjects.count > 0
            ws.ListObjects(1).Unlist
        Loop

        ' -------- Detect real data --------
        Dim lastCell As Range
        Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, _
                     LookAt:=xlPart, SearchOrder:=xlByRows, _
                     SearchDirection:=xlPrevious)
        If lastCell Is Nothing Then GoTo NextSheet

        lastRow = lastCell.Row
        LastCol = ws.Cells.Find(What:="*", LookIn:=xlValues, _
                     LookAt:=xlPart, SearchOrder:=xlByColumns, _
                     SearchDirection:=xlPrevious).Column
        firstCol = 1

        ' -------- Rename any blank headers to "Reasons" if this is Excepted_Data --------
        If sheetNames(i) = "Excepted_Data" Then
            For c = 1 To LastCol
                If Trim(ws.Cells(1, c).Value) = "" Then
                    ws.Cells(1, c).Value = "Reasons"
                End If
            Next c

            ' Optional: If you want to always ensure there is at least one "Reasons" column at the end
            If WorksheetFunction.CountIf(ws.Rows(1), "Reasons") = 0 Then
                LastCol = LastCol + 1
                ws.Cells(1, LastCol).Value = "Reasons"
            End If
        End If

        ' -------- Set data range --------
        Set dataRange = ws.Range(ws.Cells(1, firstCol), ws.Cells(lastRow, LastCol))

        ' -------- Create the table --------
        Set tableObj = ws.ListObjects.Add( _
                SourceType:=xlSrcRange, _
                Source:=dataRange, _
                XlListObjectHasHeaders:=xlYes)

        tableObj.Name = "Table_" & ws.Name
        tableObj.TableStyle = "TableStyleMedium9"
        
        If sheetNames(i) = "Excepted_Data" Then
            Dim reasonColIndex As Long
            Dim reasonColumn As Range
            tableObj.TableStyle = "TableStyleMedium10"
            ' Loop through all table columns
            For c = 1 To tableObj.ListColumns.count
                If LCase(Left(tableObj.ListColumns(c).Name, 7)) = "reasons" Then
                    reasonColIndex = c
                    ' Apply red fill immediately if you want multiple columns
                    tableObj.ListColumns(c).DataBodyRange.Interior.Color = RGB(255, 200, 200)
                End If
            Next c
            
        End If

NextSheet:
    Next i

End Function


' Ensures there are no errors in the table
Function cleanTable(flightSheet) As Collection
    Dim i As Long, j As Long
    
    ' Stores data that's removed:
    ' Each item in ExceptedData = { RowDataCollection , ReasonsCollection }
    Dim ExceptedData As Collection
    Set ExceptedData = New Collection
    
    ' Temporary structures
    Dim RowData As Collection
    Dim RemovedRow As Collection
    Dim ReasonList As Collection
    
    ' Gets the Flights_Cleaned sheet we created in the "AddSheet" function
    Dim getSheet As Worksheet
    Set getSheet = ThisWorkbook.Sheets("Flights_Cleaned")
    
    ' Gets the Excepted_Data sheet we created in the "AddSheet" function
    Dim exceptedSheet As Worksheet
    Set exceptedSheet = ThisWorkbook.Sheets("Excepted_Data")
    
    ' Checks to see if there are duplicate IDs
    Dim IDDict As Object
    Set IDDict = CreateObject("Scripting.Dictionary")
    
    For i = flightSheet.count To 2 Step -1
        
        Set RowData = flightSheet(i)
        
        
        
        ' Holds all reasons the row is invalid
        Set ReasonList = New Collection
        
        ' ---- Check rules and build reasons ----
        
        ' ---- Duplicate ID check (column 1 of RowData) ----
        Dim rowID As Variant
        rowID = RowData(1)
        ' Checks if ID is duplicated
        If IDDict.Exists(rowID) Then
            ReasonList.Add "Duplicate ID"
        Else
            IDDict.Add rowID, 1
        End If
        
        
        ' Checks if the status is ACTIVE (column 6)
        If RowData(6) <> "ACTIVE" Then
            ReasonList.Add "Status not ACTIVE"
        End If
        
        ' Checks that Price numeric & greater than $0 (column 4)
        If Not IsNumeric(RowData(4)) Then
            ReasonList.Add "Price not numeric"
        ElseIf flightSheet(i)(4) < 1 Then
            ReasonList.Add "Price less than $1"
        End If
        
        'Check if the price is numeric (otherwise it will cause errors)
        If IsNumeric(RowData(4)) Then
            ' Ensures that price is converted to AUD
            Dim priceCheck As Boolean
            priceCheck = ConvertToAUD(RowData)
        End If
        
        ' Price conversion failure (if there's an unknown currency)
        If priceCheck = False Then
            ReasonList.Add "Currency conversion failed"
        End If
        
        ' Blank cell check
        For j = 1 To RowData.count
            If IsEmpty(RowData(j)) Then
                ReasonList.Add "Blank cell in attribute " & j
            End If
        Next j
        
        ' ---- If failed any criteria, store and remove ----
        If ReasonList.count > 0 Then
            
            ' Store removed row
            Set RemovedRow = New Collection
            For j = 1 To RowData.count
                RemovedRow.Add RowData(j)
            Next j
            
            ' Create wrapper {RowData, Reasons}
            Dim Entry As Collection
            Set Entry = New Collection
            Entry.Add RemovedRow      ' Entry(1)
            Entry.Add ReasonList      ' Entry(2)
            
            ExceptedData.Add Entry    ' Add to master list
            
            ' Remove row from main table
            flightSheet.Remove i
        End If
        
    Next i
    
    ' ----- IN EXCEL SHEET PRINT CLEANED TABLE WITH HEADERS -----

    ' Header titles (replace with your actual attributes)
    Dim headers As Variant
    headers = Array("ID", "Origin", "Destination", "Price", "Currency", "Status")
    
    ' Write headers to cleaned sheet
    For j = LBound(headers) To UBound(headers)
        getSheet.Cells(1, j + 1).Value = headers(j)
    Next j
    
    ' Write cleaned flight data under headers
    For i = 1 To flightSheet.count
        Set RowData = flightSheet(i)
    
        For j = 1 To RowData.count
            getSheet.Cells(i, j).Value = RowData(j)
        Next j
    Next i
    
    
    
    ' ----- PRINT EXCEPTED DATA WITH HEADERS -----
    
    ' Write attribute headers to excepted sheet
    For j = LBound(headers) To UBound(headers)
        exceptedSheet.Cells(1, j + 1).Value = headers(j)
    Next j
    
    ' Add a header for "Reasons"
    exceptedSheet.Cells(1, UBound(headers) + 2).Value = "Reasons"
    
    
    
    ' ----- WRITE EXCEPTED ROWS -----
    
    Dim RowValues As Collection
    Dim Reasons As Collection
    Dim reasonStartCol As Long
    
    For i = 1 To ExceptedData.count
    
        Set Entry = ExceptedData(i)
        Set RowValues = Entry(1)
        Set Reasons = Entry(2)
    
        ' ----- Write the attribute values in row i+1 -----
        For j = 1 To RowValues.count
            exceptedSheet.Cells(i + 1, j).Value = RowValues(j)
        Next j
    
        ' Starting column for reasons
        reasonStartCol = RowValues.count + 1
    
        ' ----- Write ALL reasons horizontally to the right -----
        For j = 1 To Reasons.count
            exceptedSheet.Cells(i + 1, reasonStartCol + j - 1).Value = Reasons(j)
        Next j
    
    Next i
    
End Function

' Creates a table from original excel dataset
Function MakeTable(flights) As Collection
    
    ' Opens the original excel worksheet
    Dim flightSheet As Worksheet
    Set flightSheet = ThisWorkbook.Sheets("Flights")
    
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
        ' Adds cell value into the row
        For c = 1 To LastCol
            flight.Add flightSheet.Cells(r, c).Value
        Next c
        ' Adds the row into the table
        flights.Add flight
    Next r
    
    ' Returning table
    Set MakeTable = flights
End Function

' Adds Flights_Cleaned sheet
Function AddSheet(sheetName As String)
    ' Stores new worksheet
    Dim newSheet As Worksheet
    
    ' Check if sheet already exists
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets(sheetName)
    Set ExeceptedSheet = ThisWorkbook.Sheets("Excepted_Data")
    On Error GoTo 0
    
    If newSheet Is Nothing Then
        ' Add new sheet
        ThisWorkbook.Sheets.Add.Name = sheetName
        ThisWorkbook.Sheets.Add.Name = "Excepted_Data"
        MsgBox "Flights_Cleaned added"
    Else
        ' Disable confirmation dialogs (don't really need alerts when the sheets are deleted before overriding)
        Application.DisplayAlerts = False
        newSheet.Delete
        ExeceptedSheet.Delete
        ' Delete existing sheet
        Application.DisplayAlerts = True
        
        ' Adds a fresh sheet
        MsgBox "Flights_Cleaned overriden"
        ThisWorkbook.Sheets.Add.Name = sheetName
        ThisWorkbook.Sheets.Add.Name = "Excepted_Data"
    End If
    
    
    
End Function

' Once "Clean Data" Button has been clicked, this function will open
Sub CleanData()
    
    AddSheet ("Flights_Cleaned")
    
    'Creates a table from excel sheet
    Dim flights As Collection
    Set flights = MakeTable(flights)
    
    'Cleans the table
    Dim clean As Collection
    Set clean = cleanTable(flights)
    
    ' Creates a tables - "Flights_Cleaned" and "Excepted_Data"
    Dim result As Variant
    result = CreateExcelTable("Flights_Cleaned", "Excepted_Data")
    
    
End Sub
