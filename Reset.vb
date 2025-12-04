Sub Reset()

    Dim ws As Worksheet
    Dim sheetName As String
    
    ' Turn off alerts to delete sheets without confirmation
    Application.DisplayAlerts = False
    
    ' Removes all sheets, except for the one labelled 'Flights'
    ' 'Flights' is the original sheet that comes with this Excel document
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        If sheetName <> "Flights" Then
            ws.Delete
        End If
    Next ws
    
    ' Turn alerts back on
    Application.DisplayAlerts = True
    
    MsgBox "All Sheets have been reset", vbInformation


End Sub
