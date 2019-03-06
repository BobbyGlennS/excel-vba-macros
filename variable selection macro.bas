Sub selectColumnsByNames()

    Dim dataSheet As String
    Dim newDataSheet As String
    Dim totalRows As Long
    Dim i As Long
    Dim WS As Worksheet
    Dim currentColumn As Long
    Dim columnHeader As String
    Dim varToKeep As String

'USER SETTINGS ##########################################################
    'Define the name of the sheet where the data is stored
    dataSheet = "data_DG"
    newDataSheet = "selected_variables_sheet"

'DATA OPERATIONS #######################################################
    'Get variable names of columns to keep
    totalRows = Worksheets("filters").Rows(Rows.Count).End(xlUp).Row
    ReDim columnsToKeep(1 To totalRows - 1)

    For i = 1 To totalRows - 1
        columnsToKeep(i) = Worksheets("filters").Cells(i + 1, 1).Value
    Next i

    'Create new datasheet to copy variables to
    If Not WorksheetExists(newDataSheet) Then
        Set WS = Sheets.Add(After:=Sheets(Worksheets.Count))
        WS.Name = newDataSheet
    End If

    'copy columns that we need to keep to new sheet.
    For i = 1 To totalRows - 1
         currentColumn = [0]
         varToKeep = columnsToKeep(i)
         On Error Resume Next
         currentColumn = Application.Match(varToKeep, Sheets(dataSheet).Rows(1), 0)
         If Not currentColumn = 0 Then
            Worksheets(dataSheet).Columns(currentColumn).Copy Worksheets(newDataSheet).Columns(i)
         Else
            Worksheets(newDataSheet).Cells(1, i) = varToKeep + "DOES_NOT_EXIST"
         End If
    Next i

    For i = 1 To totalRows - 1
        Worksheets("newDataSheet").Cells(i, 1).Value = columnsToKeep(i)
    Next i

End Sub

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
   Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
