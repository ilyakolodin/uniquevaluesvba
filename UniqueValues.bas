Option Explicit

Const sRange = "A1:F3"

Sub main()
    Dim uniques As Collection
    Dim source As Range
    Dim ws As Worksheet
    Dim outputSheet As Worksheet
    Dim it As Variant
    Dim i As Integer

    ' Define the source range
    Set source = ActiveSheet.Range(sRange)
    
    ' Get unique values from the source range
    Set uniques = GetUniqueValues(source.Value)

    ' Check if the "UniqueValues" sheet already exists
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("UniqueValues")
    On Error GoTo 0

    ' If the sheet doesn't exist, create it
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = "UniqueValues"
    Else
        ' Clear existing data in the sheet
        outputSheet.Cells.Clear
    End If

    ' Write unique values to the new sheet
    i = 1
    For Each it In uniques
        outputSheet.Cells(i, 1).Value = it
        i = i + 1
    Next it

    ' Notify the user
    MsgBox "Unique values have been written to the 'UniqueValues' sheet."

End Sub

Public Function GetUniqueValues(ByVal values As Variant) As Collection
    Dim result As Collection
    Dim cellValue As Variant
    Dim cellValueTrimmed As String

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        cellValueTrimmed = Trim(cellValue)
        If cellValueTrimmed = "" Then GoTo NextValue
        result.Add cellValueTrimmed, cellValueTrimmed
NextValue:
    Next cellValue

    On Error GoTo 0
End Function
