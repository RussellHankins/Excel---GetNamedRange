Attribute VB_Name = "Module1"
Public Function GetNamedRange(oWorkbook As Workbook, sNamedRange As String) As Range
'Returns the Range object corresponding to the named range sNamedRange in workbook oWorkbook.
Dim iRangeCount As Long
Dim iRangeLoop As Long
Dim oRange As Range
    Set oRange = Nothing
    iRangeCount = oWorkbook.Names.Count
    If iRangeCount > 0 Then
        For iRangeLoop = 1 To iRangeCount
            If oWorkbook.Names(iRangeLoop).Name = sNamedRange Then
                Set oRange = oWorkbook.Names(iRangeLoop).RefersToRange
                Exit For
            End If
        Next
    End If
    Set GetNamedRange = oRange
End Function
