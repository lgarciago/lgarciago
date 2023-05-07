Sub ConvertUnits()
    Dim ws As Worksheet
    Dim cell As Range
    Dim inputStr As String
    Dim multiplier As Double
    Dim value As Double
    
    Set ws = ThisWorkbook.ActiveSheet
    
    For Each cell In ws.Range("B2:B" & ws.Cells(ws.Rows.Count, 2).End(xlUp).Row)
        If Not IsEmpty(cell) Then
            inputStr = cell.Value
            If Right(inputStr, 1) = "m" Then
                multiplier = 0.001
                value = Left(inputStr, Len(inputStr) - 1) * multiplier
            ElseIf Right(inputStr, 1) = "p" Then
                multiplier = 0.000000000001
                value = Left(inputStr, Len(inputStr) - 1) * multiplier
            ElseIf Right(inputStr, 1) = "u" Then
                multiplier = 0.000001
                value = Left(inputStr, Len(inputStr) - 1) * multiplier
            Else
                value = inputStr
            End If
            cell.Value = value
        End If
    Next cell
End Sub

