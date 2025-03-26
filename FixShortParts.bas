Attribute VB_Name = "FixShortParts"
Function NameFix()

'Sheets("Short Parts").Select
'Finds the last non-blank cell in a single row or column
With Worksheets("Short Parts")
Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = .Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    
    
VarOne = .Range(Cells(CDbl(lRow), CDbl(12)).Address).Value
VarTwo = .Range(Cells(CDbl(lRow), CDbl(9)).Address).Value

If VarOne = VarTwo Then GoTo EndFunc

For i = 2 To lRow

OGCompName = .Range(Cells(CDbl(i), CDbl(9)).Address).Value

If Left(OGCompName, 1) = "Y" Then
    j = 1
    Character = Right(OGCompName, 1)
    If Character = "H" Then
    Do While Character Like "[a-zA-Z]"
        j = j + 1
        Character = Left(Right(OGCompName, j), 1)
    Loop
    NewCompName = Left(OGCompName, Len(OGCompName) - (j - 1))
    .Range(Cells(CDbl(i), CDbl(12)).Address).Value = NewCompName
    Else
    .Range(Cells(CDbl(i), CDbl(12)).Address).Value = .Range(Cells(CDbl(i), CDbl(9)).Address).Value
    End If
    
    
Else

.Range(Cells(CDbl(i), CDbl(12)).Address).Value = .Range(Cells(CDbl(i), CDbl(9)).Address).Value
End If
Next i

End With

'Sheets(OGName).Select
EndFunc:
End Function
