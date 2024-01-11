Attribute VB_Name = "BOMCopyTableFunction"

Function BOM_CopyTable(locR As Variant, locC As Variant, newName As String)

'Initial Template placement
Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address)

'For loop to place template multiple times
'For i = 1 To NumberHose
'Goes to template sheet and copies the formatted table
Worksheets("Template").Range("A4:H16").copy

TargetRange.PasteSpecial xlPasteAll

'Setting range for Price Breaks
RW = locR
CL = locC + 8
    For j = 0 To breakCount - 1
    Worksheets("Template").Range("J4:J16").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(RW), CDbl(CL + j)).Address)
    TargetRange.PasteSpecial xlPasteAll
    Worksheets(newName).Range(Cells(CDbl(RW + 2), CDbl(CL + j)).Address).Value = "Price Break " & j + 1
    
    Next j
    
    Worksheets("Template").Range("L6:N15").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR + 2), CDbl(CL + breakCount + 1)).Address)
    TargetRange.PasteSpecial xlPasteAll
    
    Worksheets(newName).Range(Cells(CDbl(locR + 2), CDbl(CL + breakCount + 2)).Address).Formula2 = ("=" & Cells(CDbl(locR + 1), CDbl(locC + 1)).Address)
    
         If breakCount <= 7 Then
         looper = 7
         Else
         looper = breakCount - 1
         End If
    For j = 0 To breakCount - 1
        Worksheets(newName).Range(Cells(CDbl(locR + 4 + j), CDbl(CL + breakCount + 1)).Address).Value = partQty(j + 1)
    Next j
    
    For j = 0 To looper
    
        If j < looper Then
        Worksheets(newName).Range(Cells(CDbl(locR + 5 + j), CDbl(CL + breakCount + 2)).Address).Formula2 = ("=" & Cells(CDbl(locR + 4 + j), CDbl(CL + breakCount + 2)).Address & "-.01")
        End If
        Worksheets(newName).Range(Cells(CDbl(locR + 4 + j), CDbl(CL + breakCount + 3)).Address).Formula2 = ("=" & Cells(CDbl(locR + 2), CDbl(CL + breakCount + 2)).Address & "/" & "(1-" & Cells(CDbl(locR + 4 + j), CDbl(CL + breakCount + 2)).Address & ")")
    Next j

'changes Target Range to place below previous table
 'Set TargetRange = Cells(Rows.Count, 1).End(xlUp).Offset(2, 0)
'Next i
End Function
