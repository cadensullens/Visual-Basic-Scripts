Attribute VB_Name = "BOMCopyTableFunction"

Function BOM_CopyTable(locR As Variant, locC As Variant, newName As String)

'Initial Template placement
Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address)

'For loop to place template multiple times
'For i = 1 To NumberHose
'Goes to template sheet and copies the formatted table
Worksheets("Template").Range("A4:I7").copy

TargetRange.PasteSpecial xlPasteAll

'Setting range for Price Breaks
RW = locR
CL = locC + 8
For i = 1 To UBound(PartNames)
    For j = 2 To breakCount

    Worksheets(newName).Range(Cells(CDbl(RW + 2 + i), CDbl(CL + j - 1)).Address).Borders.LineStyle = xlContinuous
    Next j
    Next i
    
    
    Worksheets("Template").Range("L5:N10").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(CL + breakCount + 1)).Address)
    TargetRange.PasteSpecial xlPasteAll
    
    Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(CL + breakCount + 2)).Address).Formula2 = ("=" & Cells(CDbl(locR + 1), CDbl(locC + 1)).Address)
    
    
    If LCase(SpecClean) = "yes" Then

        Worksheets("Template").Range("A4").copy
         
        Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(locC + 6)).Address)
        TargetRange.PasteSpecial xlPasteAll
        'box around
        Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(locC + 6)).Address).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlThin
        Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(locC + 7)).Address).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlThin
        Worksheets(newName).Range(Cells(CDbl(locR + 1), CDbl(locC + 6)).Address).Value = "Clean Price"
    End If
    
End Function
