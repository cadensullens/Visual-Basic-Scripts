Attribute VB_Name = "BuySellCopyTableFunction"

Function BuySell_CopyTable(locR As Variant, locC As Variant, newName As String)

         'copies top of template above table and paste it
         Worksheets("Template").Range("A4:D5").copy
         Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address)
         TargetRange.PasteSpecial xlPasteAll
         
         'copies first cell of template to perserve formatting and uses for pasting
         Worksheets("Template").Range("A4").copy
         
         Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(3, 0)
         TargetRange.PasteSpecial xlPasteAll
         
         Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2, 0)
         TargetRange.PasteSpecial xlPasteAll
        
         Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2, 2)
         TargetRange.PasteSpecial xlPasteAll
         
         'places borders around the new cells
         Worksheets(newName).Range(Cells(CDbl(locR + 3), CDbl(locC)).Address).BorderAround _
         LineStyle:=xlContinuous, Weight:=xlThin
         Worksheets(newName).Range(Cells(CDbl(locR + 3), CDbl(locC + 1)).Address).BorderAround _
         LineStyle:=xlContinuous, Weight:=xlThin
         Worksheets(newName).Range(Cells(CDbl(locR + 2), CDbl(locC + 1)).Address).BorderAround _
         LineStyle:=xlContinuous, Weight:=xlThin
         Worksheets(newName).Range(Cells(CDbl(locR + 2), CDbl(locC + 3)).Address).BorderAround _
         LineStyle:=xlContinuous, Weight:=xlThin
         
         'copies margin table
         Worksheets("Template").Range("L6:N8").copy
         Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC + 5)).Address)
         TargetRange.PasteSpecial xlPasteAll
         
         'sets formula for base price
         Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC + 6)).Address).Formula2 = ("=" & Cells(CDbl(locR + 1), CDbl(locC + 1)).Address)
         
         'Places formulas for margin in correct cells
'         For i = 1 To 1
'         Worksheets(newName).Range(Cells(CDbl(locR + 2 + i), CDbl(locC + 7)).Address).Formula2 = ("=" & Cells(CDbl(locR), CDbl(locC + 6)).Address & "/" & "(1-" & Cells(CDbl(locR + 2 + i), CDbl(locC + 6)).Address & ")")
'         Next i
         
End Function
