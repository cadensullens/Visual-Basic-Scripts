Attribute VB_Name = "FRunBOMCost"


Function CollectBOMNames()
Dim ActTrue As String

'Sheets("BOM Master").Select

Dim BOMParts() As String
Dim i  As Long
Dim lRow As Long
'Find the last non-blank cell in column A(1)
    lRow = Worksheets("BOM Master").Cells(Rows.Count, 1).End(xlUp).Row
'One shorter than lRow due to header
ReDim BOMParts(1 To lRow - 1)
'starts at 2 due to headers in table
For i = 2 To lRow
    BOMParts(i - 1) = Worksheets("BOM Master").Range(Cells(CDbl(i), CDbl(1)).Address).Value
Next i


With Worksheets("BOM PriceSheet")

For i = 1 To UBound(BOMParts)
 .Range(Cells(CDbl(i + 1), CDbl(1)).Address).Value = BOMParts(i)
Next i

End With
Call RunBOMCosts(BOMParts)
End Function

Function RunBOMCosts(BOMParts() As String)
Dim table1 As ListObject
Dim table2 As ListObject
Dim CustomCheck As Boolean
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim i As Integer
Dim j As Integer


Set ws1 = Worksheets("TiteFlex Pricing")
Set table1 = ws1.ListObjects("TiteFlex_Pricing")
Set ws2 = Worksheets("Custom Prices")
Set table2 = ws2.ListObjects("Custom_Prices")

For i = 1 To UBound(BOMParts)
CustomCheck = False
PriceMissing = False
Call Check_BOM(BOMParts(i))

For j = LBound(PartNames) To UBound(PartNames)
On Error GoTo Errhandler
Start3G:
    
     'Titeflex Pricing finds price on that sheet
      'Vlookup as sheet does not have duplicate P/N
     errNum = 3
     Price = Application.WorksheetFunction.VLookup(PartNames(j), table1.Range.Columns("A:F"), 4, False)
     ReDim Preserve PriceList(1 To j)
     PriceList(j) = Price
     GoTo ContinueG
     
CustomPrice:
     'errMsg = "Component " & PartNames(j) & " is NOT on the Custom component pricing Sheet, Confirm Spelling of Part and Date. If correct, then Part is not on the Custom component pricing Sheet."
     errNum = 31
     Price = Application.WorksheetFunction.VLookup(PartNames(j), table2.Range.Columns("A:C"), 2, False)
     
     ReDim Preserve PriceList(1 To j)
     PriceList(j) = Price
     CustomCheck = True
     GoTo ContinueG
    errNum = 0
     
Errhandler:
If errNum = 3 Then
    Resume CustomPrice
End If
    
 
If errNum = 31 Then
    Price = 0
    ReDim Preserve PriceList(1 To j)
    PriceList(j) = Price
    PriceMissing = True
    Resume ContinueG
End If

If errNum = 0 Then
    MsgBox "error"
    GoTo EndSub:
End If
    
    
ContinueG:
     ReDim Preserve Grand(1 To j)
     Grand(j) = compQTY(j) * Round(PriceList(j), 2)

Next j

With Worksheets("BOM PriceSheet")
    If CustomCheck = False Then
        'Final Sum for extra options
        Grandsum = Round(Application.WorksheetFunction.Sum(Grand), 2) + (10 * WireHole) + BarbRoy
        .Range(Cells(CDbl(i + 1), CDbl(3)).Address).Value = Grandsum
    Else
        
        'Creates Value for Grand Total
        Grandsum = Round(Application.WorksheetFunction.Sum(Grand), 2) + (10 * WireHole) + BarbRoy
        .Range(Cells(CDbl(i + 1), CDbl(3)).Address).Value = Grandsum
        If PriceMissing = False Then
            PriceText = "Review for Custom Pricing"
        Else
            PriceText = "Some Pricing missing"
        End If
        .Range(Cells(CDbl(i + 1), CDbl(2)).Address).Value = PriceText
    End If
End With
Next i

EndSub:
End Function
