Attribute VB_Name = "AddComponentFunction"
Public PriceC As Double

Public Function Add_Component(CompName As String, PoDate As String)
'Updates
ThisWorkbook.Connections("Query - Custom Prices").Refresh

On Error GoTo Errhandler

Dim wb As Workbook
Dim CheckCustom As Range
'Opens BOM workbook
Set wb = Workbooks.Open("Your Sharepoint URL")
Sheets("Component Pricing").Select
Set ws = Worksheets("Component Pricing")
Set table = ws.ListObjects("ComponentPricing")
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Removing front of partnames
If Left(CompName, 6) = "OPINV:" Then
CompNamef = Right(CompName, Len(CompName) - 6)
Else
CompNamef = CompName
End If

errNum = 2

Set CheckCustom = ws.Range("A1:A" & lRow).Find(CompNamef)

If Not CheckCustom Is Nothing Then
  With Workbooks("BOMsForHoses.xlsx").Worksheets("Component Pricing")
    .Range(Cells(CheckCustom.Row, 1).Address).Value = CompNamef
    .Range(Cells(CheckCustom.Row, 2).Address).Value = PriceC
    .Range(Cells(CheckCustom.Row, 3).Address).Value = PoDate
    End With
Else
     With Workbooks("BOMsForHoses.xlsx").Worksheets("Component Pricing")
    .Range(Cells(CDbl(lRow + 1), 1).Address).Value = CompNamef
    .Range(Cells(CDbl(lRow + 1), 2).Address).Value = PriceC
    .Range(Cells(CDbl(lRow + 1), 3).Address).Value = PoDate

  End With
End If

wb.Close True 'save changes
ThisWorkbook.Connections("Query - Custom Prices").Refresh
    
GoTo EndSub

Errhandler:

If errNum = 1 Then
MsgBox ("comp " & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 3 Then
MsgBox ("comp " & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If


EndSub:

ThisWorkbook.Connections("Query - Custom Prices").Refresh

End Function
