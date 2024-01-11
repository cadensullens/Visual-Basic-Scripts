Attribute VB_Name = "OpenBOMFunction"
Function Open_BOM()
On Error GoTo Errhandler

Dim wb As Workbook

'Opens BOM workbook
Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/BOMsForHoses.xlsx")
Sheets("BOM Master").Select
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
   'MsgBox "Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol
'Broken need to fix need to set wb BOM workbook
'Test tomorrow
errNum = 0
Startover:
With Workbooks("BOMsForHoses.xlsx").Worksheets("BOM Master")
For i = LBound(PartNames) To UBound(PartNames)
    .Range(Cells(CDbl(lRow + 1), CDbl(2 + i * 2)).Address).Value = PartNames(i)
    .Range(Cells(CDbl(lRow + 1), CDbl(3 + i * 2)).Address).Value = compQTY(i)
Next i

.Range(Cells(CDbl(lRow + 1), 1).Address).Value = hose
.Range(Cells(CDbl(lRow + 1), 2).Address).Value = WireHole
.Range(Cells(CDbl(lRow + 1), 3).Address).Value = BarbRoy
End With
wb.Close True 'save changes
GoTo EndSub
Errhandler:

If errNum = 0 Then
wb.Close False 'does not save changes
GoTo EndSub
End If

EndSub:
End Function



