Attribute VB_Name = "AddComponentFunction"
Public PriceC As Double


Public Function Add_Component(compname As String, addname As Double)
'Updates
ThisWorkbook.Connections("Query - Custom Prices").Refresh

On Error GoTo Errhandler

If addname = 1 Then GoTo Skip

errNum = 1
compname = Application.InputBox( _
Title:="Component Name", _
Prompt:="What is the Component Name?", _
Type:=1 + 2)

If compname = "0" Or compname = "False" Then GoTo EndSub

Skip:
errNum = 2
PriceC = Application.InputBox( _
Title:="Component Price", _
Prompt:="What is the Price of the Component for " & compname & " ?", _
Type:=1)

If VarType(PriceC) = 11 Then GoTo EndSub

errNum = 3
PODate = Application.InputBox( _
Title:="PO Date", _
Prompt:="What was the PO Date(Format:XX/XX/XXXX)?", _
Type:=1 + 2)

If PODate = "0" Or PODate = "False" Then GoTo EndSub

Dim wb As Workbook

'Opens BOM workbook
Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/BOMsForHoses.xlsx")
Sheets("Component Pricing").Select
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Removing front of partnames
If Left(compname, 6) = "OPINV:" Then
compnamef = Right(compname, Len(compname) - 6)
Else
compnamef = compname
End If

  With Workbooks("BOMsForHoses.xlsx").Worksheets("Component Pricing")
    .Range(Cells(CDbl(lRow + 1), 1).Address).Value = compnamef
    .Range(Cells(CDbl(lRow + 1), 2).Address).Value = PriceC
    .Range(Cells(CDbl(lRow + 1), 3).Address).Value = PODate

  End With

wb.Close True 'save changes
ThisWorkbook.Connections("Query - Custom Prices").Refresh
    
GoTo EndSub

Errhandler:

If errNum = 1 Then
MsgBox ("comp " & errNum)
wb.Close False 'does not save changes
GoTo EndSub
End If

If errNum = 2 Then
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
