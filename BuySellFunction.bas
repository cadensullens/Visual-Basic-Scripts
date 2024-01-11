Attribute VB_Name = "BuySellFunction"
Public Vendor As String
Public PriceBS As String
Public LeadtimeBS As Double
Public QuoteDate As String
Public ValidFor As Double
Public Expire As String
Public MOQ As Double

Function Buy_Sell(hose As String)

BuySell = 0
Dim table As ListObject
Dim ws As Worksheet
Set ws = Worksheets("Buy-Sell")
Set table = ws.ListObjects("BuySell")

For j = 1 To Len(hose)
Character = Mid(hose, j, 1)
If Character Like "[a-zA-Z-]" Then GoTo StringCheckBuy
Next j
'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheckBuy

StringCheckBuy:

    HoseCheck1 = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")

GoTo BoolCheckBuy

DoubleCheckBuy:
    HoseCheck1 = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=" & CDbl(hose) & "")
    doublehose = 1

BoolCheckBuy:
     Dim Bool1() As Double
     Dim check1 As Double
     For j = LBound(HoseCheck1) To UBound(HoseCheck1)
        If HoseCheck1(j, 1) = False Then
        check1 = 0
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        Else
        check1 = 1
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        End If
        Next j
  errNum = 4
If Application.WorksheetFunction.Sum(Bool1) > 0 Then
BuySell = 1
Else
GoTo EndBS
End If

If doublehose = 1 Then
Vendor = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:F"), 2, False)
PriceBS = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:F"), 3, False)
LeadtimeBS = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:F"), 4, False)
QuoteDate = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:F"), 5, False)
ValidFor = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:F"), 6, False)
Expire = CDate(CDbl(QuoteDate) + ValidFor)
MOQ = Application.WorksheetFunction.VLookup(CDbl(hose), table.Range.Columns("A:G"), 7, False)
Else
Vendor = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:F"), 2, False)
PriceBS = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:F"), 3, False)
LeadtimeBS = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:F"), 4, False)
QuoteDate = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:F"), 5, False)
ValidFor = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:F"), 6, False)
Expire = CDate(CDbl(QuoteDate) + ValidFor)
MOQ = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:G"), 7, False)
End If
EndBS:
End Function
