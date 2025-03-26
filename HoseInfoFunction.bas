Attribute VB_Name = "HoseInfoFunction"
Public HoseErr As Double
Public BuySell As Double
Public hoseNames() As String
Public BuildSkip As Double
Public Miss As Double


Function HoseInfo(hose As String)

HoseErr = 0
BuySell = 0
Count = 0
Miss = 0

Dim table As ListObject
Dim ws As Worksheet
'Dim Count As Double


Set ws = Worksheets("BOM Master")
Set table = ws.ListObjects("BOMMaster")
Dim table1 As ListObject
Dim ws1 As Worksheet

Set ws1 = Worksheets("Buy-Sell")
Set table1 = ws1.ListObjects("BuySell")

On Error GoTo Errhandler
errNum = 1

'Check for hose on BOM
For j = 1 To Len(hose)
Character = Mid(hose, j, 1)
If Character Like "[a-zA-Z-]" Then GoTo StringCheck
Next j

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheck

StringCheck:

    HoseCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")

GoTo BoolCheck

DoubleCheck:
    HoseCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=" & CDbl(hose) & "")

errNum = 1
BoolCheck:
     Dim Bool() As Double
     Dim check As Double
     For j = LBound(HoseCheck) To UBound(HoseCheck)
        If HoseCheck(j, 1) = False Then
        check = 0
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        Else
        check = 1
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        End If
        Next j
  errNum = 3
If Application.WorksheetFunction.Sum(Bool) = 0 Then

For j = 1 To Len(hose)
Character = Mid(hose, j, 1)
If Character Like "[a-zA-Z-]" Then GoTo StringCheckBuy
Next j

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheckBuy

StringCheckBuy:

    HoseCheck1 = ws1.Evaluate(table1.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")

GoTo BoolCheckBuy

DoubleCheckBuy:
    HoseCheck1 = ws1.Evaluate(table1.ListColumns(1).DataBodyRange.Address & "=" & CDbl(hose) & "")

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
If Application.WorksheetFunction.Sum(Bool1) = 0 Then GoTo Errhandler
End If

errNum = 5
Count = Count + 1
ReDim Preserve hoseNames(1 To Count)
hoseNames(Count) = hose
HoseErr = 0

GoTo EndSub

Errhandler:
If errNum = 1 Then
MsgBox (errNum & " HoseInfo")
GoTo EndSub
End If

If errNum = 2 Then
MsgBox (errNum & " HoseInfo")
GoTo EndSub
End If

If errNum = 5 Then
MsgBox (errNum & " HoseInfo")
GoTo EndSub
End If

If errNum = 4 Then
HoseErr = 1
GoTo EndSub
End If


EndSub:
End Function
