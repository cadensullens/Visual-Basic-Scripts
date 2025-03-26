Attribute VB_Name = "saveStatusFunction"


Function saveStatus(Po As String, locR As Variant, locC As Variant, newName As String)
On Error GoTo Errhandler
Dim errNum As Double

errNum = 1

If copyTemp = 1 Then
'Goes to template sheet and copies the formatted table
Worksheets("Template").Range("A18:F18").copy
Set TargetRange = Worksheets(newName).Range(Cells(CDbl(locR), CDbl(locC)).Address)
TargetRange.PasteSpecial xlPasteAll

For i = 1 To POHits

    If i Mod 2 = 0 Then
    Index = locR + i
    Worksheets("Template").Range("A20:F20").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(Index), CDbl(locC)).Address)
    TargetRange.PasteSpecial xlPasteAll
    Else
    Index = locR + i
    Worksheets("Template").Range("A19:F19").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(Index), CDbl(locC)).Address)
    TargetRange.PasteSpecial xlPasteAll
    End If
Next i
End If

errNum = 3
For i = 1 To POHits
    With Worksheets(newName)

      Index = locR + i
     'PO #
     .Range(Cells(CDbl(Index), CDbl(locC)).Address).Value = Po
     'SO Number
     .Range(Cells(CDbl(Index), CDbl(locC + 1)).Address).Value = SONumber(i)
     'Customer Date
     .Range(Cells(CDbl(Index), CDbl(locC + 2)).Address).Value = CustDate(i)
     'Completion/Recovery Date
     .Range(Cells(CDbl(Index), CDbl(locC + 3)).Address).Value = CompDate(i)
     'Qty
     .Range(Cells(CDbl(Index), CDbl(locC + 4)).Address).Value = BuildQty(i)
     'Job Status
     .Range(Cells(CDbl(Index), CDbl(locC + 5)).Address).Value = JobStat(i)
     
    End With
Next i

     
errNum = 4
'Makes columns correct size
Worksheets(newName).Columns("A:R").AutoFit

GoTo EndProc

Errhandler:
If errNum = 1 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

If errNum = 3 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

If errNum = 4 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

EndProc:
errNum = 4
'Makes columns correct size
Worksheets(newName).Columns("A:R").AutoFit
CopyCheck = 0
End Function
