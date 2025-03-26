Attribute VB_Name = "saveStatusFunction"


Function saveStatus(PO As String, locR As Variant, locC As Variant, newName As String)
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
    index = locR + i
    Worksheets("Template").Range("A20:F20").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(index), CDbl(locC)).Address)
    TargetRange.PasteSpecial xlPasteAll
    Else
    index = locR + i
    Worksheets("Template").Range("A19:F19").copy
    Set TargetRange = Worksheets(newName).Range(Cells(CDbl(index), CDbl(locC)).Address)
    TargetRange.PasteSpecial xlPasteAll
    End If
Next i
End If

errNum = 3
For i = 1 To POHits
    With Worksheets(newName)

      index = locR + i
     'PO #
     .Range(Cells(CDbl(index), CDbl(locC)).Address).Value = PO
     'SO Number
     .Range(Cells(CDbl(index), CDbl(locC + 1)).Address).Value = SONumber(i)
     'Customer Date
     .Range(Cells(CDbl(index), CDbl(locC + 2)).Address).Value = CustDate(i)
     'Completion/Recovery Date
     .Range(Cells(CDbl(index), CDbl(locC + 3)).Address).Value = CompDate(i)
     'Qty
     .Range(Cells(CDbl(index), CDbl(locC + 4)).Address).Value = BuildQty(i)
     'Job Status
     .Range(Cells(CDbl(index), CDbl(locC + 5)).Address).Value = JobStat(i)
     
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
