Attribute VB_Name = "CheckBOMFunction"
Public CheckBOMerr As Double

Function Check_BOM(hose As String)
'Reset CheckBOMerr variable
CheckBOMerr = 0

Dim table As ListObject
Dim ws As Worksheet
Set ws = Worksheets("BOM Master")
Set table = ws.ListObjects("BOMMaster")

On Error GoTo Errhandler
errNum = 3
WireHole = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:B"), 2, False)
BarbRoy = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:C"), 3, False)
For i = 1 To 10
    Build = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:W"), 2 + i * 2, False)
    QTY = Application.WorksheetFunction.VLookup(hose, table.Range.Columns("A:W"), 3 + i * 2, False)
    Dim index As Long
    
      index = InStr(Build, ":")
      
      Build = Mid(Build, index + 1)
errNum = 4
    If Build = "" Then Exit For
    ReDim Preserve PartNames(1 To i)
        PartNames(i) = Build
    ReDim Preserve compQTY(1 To i)
        compQTY(i) = QTY
Next i

errNum = 0
GoTo EndProc:

Errhandler:


If errNum = 3 Then
MsgBox (CStr(errNum) & " CheckBOM")
CheckBOMerr = 1
GoTo EndProc
End If

If errNum = 4 Then
MsgBox (CStr(errNum) & " CheckBOM")
CheckBOMerr = 1
GoTo EndProc
End If

EndProc:
End Function
