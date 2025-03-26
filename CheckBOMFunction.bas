Attribute VB_Name = "CheckBOMFunction"
Public CheckBOMerr As Double

Function Check_BOM(hose As String)
'Reset CheckBOMerr variable
CheckBOMerr = 0
On Error GoTo Errhandler

'Dim wb As Workbook
errNum = 2
'Opens BOM workbook
'Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/Quoting/Quote Sheet Files/BOMsForHoses.xlsx")
'Sheets("BOM Master").Select
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Worksheets("BOM Master").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Worksheets("BOM Master").Cells(1, Columns.Count).End(xlToLeft).Column
    
    'wb.Close False 'does not save changes
    
Dim table As ListObject
Dim ws As Worksheet
Set ws = Worksheets("BOM Master")
Set table = ws.ListObjects("BOMMaster")

Col = (lCol - 3) / 2
lcoladd = Mid(Cells(1, lCol).Address, 2, 2)

'Check for hose on BOM
    For j = 1 To Len(hose)
    Character = Mid(hose, j, 1)
    If Character Like "[a-zA-Z-]" Then GoTo Stringer
    Next j

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheck
Stringer:
    Hoser = hose
GoTo Normal

DoubleCheck:
    Hoser = CDbl(hose)
Hoser = hose
    
Normal:
errNum = 3
WireHole = Application.WorksheetFunction.VLookup(Hoser, table.Range.Columns("A:B"), 2, False)
BarbRoy = Application.WorksheetFunction.VLookup(Hoser, table.Range.Columns("A:C"), 3, False)
SpecClean = Application.WorksheetFunction.VLookup(Hoser, table.Range.Columns("A:D"), 4, False)
For i = 1 To Col
    Build = Application.WorksheetFunction.VLookup(Hoser, table.Range.Columns("A:" & lcoladd & ""), 3 + i * 2, False)
    QTY = Application.WorksheetFunction.VLookup(Hoser, table.Range.Columns("A:" & lcoladd & ""), 4 + i * 2, False)
    
      Index = InStr(Build, ":")
      
      Build = Mid(Build, Index + 1)
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

If errNum = 2 Then
MsgBox (CStr(errNum) & " CheckBOM")
CheckBOMerr = 1
GoTo EndProc
End If

If errNum = 3 Then
Debug.Print "CheckBOM 3"
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
