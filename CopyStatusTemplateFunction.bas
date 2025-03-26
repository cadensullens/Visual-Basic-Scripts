Attribute VB_Name = "CopyStatusTemplateFunction"
Function CopyStatus_Template()

On Error GoTo Errhandler
' variable from Part info 0 makes a new sheet for saving the data 1 uses an existing sheet
If copyTemp = 0 Then
GoTo Errhandler
ElseIf copyTemp = 1 Then
GoTo Start2
ElseIf copyTemp = 2 Then
GoTo Finish
End If


Start2:
'For the add to existing sheet function from Partinfo user form
Set location = Application.InputBox( _
Title:="Cell select", _
Prompt:="Select the cell to begin Template", _
Type:=8)

Finish:
errNum = 6
If copyTemp = 1 Then
'get the row number from the location pick above
locR = location.Row
locC = location.Column
'Gets the sheet name
newName = location.Worksheet.Name

GoTo saveStatus

ElseIf copyTemp = 2 Then

Call newStatusSheet
locR = 4
locC = 1
'Goes to template sheet and copies the formatted table
Worksheets("Template").Range("A18:F18").copy
Set TargetRange = Worksheets(newName).Range("A4")
TargetRange.PasteSpecial xlPasteAll

For i = 1 To POHits
    If i Mod 2 = 0 Then
    Worksheets("Template").Range("A20:F20").copy
    Set TargetRange = Worksheets(newName).Range("A4").Offset(i, 0)
    TargetRange.PasteSpecial xlPasteAll
    Else
    Worksheets("Template").Range("A19:F19").copy
    Set TargetRange = Worksheets(newName).Range("A4").Offset(i, 0)
    TargetRange.PasteSpecial xlPasteAll
    End If
Next i
GoTo saveStatus
End If
 
saveStatus:

Call saveStatus(PO, locR, locC, newName)


GoTo EndSub

Errhandler:


EndSub:
'Ends the copy mode of the selected cells
Application.CutCopyMode = False
End Function
