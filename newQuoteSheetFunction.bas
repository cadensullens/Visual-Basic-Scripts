Attribute VB_Name = "newQuoteSheetFunction"
Function newQuoteSheet()

On Error GoTo Errhandler
Startn:
newName = Application.InputBox( _
Title:="Name of New Sheet", _
Prompt:="Type the name of the new sheet", _
Type:=1 + 2)

'making sure name is a string after input box
newName = CStr(newName)

'if cancel is entered then end new sheet creation
If newName = "False" Then GoTo EndProc

'Adding Button connected to macro
errNum = 1
Worksheets.Add(After:=Sheets("Button")).Name = newName
errNum = 2
Dim Button As Shape
Set Button = Worksheets(newName).Shapes.AddShape(msoShapeRoundedRectangle, 10, 5, 150, 30)

'Adds Bevel to the Shape
Button.ThreeD.BevelTopType = msoBevelSoftRound

'Adds text and formatting
With Button.TextFrame2.TextRange
    .Text = "Look up a Hose"
    .Font.Bold = msoTrue
    .Font.Fill.ForeColor.RGB = RGB(256, 256, 256)
    .Font.Size = 18
    '.Font.Shadow = True
End With
    
'Center Alignment
  Button.TextFrame.HorizontalAlignment = xlHAlignCenter
  Button.TextFrame.VerticalAlignment = xlVAlignCenter
'Makes sure color is Theme Matched
Button.Fill.ForeColor.RGB = RGB(165, 181, 146)
'Assigns Macro to button
Button.OnAction = "LookUpHose.Enter_Comp"

'Freeze button in place on sheet
With ActiveWindow
    If .FreezePanes Then .FreezePanes = False
    .SplitColumn = 0
    .SplitRow = 3
    .FreezePanes = True
End With

GoTo EndProc

Errhandler:
If errNum = 1 Then
'Finish this up
MsgBox ("Sheet name is a repeat, Please Enter a Unique name for the Sheet.")
ActiveSheet.Delete
GoTo Startn:
End If

If errNum = 2 Then
'Finish this up
MsgBox (errNum & " new Quote")
GoTo EndProc:
End If

EndProc:
End Function
