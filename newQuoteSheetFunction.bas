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
For Each sh In ActiveWorkbook.Sheets
    If LCase(sh.Name) = LCase(newName) Then
        MsgBox ("Sheet name is a repeat, Please Enter a Unique name for the Sheet.")
        GoTo Startn
    End If
Next

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

errNum = 2
Dim Button2 As Shape
Set Button2 = Worksheets(newName).Shapes.AddShape(msoShapeRoundedRectangle, 10, 5, 175, 30)

'Adds Bevel to the Shape
Button2.ThreeD.BevelTopType = msoBevelSoftRound

'Adds text and formatting
With Button2.TextFrame2.TextRange
    .Text = "Add Quote to Metric"
    .Font.Bold = msoTrue
    .Font.Fill.ForeColor.RGB = RGB(256, 256, 256)
    .Font.Size = 18
    '.Font.Shadow = True
End With
Button2.Left = 175
'Center Alignment
  Button2.TextFrame.HorizontalAlignment = xlHAlignCenter
  Button2.TextFrame.VerticalAlignment = xlVAlignCenter
'Makes sure color is Theme Matched
Button2.Fill.ForeColor.RGB = RGB(165, 181, 146)
'Assigns Macro to button
Button2.OnAction = "QuoteMetric.CallQuote"

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
MsgBox (errNum & " new Quote")
GoTo EndProc:
End If

If errNum = 2 Then
'Finish this up
MsgBox (errNum & " new Quote")
GoTo EndProc:
End If

EndProc:
End Function
