Attribute VB_Name = "AskQtyFunction"
Public comp As Variant

Function Ask_Qty(PartNames() As String)

For i = 1 To UBound(PartNames)
errNum = 1
comp = Application.InputBox( _
Title:="Component Quantity", _
Prompt:="How much of " & PartNames(i) & " will be used?", _
Type:=1)
If VarType(comp) = 11 Then GoTo EndSub


ReDim Preserve compQTY(1 To i)
compQTY(i) = comp
Next i
GoTo EndSub
Errhandler:
MsgBox (errNum & " Ask Qty")

EndSub:

End Function
