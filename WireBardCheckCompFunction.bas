Attribute VB_Name = "WireBardCheckCompFunction"
Function WireHole_Barb()

On Error GoTo Errhandler
'Check for hose on BOM
errNum = 1
WireHole = Application.InputBox( _
Title:="Components Count", _
Prompt:="How many WireHoles are there?", _
Type:=1)
If VarType(WireHole) = 11 Then GoTo EndSub

errNum = 2
BarbRoy = Application.InputBox( _
Title:="Barb Royalty", _
Prompt:="What is the Barb Royalty Amount?", _
Type:=1)
If VarType(BarbRoy) = 11 Then GoTo EndSub
GoTo EndSub

Errhandler:
MsgBox (errNum & " wire")
EndSub:
End Function
