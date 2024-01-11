Attribute VB_Name = "PriceBreaksFunction"
Function PriceBreaksFunc()

Dim PartBreak As String
Dim break As Double
'Reset until changed in If statement
priceend = 0

On Error GoTo Errhandler
Start:
errNum = 0
PartBreak = MsgBox("Do you want to enter in Part/Price Breaks(yes/no)?", vbYesNo, "Part Breaks")

If PartBreak <> 6 Then
    For i = 1 To breakCount
    partQty(i) = 0
    Next i
    breakCount = 0
    GoTo EndSub
End If
    

Start3:
errNum = 3
breakCount = Application.InputBox( _
Title:="Part Breaks Amount", _
Prompt:="How many Part/Price Breaks do you have?", _
Type:=1)
test = VarType(breakCount)
If breakCount = False Then GoTo EndSub
errNum = 5

Start4:
errNum = 4

For i = 1 To breakCount
    break = Application.InputBox( _
    Title:="Price Break" & " " & i, _
    Prompt:="Enter in Part/Price Break Qty", _
    Type:=1)
    If break = False Then GoTo EndSub
    ReDim Preserve partQty(1 To i)
    partQty(i) = break
    Next i
    
errNum = 0
GoTo EndSub

Errhandler:

If errNum = 0 Then
MsgBox ("Please Click 'Yes' or 'No'")
GoTo Start
End If

If errNum = 3 Then
MsgBox ("Enter Part/Price break Total as a Number only")
GoTo Start3
End If

If errNum = 4 Then
MsgBox ("Enter Part/Price break values as a Number only")
GoTo Start4
End If

If errNum = 5 Then
MsgBox ("Cannot have more thatn 6 Price Breaks. Please enter a number 6 or lower")
GoTo Start3
End If


EndSub:

End Function
