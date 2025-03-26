Attribute VB_Name = "HoseInfoFunction"
Public HoseErr As Double
Public BuySell As Double
Public hoseNames() As String
Public BuildSkip As Double


Function HoseInfo()

HoseErr = 0
BuySell = 0
Count = 0

Dim table As ListObject
Dim ws As Worksheet
'Dim Count As Double


Set ws = Worksheets("BOM Master")
Set table = ws.ListObjects("BOMMaster")
Dim table1 As ListObject
Dim ws1 As Worksheet

Set ws1 = Worksheets("Buy-Sell")
Set table1 = ws1.ListObjects("BuySell")

On Error GoTo Errhandler
errNum = 1
NumberHose = Application.InputBox( _
Title:="Hose Count", _
Prompt:="How many Hoses are you looking up?", _
Type:=1)
If NumberHose = False Then GoTo EndSub

errNum = 2
For i = 1 To NumberHose
hose = Application.InputBox( _
Title:="Hose Name" & " " & i, _
Prompt:="What is the name of Hose #" & i & " ?", _
Type:=1 + 2)

If hose = "0" Or hose = "False" Then GoTo EndSub

'Check for hose on BOM
For j = 1 To Len(hose)
Character = Mid(hose, j, 1)
If Character Like "[a-zA-Z-]" Then GoTo StringCheck
Next j

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheck

StringCheck:

    HoseCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")

GoTo BoolCheck

DoubleCheck:
    HoseCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=" & CDbl(hose) & "")

errNum = 1
BoolCheck:
     Dim Bool() As Double
     Dim check As Double
     For j = LBound(HoseCheck) To UBound(HoseCheck)
        If HoseCheck(j, 1) = False Then
        check = 0
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        Else
        check = 1
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        End If
        Next j
  errNum = 3
If Application.WorksheetFunction.Sum(Bool) = 0 Then

For j = 1 To Len(hose)
Character = Mid(hose, j, 1)
If Character Like "[a-zA-Z-]" Then GoTo StringCheckBuy
Next j

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheckBuy

StringCheckBuy:

    HoseCheck1 = ws1.Evaluate(table1.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")

GoTo BoolCheckBuy

DoubleCheckBuy:
    HoseCheck1 = ws1.Evaluate(table1.ListColumns(1).DataBodyRange.Address & "=" & CDbl(hose) & "")

BoolCheckBuy:
     Dim Bool1() As Double
     Dim check1 As Double
     For j = LBound(HoseCheck1) To UBound(HoseCheck1)
        If HoseCheck1(j, 1) = False Then
        check1 = 0
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        Else
        check1 = 1
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        End If
        Next j
  errNum = 4
If Application.WorksheetFunction.Sum(Bool1) = 0 Then GoTo Errhandler
End If


Count = Count + 1
ReDim Preserve hoseNames(1 To Count)
hoseNames(Count) = hose
Start:
Next i

GoTo EndSub

Errhandler:
If errNum = 1 Then
MsgBox (errNum & " HoseInfo")
GoTo EndSub
End If

If errNum = 2 Then
MsgBox (errNum & " HoseInfo")
GoTo EndSub
End If

If errNum = 4 Then
    If NumberHose > 1 Then
    Response = MsgBox("Hose, " & hose & ", not found on BOM list or Buy/Sell. Would you like to enter the Hose now?", vbYesNo, "Hose Not Found")
    'check response from Msgbox to run a function
        If Response = 6 Then
        'for skipping name ask again in build function
        BuildSkip = 1
        'Determine Build Setup as Maker or Buy/Sell
        Response = MsgBox("Click 'Yes' for Maker, Click 'No' for Buy/Sell", vbYesNo, "Choose Build Type")
        
            If Response = 6 Then
            Call Build_Comp
            'decrease Number hose for later functions
            NumberHose = NumberHose - 1
            BuildSkip = 0
            HoseErr = 1
            Else
            Call BuySell_Update
            'decrease Number hose for later functions
            NumberHose = NumberHose - 1
            BuildSkip = 0
            HoseErr = 1
            End If
        
        'Code for 'No' response
        Else
        'decrease Number hose for later functions
        NumberHose = NumberHose - 1
        GoTo EndSub
        
        'For response = 6 for Hose not found message
        End If
        
    GoTo Start
    
    Else
    
    Response = MsgBox("Hose, " & hose & ", not found on BOM list or Buy/Sell. Would you like to enter the Hose now?", vbYesNo, "Hose Not Found")
    
    'check response from Msgbox to run a function
            If Response = 6 Then
            'for skipping name ask again in build function
            BuildSkip = 1
            'Determine Build Setup as Maker or Buy/Sell
            Response = MsgBox("Click 'Yes' for Maker, Click 'No' for Buy/Sell", vbYesNo, "Choose Build Type")
            
                If Response = 6 Then
                Call Build_Comp
                'decrease Number hose for later functions
                NumberHose = NumberHose - 1
                BuildSkip = 0
                HoseErr = 1
                Else
                Call BuySell_Update
                'decrease Number hose for later functions
                NumberHose = NumberHose - 1
                BuildSkip = 0
                HoseErr = 1
                End If
            
            Else
            HoseErr = 1
            GoTo EndSub
            End If
            
            GoTo EndSub
        'For NumberHose
        End If
'for errNum
End If


EndSub:
End Function
