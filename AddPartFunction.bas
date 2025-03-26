Attribute VB_Name = "AddPartFunction"
Public buildHose() As String
Public buildName As String

Function Add_Part(buildHose() As String)

For i = LBound(buildHose) To UBound(buildHose)
hose = buildHose(i)

    Response = MsgBox("Hose, " & hose & ", not found on BOM list or Buy/Sell. Would you like to enter the Hose now?", vbYesNo, "Hose Not Found")
    'check response from Msgbox to run a function
        If Response = 6 Then
        'for skipping name ask again in build function
        BuildSkip = 1
        'Determine Build Setup as Maker or Buy/Sell
        Response = MsgBox("Click 'Yes' for Maker, Click 'No' for Buy/Sell", vbYesNo, "Choose Build Type")
        
            If Response = 6 Then
            copyTemp = 3
            Call Build_Comp
            'decrease Number hose for later functions
'            NumberHose = NumberHose - 1
            BuildSkip = 0
            copyTemp = 3
            Miss = Miss - 1
            Else
            copyTemp = 3
            Call BuySell_Update
            'decrease Number hose for later functions
'            NumberHose = NumberHose - 1
            BuildSkip = 0
            copyTemp = 3
            Miss = Miss - 1
            End If
        
        'Code for 'No' response
        Else
        
        GoTo EndSub
        
        
        End If
Next i
EndSub:
End Function
