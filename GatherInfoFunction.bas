Attribute VB_Name = "GatherInfoFunction"
Public Gathererr As Double

Function Gather_Info(hoseNames As String)

Gathererr = 0
If hoseNames = "" Then
Gathererr = 1
GoTo EndProcG

End If
'Buy sell has diff userform

Call Buy_Sell(hoseNames)
If BuySell = 1 Then
Call Buy_Sell_Fill(hoseNames)
GoTo EndProcG
End If

Call Check_BOM(hoseNames)

If CheckBOMerr = 1 Then
Gathererr = 1
GoTo EndProcG
End If




'Declaring Variables for use
Dim table As ListObject
Dim table2 As ListObject
Dim table3 As ListObject
Dim table5 As ListObject
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet
Dim ws5 As Worksheet


'Setting Variables to values
Set ws = Worksheets("Short Parts")
Set table = ws.ListObjects("Detail")
Set ws2 = Worksheets("TiteFlex Backlog")
Set table2 = ws2.ListObjects("Backlog")
Set ws3 = Worksheets("TiteFlex Pricing")
Set table3 = ws3.ListObjects("TiteFlex_Pricing")
Set ws4 = Worksheets("QB Inventory")
Set table4 = ws4.ListObjects("Inventory")
Set ws5 = Worksheets("Custom Prices")
Set table5 = ws5.ListObjects("Custom_Prices")

'Error Variables
Dim errMsg As String
Dim errNum As Integer


On Error GoTo Errhandler

'sets user input to a variable
Start1G:
stringDate = CStr(CDbl(CDate(DueDate)))


 If breakCount <> 0 Then
 ReDim PriceBreaks(1 To UBound(PartNames), 1 To breakCount)
 End If
 
 
 
For i = LBound(PartNames) To UBound(PartNames)
     
     'This equation below is the equivalent of what is happening below
    'the two evaluate functions are checking the columns assigned for a matching string
    'The date is looking for less than or equal to capture everything from the due date back to present
    '=SUM(FILTER(Detail[Quantity],("R160-6"=Detail[Component PN]) *(O10>=Detail[DUE DATE])))
     errMsg = "Component " & PartNames(i) & " is NOT on the Short Parts list, Confirm Spelling of Part and Date. If correct then, Part is not on Short Parts."
     errNum = 1
     dateCheck = ws.Evaluate(table.ListColumns(8).DataBodyRange.Address & "<=" & stringDate)
     CheckPart = ws.Evaluate(table.ListColumns(9).DataBodyRange.Address & "=""" & PartNames(i) & """")
     
     Dim BoolG1() As Double
     Dim checkG1 As Double
     For j = LBound(dateCheck) To UBound(dateCheck)
        If dateCheck(j, 1) = False Then
        checkG1 = 0
        ReDim Preserve BoolG1(1 To j)
        BoolG1(j) = checkG1
        Else
        checkG1 = 1
        ReDim Preserve BoolG1(1 To j)
        BoolG1(j) = checkG1
        End If
        Next j
        
    If Application.WorksheetFunction.Sum(BoolG1) = 0 Then GoTo Errhandler
    
     Dim BoolG2() As Double
     Dim checkG2 As Double
     For j = LBound(CheckPart) To UBound(CheckPart)
        If CheckPart(j, 1) = False Then
        checkG2 = 0
        ReDim Preserve BoolG2(1 To j)
        BoolG2(j) = checkG2
        Else
        checkG2 = 1
        ReDim Preserve BoolG2(1 To j)
        BoolG2(j) = checkG2
        End If
        Next j
    
    If Application.WorksheetFunction.Sum(BoolG2) = 0 Then GoTo Errhandler
        
     'Two new variables to run the For loop
     Dim together() As Double
     Dim List As Double
     'The For loop compares the two variables for matches of "true" for the same index
     'If it matches it will grab the value from the quantity cell and save the value
     'Otherwise it will save a 0 if not a match
     For j = LBound(CheckPart) To UBound(CheckPart)
        If CheckPart(j, 1) = True And dateCheck(j, 1) = True Then
        List = CDec(table.ListColumns(10).Range(j + 1))
        ReDim Preserve together(1 To j)
        together(j) = List
        Else
        List = 0
        ReDim Preserve together(1 To j)
        together(j) = List
        End If
        Next j
        
     'turns For Loop arrays (together, together2) into a sum for Short Parts
     ShortPart = Application.WorksheetFunction.Sum(together)
     'Adding to list for use in sheet creation function
     ReDim Preserve ShortPartList(1 To i)
     ShortPartList(i) = ShortPart
    
Start2G:
    
     'Backlog on Order
     'This runs exactly the same as above function
      errMsg = "Component " & PartNames(i) & " is NOT on the TiteFlex Backlog, Confirm Spelling of Part and Date. If correct then, Part is not on the TiteFlex Backlog."
      errNum = 2
     Ordered = ws2.Evaluate(table2.ListColumns(4).DataBodyRange.Address & "=""" & PartNames(i) _
     & """")
     BackDate = ws2.Evaluate(table2.ListColumns(8).DataBodyRange.Address & "<=" & stringDate)
     
     Dim Bool3() As Double
     Dim check3 As Double
     For j = LBound(Ordered) To UBound(Ordered)
        If Ordered(j, 1) = False Then
        check3 = 0
        ReDim Preserve Bool3(1 To j)
        Bool3(j) = check3
        Else
        check3 = 1
        ReDim Preserve Bool3(1 To j)
        Bool3(j) = check3
        End If
        Next j
        
    If Application.WorksheetFunction.Sum(Bool3) = 0 Then GoTo Errhandler
    
     Dim Bool4() As Double
     Dim check4 As Double
     For j = LBound(BackDate) To UBound(BackDate)
        If BackDate(j, 1) = False Then
        check4 = 0
        ReDim Preserve Bool4(1 To j)
        Bool4(j) = check4
        Else
        check4 = 1
        ReDim Preserve Bool4(1 To j)
        Bool4(j) = check4
        End If
        Next j
    
    If Application.WorksheetFunction.Sum(Bool4) = 0 Then GoTo Errhandler
    
      Dim together2() As Double
      Dim List2 As Double
      For j = LBound(Ordered) To UBound(Ordered)
        If Ordered(j, 1) = True And BackDate(j, 1) = True Then
        List2 = CDec(table2.ListColumns(5).Range(j + 1))
        ReDim Preserve together2(1 To j)
        together2(j) = List2
        Else
        List2 = 0
        ReDim Preserve together2(1 To j)
        together2(j) = List2
        End If
        Next j
         
    'turns For Loop arrays (together, together2) into a sum For Backlog
     Backlog = Application.WorksheetFunction.Sum(together2)
     
     ReDim Preserve BacklogList(1 To i)
     BacklogList(i) = Backlog


Start4G:
    
     errMsg = "Component " & PartNames(i) & " is NOT on the Inventory Sheet, Confirm Spelling of Part and Date. If correct, then Part is not on the Inventory Sheet."
     errNum = 4
     'Have to add OPINV inventory suffix for searhcing QB sheet
     qbName = "OPINV:" + PartNames(i)
     onHand = Application.WorksheetFunction.VLookup(qbName, table4.Range.Columns("A:B"), 2, False)
     onHand = Round(onHand, 2)
     ReDim Preserve onHandList(1 To i)
     onHandList(i) = onHand
     
Start3G:
    
     'Titeflex Pricing finds price on that sheet
      'Vlookup as sheet does not have duplicate P/N
     errMsg = "Component " & PartNames(i) & " is NOT on the TiteFlex pricing Sheet, Confirm Spelling of Part and Date. If correct, then Part is not on the TiteFlex pricing Sheet. The Custom Component Sheet will now be checked."
     errNum = 3
     Price = Application.WorksheetFunction.VLookup(PartNames(i), table3.Range.Columns("A:F"), 4, False)
     ReDim Preserve PriceList(1 To i)
     PriceList(i) = Price
     GoTo Leadtime
     
CustomPrice:
     errMsg = "Component " & PartNames(i) & " is NOT on the Custom component pricing Sheet, Confirm Spelling of Part and Date. If correct, then Part is not on the Custom component pricing Sheet."
     errNum = 31
     Price = Application.WorksheetFunction.VLookup(PartNames(i), table5.Range.Columns("A:C"), 2, False)
     
     ReDim Preserve PriceList(1 To i)
     PriceList(i) = Price
     Leadtime = 0
     ReDim Preserve LeadTimeList(1 To i)
     LeadTimeList(i) = Leadtime
     GoTo ContinueG
     
Leadtime:
     'TiteFlex Leadtime
     'Vlookup as sheet does not have duplicate P/N
     Leadtime = Application.WorksheetFunction.VLookup(PartNames(i), table3.Range.Columns("A:F"), 5, False)
     ReDim Preserve LeadTimeList(1 To i)
     LeadTimeList(i) = Leadtime
 
 GoTo ContinueG
 
Errhandler:
    MsgBox errMsg
    
    If errNum = 0 Then
    Gathererr = 1
    GoTo EndProcG
    End If
    
    If errNum = 1 Then
    ShortPart = 0
    ReDim Preserve ShortPartList(1 To i)
     ShortPartList(i) = ShortPart
    GoTo Start2G
    End If
    
    If errNum = 2 Then
    Backlog = 0
    ReDim Preserve BacklogList(1 To i)
     BacklogList(i) = Backlog
    GoTo Start4G
    End If
    
    If errNum = 3 Then
    Resume CustomPrice
    End If
    
    If errNum = 31 Then
    Response = MsgBox("Do you want to add " & PartNames(i) & " pricing now?", vbYesNo, "Add Price for Component")
        If Response = 6 Then
        Call Add_Component(PartNames(i), 1)
        ReDim Preserve PriceList(1 To i)
         PriceList(i) = PriceC
        Leadtime = 0
        ReDim Preserve LeadTimeList(1 To i)
         LeadTimeList(i) = Leadtime
        Resume ContinueG
        Else
        Price = 0
        ReDim Preserve PriceList(1 To i)
         PriceList(i) = Price
        Leadtime = 0
        ReDim Preserve LeadTimeList(1 To i)
         LeadTimeList(i) = Leadtime
        Resume ContinueG
        End If
     
    End If
    
    If errNum = 4 Then
    onHand = 0
    ReDim Preserve onHandList(1 To i)
     onHandList(i) = onHand
    Resume Start3G
    End If
    
ContinueG:

errNum = 0
errMsg = "There was an error in finding information for this Hose. Please Try Again."
        'Creates qtys for price break points amounts
    For j = 1 To breakCount

    PriceBreaks(i, j) = (CDbl(Backlog) + CDbl(onHand)) - (CDbl(ShortPart) + (partQty(j) * compQTY(i)))

    Next j
    
     ReDim Preserve Grand(1 To i)
     Grand(i) = compQTY(i) * Round(PriceList(i), 2)
     ReDim Preserve LongLead(1 To i)
     LongLead(i) = Leadtime

Next i

If CopyCheck = 1 Then
    'Final Sum for extra options
    Grandsum = Round(Application.WorksheetFunction.Sum(Grand), 2) + (10 * WireHole) + BarbRoy
    'Find largest value
    max = LongLead(1) 'set the first leadtime as the max
    For j = LBound(PartNames) To UBound(PartNames)
        If LongLead(j) >= max Then max = LongLead(j) 'if another element is larger, then it is the max
    Next j

Else
    'If hose was found on BOM then the name will be placed in
    If hoseNames <> "" Then
         PartInfo.partname.Caption = "Hose:" & " " & hoseNames
    End If
    'Shows dueDate entered
    If CDate(DueDate) = "12/12/9999" Then
    PartInfo.DateEnter.Value = ""
    Else
    PartInfo.DateEnter.Value = CDate(DueDate)
    End If
    
    If LeadEntry = "" Then
    PartInfo.Leadtime.Value = ""
    Else
    PartInfo.Leadtime.Value = LeadEntry & " Weeks"
    End If
    
    'Creates Value for Grand Total
    Grandsum = Round(Application.WorksheetFunction.Sum(Grand), 2) + (10 * WireHole) + BarbRoy
    PartInfo.Grand.Value = "$" & Grandsum
    
    'Find largest value
    max = LongLead(1) 'set the first leadtime as the max
    For j = LBound(PartNames) To UBound(PartNames)
    If LongLead(j) >= max Then max = LongLead(j) 'if another element is larger, then it is the max
    Next j
    PartInfo.Longest.Value = max & " Weeks"
    
End If
CopyCheck = 0
EndProcG:
End Function
