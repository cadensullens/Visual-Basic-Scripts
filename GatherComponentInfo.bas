Attribute VB_Name = "GatherComponentInfo"
Function Gather_Component_Info(PartNames() As String)


'Declaring Variables for use
Dim table As ListObject
Dim table2 As ListObject
Dim table3 As ListObject
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet

'Setting Variables to values
Set ws = Worksheets("Short Parts")
Set table = ws.ListObjects("Detail")
Set ws2 = Worksheets("Purchase Order Tracker")
Set table2 = ws2.ListObjects("Purchase_Order_Tracker")
Set ws3 = Worksheets("TiteFlex Pricing")
Set table3 = ws3.ListObjects("TiteFlex_Pricing")
Set ws4 = Worksheets("Qb inventory")
Set table4 = ws4.ListObjects("Inventory")
Set ws5 = Worksheets("Custom Prices")
Set table5 = ws5.ListObjects("Custom_Prices")

'Error Variables
Dim errMsg As String
Dim errNum As Integer

On Error GoTo Errhandler
OldPrices = ""
OldPriceQTY = 0
'sets user input to a variable
Start1:
stringDate = CStr(CDbl(CDate(DueDate)))

 If breakCount <> 0 Then
 ReDim PriceBreaks(1 To UBound(PartNames), 1 To breakCount)
 End If


For i = LBound(PartNames) To UBound(PartNames)
'This equation below is the equivalent of what is happening below
'the two evaluate functions are checking the columns assigned for a matching string
'The date is looking for less than or equal to capture everything from the due date back to present
'=SUM(FILTER(Detail[Quantity],("R160-6"=Detail[Component PN]) *(O10>=Detail[DUE DATE])))

'Removing front of partnames
If Left(PartNames(i), 6) = "OPINV:" Then
PartNames(i) = Right(PartNames(i), Len(PartNames(i)) - 6)
End If
'Short parts function

     errNum = 1
     dateCheck = ws.Evaluate(table.ListColumns(8).DataBodyRange.Address & "<=" & stringDate)
     CheckPart = ws.Evaluate(table.ListColumns(12).DataBodyRange.Address & "=""" & PartNames(i) & """")
     
     Dim Bool() As Double
     Dim check As Double
     For j = LBound(dateCheck) To UBound(dateCheck)
        If dateCheck(j, 1) = False Then
        check = 0
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        Else
        check = 1
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        End If
        Next j
        
    If Application.WorksheetFunction.Sum(Bool) = 0 Then GoTo Errhandler
    
     Dim Bool2() As Double
     Dim check2 As Double
     For j = LBound(CheckPart) To UBound(CheckPart)
        If CheckPart(j, 1) = False Then
        check2 = 0
        ReDim Preserve Bool2(1 To j)
        Bool2(j) = check2
        Else
        check2 = 1
        ReDim Preserve Bool2(1 To j)
        Bool2(j) = check2
        End If
        Next j
    
    If Application.WorksheetFunction.Sum(Bool2) = 0 Then GoTo Errhandler
        
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
    
Start2:
    
     'Backlog on Order
     'This runs exactly the same as above function
      errNum = 2
     Ordered = ws2.Evaluate(table2.ListColumns(5).DataBodyRange.Address & "=""" & PartNames(i) _
     & """")
     BackDate = ws2.Evaluate(table2.ListColumns(9).DataBodyRange.Address & "<=" & stringDate)
     
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
        List2 = CDec(table2.ListColumns(6).Range(j + 1))
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
     
Start4:
    
     errNum = 4
     'Have to add OPINV inventory suffix for searhcing QB sheet
     qbName = "OPINV:" + PartNames(i)
     onHand = Application.WorksheetFunction.VLookup(qbName, table4.Range.Columns("A:B"), 2, False)
     ReDim Preserve onHandList(1 To i)
     onHandList(i) = Round(onHand, 2)
     
Start3:
    
     'Titeflex Pricing finds price on that sheet
      'Vlookup as sheet does not have duplicate P/N
     errNum = 3
     Price = Application.WorksheetFunction.VLookup(PartNames(i), table3.Range.Columns("A:F"), 4, False)
     ReDim Preserve PriceList(1 To i)
     PriceList(i) = Price
     GoTo Leadtime
     
CustomPrice:
     errMsg = "Component " & PartNames(i) & " is NOT on the Custom component pricing Sheet, Confirm Spelling of Part and Date. If correct, then Part is not on the Custom component pricing Sheet."
     errNum = 31
     Price = Application.WorksheetFunction.VLookup(PartNames(i), table5.Range.Columns("A:C"), 2, False)
     PoDate = Application.WorksheetFunction.VLookup(PartNames(i), table5.Range.Columns("A:C"), 3, False)
     
        If CDate(PoDate) < (Date - 90) Then
           Response = MsgBox("Pricing for " & PartNames(i) & " is from " & CDate(PoDate) & ". Do you want to update Pricing?", vbYesNo, "Update Pricing")
           If Response = 6 Then
               CustomComp.Caption = "Component: " & PartNames(i)
               CustomComp.CompNameEntry.Value = PartNames(i)
               CustomComp.Show
               ReDim Preserve PriceList(1 To i)
               PriceList(i) = PriceC
               Leadtime = 0
               ReDim Preserve LeadTimeList(1 To i)
                LeadTimeList(i) = Leadtime
               GoTo Continue
           Else
           'original start for podate
               If OldPrices = "" Then
               OldPrices = PartNames(i)
               OldPriceQTY = 1
               Else
               OldPrices = OldPrices & ", " & PartNames(i)
               OldPriceQTY = OldPriceQTY + 1
               End If
           End If
         End If
        
     ReDim Preserve PriceList(1 To i)
     PriceList(i) = Price
     Leadtime = 0
     ReDim Preserve LeadTimeList(1 To i)
     LeadTimeList(i) = Leadtime
     
     GoTo Continue
     
Leadtime:
     'TiteFlex Leadtime
     'Vlookup as sheet does not have duplicate P/N
     Leadtime = Application.WorksheetFunction.VLookup(PartNames(i), table3.Range.Columns("A:F"), 5, False)
     ReDim Preserve LeadTimeList(1 To i)
     LeadTimeList(i) = Leadtime
     
 
 GoTo Continue
 
Errhandler:
    
    
    If errNum = 0 Then
    MsgBox errMsg
    Gathererr = 1
    GoTo EndProc
    End If
    
    If errNum = 1 Then
    ShortPart = 0
    ReDim Preserve ShortPartList(1 To i)
     ShortPartList(i) = ShortPart
    GoTo Start2
    End If
    
    If errNum = 2 Then
    Backlog = 0
    ReDim Preserve BacklogList(1 To i)
     BacklogList(i) = Backlog
    GoTo Start4
    End If
    
    If errNum = 3 Then
    Resume CustomPrice
    End If
    
    If errNum = 31 Then
    MsgBox errMsg
    Response = MsgBox("Do you want to add " & PartNames(i) & " pricing now?", vbYesNo, "Add Price for Component")
        If Response = 6 Then
        CustomComp.Caption = "Component: " & PartNames(i)
        CustomComp.CompNameEntry.Value = PartNames(i)
        CustomComp.Show
        ReDim Preserve PriceList(1 To i)
         PriceList(i) = PriceC
        Leadtime = 0
        ReDim Preserve LeadTimeList(1 To i)
         LeadTimeList(i) = Leadtime
        Resume Continue
    Else
        Price = 0
        ReDim Preserve PriceList(1 To i)
         PriceList(i) = Price
        Leadtime = 0
        ReDim Preserve LeadTimeList(1 To i)
         LeadTimeList(i) = Leadtime
        Resume Continue
        End If
    End If
    
    If errNum = 4 Then
    onHand = 0
    ReDim Preserve onHandList(1 To i)
     onHandList(i) = onHand
    Resume Start3
    End If
    
Continue:

errNum = 0
errMsg = "There was an error in finding information for this Hose. Please Try Again."

    'Creates qtys for price break points amounts
    For j = 1 To breakCount

    PriceBreaks(i, j) = (CDbl(Backlog) + CDbl(onHand)) - (CDbl(ShortPart) + (partQty(j) * compQTY(i)))

    Next j
    
     ReDim Preserve Grand(1 To i)
     Grand(i) = compQTY(i) * Round(Price, 2)
     ReDim Preserve LongLead(1 To i)
     LongLead(i) = Leadtime


Next i

    If hose <> "" Then
         PartInfo.partname.Caption = "Hose:" & " " & hose
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
    PartInfo.Grandtext.Value = Grandsum
    
    'Find largest value
    max = LongLead(1) 'set the first leadtime as the max
    For j = LBound(PartNames) To UBound(PartNames)
    If LongLead(j) >= max Then max = LongLead(j) 'if another element is larger, then it is the max
    Next j
    PartInfo.Longest.Value = max & " Weeks"
    
    PartInfo.SpecialClean.Value = SpecClean
    
    If OldPriceQTY > 1 Then
        OldPriceText = "Component Price for " & OldPrices & " are Older Than 90 Days, Please Review Validity for accurate quote."
    ElseIf OldPriceQTY = 1 Then
        OldPriceText = "Component Price for " & OldPrices & " is Older Than 90 Days, Please Review Validity for accurate quote."
    Else
        OldPriceText = ""
    End If
    
EndProc:

End Function


