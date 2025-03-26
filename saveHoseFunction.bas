Attribute VB_Name = "saveHoseFunction"
Function saveHose(HoseName As String, locR As Variant, locC As Variant, newName As String)

On Error GoTo Errhandler
CopyCheck = 0
errNum = 1
If BuySell <> 1 Then

For i = LBound(PartNames) To UBound(PartNames)

errNum = 3
    With Worksheets(newName)

      Index = locR + 2 + i
     'Enter in component names in table
     .Range(Cells(CDbl(Index), CDbl(locC)).Address).Value = PartNames(i)
     'Qty to build hose
     .Range(Cells(CDbl(Index), CDbl(locC + 1)).Address).Value = compQTY(i)
     'Price info for Hose
     .Range(Cells(CDbl(Index), CDbl(locC + 2)).Address).Value = "$" & Round(PriceList(i), 2)
     'On hand value for component
     .Range(Cells(CDbl(Index), CDbl(locC + 3)).Address).Value = Round(onHandList(i), 2)
     'On order information for component
     .Range(Cells(CDbl(Index), CDbl(locC + 4)).Address).Value = BacklogList(i)
     'Current orders parts claimed/needed
     .Range(Cells(CDbl(Index), CDbl(locC + 5)).Address).Value = Round(ShortPartList(i), 2)
     'Difference to determine if enough margin to build hose
     DiffValue = Round((CDbl(BacklogList(i)) + CDbl(onHandList(i))) - CDbl(ShortPartList(i)), 2)
     .Range(Cells(CDbl(Index), CDbl(locC + 6)).Address).Value = DiffValue

     'Time to get new components
     .Range(Cells(CDbl(Index), CDbl(locC + 7)).Address).Value = LeadTimeList(i)
     
     
     'For Loop to enter in Price Break amounts
     For j = 1 To breakCount
     
     .Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2 + i, 7 + j).Value = PriceBreaks(i, j)
     
     Next j
     
    End With
Next i

For i = LBound(PartNames) To UBound(PartNames)
 With Worksheets(newName)
 
    For j = 1 To breakCount
     If PriceBreaks(i, j) < 0 Then
        .Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2 + i, 7 + j).Interior.Color = RGB(255, 199, 206)
        .Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2 + i, 7 + j).Font.Color = RGB(156, 0, 6)
     End If
     Next j
     
    Index = locR + 2 + i
    DiffValue = .Range(Cells(CDbl(Index), CDbl(locC + 6)).Address).Value
    If DiffValue < 0 Then
       .Range(Cells(CDbl(Index), CDbl(locC + 6)).Address).Interior.Color = RGB(255, 199, 206)
        .Range(Cells(CDbl(Index), CDbl(locC + 6)).Address).Font.Color = RGB(156, 0, 6)
    End If
    End With
    
Next i
    
'add breakcounts in
For i = 1 To breakCount
    Index = locR + 2 + i
    With Worksheets(newName)
    .Range(Cells(CDbl(Index), CDbl(locC + 9 + breakCount)).Address).Value = partQty(i)
    End With
    
Next i

If breakCount < 4 Then
marginRepeat = 4
Else
marginRepeat = breakCount
End If

'expands margin table
For i = 1 To marginRepeat
    With Worksheets(newName)
    Index = locR + 2 + i

    If i = 1 Then
    .Range(Cells(CDbl(Index), CDbl(locC + 10 + breakCount)).Address).Value = MarginStart / 100
    .Range(Cells(CDbl(Index), CDbl(locC + 11 + breakCount)).Address).Formula2 = ("=" & Cells(CDbl(locR + 1), CDbl(locC + 10 + breakCount)).Address & "/" & "(1- [@[MM%]])")
    Else
    .Range(Cells(CDbl(Index), CDbl(locC + 10 + breakCount)).Address).Formula2 = "=" & Cells(CDbl(Index - 1), CDbl(locC + 10 + breakCount)).Address & "-" & (Increments / 100)
    .Range(Cells(CDbl(Index), CDbl(locC + 11 + breakCount)).Address).Formula2 = ("=" & Cells(CDbl(locR + 1), CDbl(locC + 10 + breakCount)).Address & "/" & "(1- [@[MM%]])")
    End If
    
    End With
Next i



With Worksheets(newName)
        'All this needs to be fixed based on gather info

     .Range(Cells(CDbl(locR), CDbl(locC + 1)).Address).Value = HoseName
     
     If LCase(SpecClean) = "yes" Then
     '=SUM(Bxx:Bxx*Cxx:Cxx) + (10 * wirehole amount) + BarbRoyalty amount
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 1)).Address).Formula2 = "=SUM(" & Cells(CDbl(locR + 3), CDbl(locC + 1)).Address & ":" & Cells(CDbl(locR + 2 + UBound(PartNames)), _
     CDbl(locC + 1)).Address & "*" & Cells(CDbl(locR + 3), CDbl(locC + 2)).Address & ":" & Cells(CDbl(locR + 2 + UBound(PartNames)), CDbl(locC + 2)).Address & ") +" & Cells(CDbl(locR + 1), CDbl(locC + 7)).Address & "+" & (10 * WireHole) + BarbRoy
     Else
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 1)).Address).Formula2 = "=SUM(" & Cells(CDbl(locR + 3), CDbl(locC + 1)).Address & ":" & Cells(CDbl(locR + 2 + UBound(PartNames)), _
     CDbl(locC + 1)).Address & "*" & Cells(CDbl(locR + 3), CDbl(locC + 2)).Address & ":" & Cells(CDbl(locR + 2 + UBound(PartNames)), CDbl(locC + 2)).Address & ") +" & (10 * WireHole) + BarbRoy
     End If
     
     If CDate(DueDate) = "12/12/9999" Then
     .Range(Cells(CDbl(locR), CDbl(locC + 3)).Address) = ""
     Else
     .Range(Cells(CDbl(locR), CDbl(locC + 3)).Address) = CDate(DueDate)
     End If
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 3)).Address).Value = max & " Weeks"
     
     .Range(Cells(CDbl(locR), CDbl(locC + 5)).Address) = LeadEntry & " Weeks"
     
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 5)).Address) = SpecClean
     
     If SpecClean = "Yes" Then
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 5)).Address).Interior.Color = RGB(255, 127, 127)
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 7)).Address).Value = CleanCustomPrice
     End If
     
End With
Else

With Worksheets(newName)
     'Inserts Info for each new hose in blocks above table
     .Range(Cells(CDbl(locR), CDbl(locC + 1)).Address).Value = hose
     
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 1)).Address).Value = PriceBS
     
     .Range(Cells(CDbl(locR), CDbl(locC + 2)).Address).Value = "Quote Date"
     
     .Range(Cells(CDbl(locR), CDbl(locC + 3)).Address) = CDate(QuoteDate)
     
     .Range(Cells(CDbl(locR + 3), CDbl(locC)).Address).Value = "Max LeadTime"
     
     .Range(Cells(CDbl(locR + 3), CDbl(locC + 1)).Address).Value = LeadtimeBS
     
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 2)).Address).Value = "Valid Until:"
     
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 3)).Address).Value = Expire
     
     If CDate(Expire) < Date Then
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 3)).Address).Interior.Color = RGB(255, 127, 127)
     End If
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC)).Address).Value = "Vendor"
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 1)).Address).Value = Vendor
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 2)).Address).Value = "Quantity Quoted"
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 3)).Address).Value = MOQ
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 5)).Address).Value = MOQ
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 7)).Address).Formula2 = ("=" & Cells(CDbl(locR), CDbl(locC + 6)).Address & "/" & "(1- [@[MM%]])")
     
End With
End If

GoTo EndProc

Errhandler:
If errNum = 1 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

If errNum = 3 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

If errNum = 4 Then
MsgBox (CStr(errNum) & " SaveHose")
GoTo EndProc
End If

EndProc:
errNum = 4
'Makes columns correct size
Worksheets(newName).Columns("A:R").AutoFit

End Function
