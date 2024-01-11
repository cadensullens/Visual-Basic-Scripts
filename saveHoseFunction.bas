Attribute VB_Name = "saveHoseFunction"
Function saveHose(Hosename As String, locR As Variant, locC As Variant, newName As String)

On Error GoTo Errhandler
CopyCheck = 0
errNum = 1
If BuySell <> 1 Then
With Worksheets(newName)
        'All this needs to be fixed based on gather info

     .Range(Cells(CDbl(locR), CDbl(locC + 1)).Address).Value = Hosename
     '=SUM(Bxx:Bxx*Cxx:Cxx) + (10 * wirehole amount) + BarbRoyalty amount
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 1)).Address).Formula2 = "=SUM(" & Cells(CDbl(locR + 3), CDbl(locC + 1)).Address & ":" & Cells(CDbl(locR + 12), _
     CDbl(locC + 1)).Address & "*" & Cells(CDbl(locR + 3), CDbl(locC + 2)).Address & ":" & Cells(CDbl(locR + 12), CDbl(locC + 2)).Address & ") +" & (10 * WireHole) + BarbRoy
     If CDate(DueDate) = "12/12/9999" Then
     .Range(Cells(CDbl(locR), CDbl(locC + 3)).Address) = ""
     Else
     .Range(Cells(CDbl(locR), CDbl(locC + 3)).Address) = CDate(DueDate)
     End If
     .Range(Cells(CDbl(locR + 1), CDbl(locC + 3)).Address).Value = max & " Weeks"
     
     .Range(Cells(CDbl(locR), CDbl(locC + 5)).Address) = LeadEntry & " Weeks"
     
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
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC)).Address).Value = "Vendor"
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 1)).Address).Value = Vendor
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 2)).Address).Value = "Quantity Quoted"
     
     .Range(Cells(CDbl(locR + 2), CDbl(locC + 3)).Address).Value = MOQ
     
End With
GoTo EndProc
End If


For i = LBound(PartNames) To UBound(PartNames)

errNum = 3
    With Worksheets(newName)

      index = locR + 2 + i
     'Enter in component names in table
     .Range(Cells(CDbl(index), CDbl(locC)).Address).Value = PartNames(i)
     'Qty to build hose
     .Range(Cells(CDbl(index), CDbl(locC + 1)).Address).Value = compQTY(i)
     'Price info for Hose
     .Range(Cells(CDbl(index), CDbl(locC + 2)).Address).Value = "$" & Round(PriceList(i), 2)
     'On hand value for component
     .Range(Cells(CDbl(index), CDbl(locC + 3)).Address).Value = Round(onHandList(i), 2)
     'On order information for component
     .Range(Cells(CDbl(index), CDbl(locC + 4)).Address).Value = BacklogList(i)
     'Current orders parts claimed/needed
     .Range(Cells(CDbl(index), CDbl(locC + 5)).Address).Value = Round(ShortPartList(i), 2)
     'Difference to determine ifenough margin to build hose
     .Range(Cells(CDbl(index), CDbl(locC + 6)).Address).Value = Round((CDbl(BacklogList(i)) + CDbl(onHandList(i))) - CDbl(ShortPartList(i)), 2)
     'Time to get new components
     .Range(Cells(CDbl(index), CDbl(locC + 7)).Address).Value = LeadTimeList(i)
     
     
     'For Loop to enter in Price Break amounts
     For j = 1 To breakCount
     
     .Range(Cells(CDbl(locR), CDbl(locC)).Address).Offset(2 + i, 7 + j).Value = PriceBreaks(i, j)
     
     Next j
     
    End With
Next i

For i = 1 To breakCount
    With Worksheets(newName)
    index = locR + 3 + i
    .Range(Cells(CDbl(index), CDbl(locC + 9 + breakCount)).Address).Value = partQty(i)
    End With
Next i
     
errNum = 4
'Makes columns correct size
Worksheets(newName).Columns("A:R").AutoFit

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
