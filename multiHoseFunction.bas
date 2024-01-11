Attribute VB_Name = "multiHoseFunction"
Public CopyCheck As Double

Function multiHose(hose() As String)


Dim loc As Double

On Error GoTo Errhandler
For k = LBound(hose) To UBound(hose)
'variable to go gather the information on the hose builds
errNum = 1
CopyCheck = 1
Call Gather_Info(hose(k))
If Gathererr = 1 And k <> NumberHose Then
Gathererr = 0
GoTo pass
ElseIf Gathererr = 0 Then GoTo skipM
Else
GoTo EndProc
End If


skipM:
loc = k * 14 - 10

If BuySell <> 1 Then
Call BOM_CopyTable(loc, 1, newName)

With ActiveSheet
     'Inserts Info for each new hose in blocks above table
     .Range("B" & loc).Value = hose(k)
     .Range("B" & loc + 1).Formula2 = "=SUM(B" & loc + 3 & ":B" & loc + 12 & "*C" & loc + 3 & ":C" & loc + 12 & ") + " & (10 * WireHole) + BarbRoy
     If CDate(DueDate) = "12/12/9999" Then
     .Range("D" & loc) = ""
     Else
     .Range("D" & loc) = CDate(DueDate)
     End If
     .Range("D" & loc + 1).Value = max & " Weeks"
     
     .Range("F" & loc) = LeadEntry & " Weeks"
     
End With
Else

'Goes to Function that copyies and formats information for buy sells
    Call BuySell_CopyTable(loc, 1, newName)

    With ActiveSheet

     'Inserts Info for each new hose in blocks above table
     .Range("B" & loc).Value = hose(k)
     
     .Range("B" & loc + 1).Value = PriceBS
     
     .Range("C" & loc).Value = "Quote Date"
     
     .Range("D" & loc) = CDate(QuoteDate)
     
     .Range("A" & loc + 3).Value = "Max LeadTime"
     
     .Range("B" & loc + 3).Value = LeadtimeBS
     
     .Range("C" & loc + 1).Value = "Valid Until:"
     
     .Range("D" & loc + 1).Value = Expire
     
     .Range("A" & loc + 2).Value = "Vendor"
     
     .Range("B" & loc + 2).Value = Vendor
     
     .Range("C" & loc + 2).Value = "Quantity Quoted"
     
     .Range("D" & loc + 2).Value = MOQ
     
     End With
GoTo pass
End If


    errNum = 2
    'Fills in Formatted table with information form Gather Info one block
    For i = LBound(PartNames) To UBound(PartNames)
    
    errNum = 3
    With ActiveSheet

    index = loc + 2 + i
    'Enter in component names in table
    .Range("A" & index).Value = PartNames(i)
    'Qty to build hose
    .Range("B" & index).Value = compQTY(i)
    'Price info for Hose
    .Range("C" & index).Value = "$" & Round(PriceList(i), 2)
    'On hand value for component
    .Range("D" & index).Value = onHandList(i)
    'On order information for component
    .Range("E" & index).Value = BacklogList(i)
    'Current orders parts claimed/needed
    .Range("F" & index).Value = Round(ShortPartList(i), 2)
    'Difference to determine ifenough margin to build hose
    .Range("G" & index).Value = (CDbl(BacklogList(i)) + CDbl(onHandList(i))) - CDbl(ShortPartList(i))
    'Time to get new components
    .Range("H" & index).Value = LeadTimeList(i)
     
    errNum = 4
        'For Loop to enter in Price Break amounts
        For j = 1 To breakCount
        
        .Range("A" & loc).Offset(2 + i, 7 + j).Value = PriceBreaks(i, j)
         
        Next j
     
    End With

    Next i


pass:
Next k

errNum = 5
'Makes columns correct size
Worksheets(newName).Columns("A:R").AutoFit
'Ends the copy mode of the selected cells

GoTo EndProc

Errhandler:
If errNum = 1 Then
MsgBox (CStr(errNum) & " multiHose")
GoTo EndProc
End If

If errNum = 2 Then
MsgBox (CStr(errNum) & " multiHose")
GoTo EndProc
End If

If errNum = 3 Then
MsgBox (CStr(errNum) & " multiHose")
GoTo EndProc
End If

If errNum = 4 Then
MsgBox (CStr(errNum) & " multiHose")
GoTo EndProc
End If

If errNum = 5 Then
MsgBox (CStr(errNum) & " multiHose")
GoTo EndProc
End If

EndProc:
End Function

