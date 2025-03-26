Attribute VB_Name = "BacklogStatusFunction"


Function BackLogCheck(Po As String)

Dim Add As Double
Dim table As ListObject
On Error GoTo Errhandler
Set wb = Workbooks.Open("Your Sharepoint URL")

Set table = Workbooks("Backlog Report.xlsb").Worksheets("Backlog").ListObjects("Table1")


For i = 1 To Len(Po)
Character = Mid(Po, i, 1)
If Character Like "[a-zA-Z]" Then GoTo StringCheck
Next i

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheck

StringCheck:
With ActiveSheet
    CheckPO = .Evaluate(table.ListColumns(6).DataBodyRange.Address & "=""" & Po & """")
End With
GoTo BoolCheck

DoubleCheck:
With ActiveSheet
    CheckPO = .Evaluate(table.ListColumns(6).DataBodyRange.Address & "=" & CDbl(Po) & "")
End With

BoolCheck:
     Dim Bool() As Double
     Dim check As Double
    For j = LBound(CheckPO) To UBound(CheckPO)
        If CheckPO(j, 1) = False Then
            check = 0
        ReDim Preserve Bool(1 To j)
            Bool(j) = check
        Else
            check = 1
            ReDim Preserve Bool(1 To j)
            Bool(j) = check
        End If
     Next j

    If Application.WorksheetFunction.Sum(Bool) = 0 Then
        GoTo Errhandler
    Else
        POHits = Application.WorksheetFunction.Sum(Bool)
        
        For k = 1 To UBound(Bool)
            If Bool(k) = 1 Then
            Add = Add + 1
                With Workbooks("Backlog Report.xlsb").Worksheets("Backlog")
                    ReDim Preserve SONumber(1 To Add)
                        SONumber(Add) = .Range(Cells(k + 1, 5).Address).Value
                    ReDim Preserve BuildQty(1 To Add)
                        BuildQty(Add) = .Range(Cells(k + 1, 9).Address).Value
                    ReDim Preserve CustDate(1 To Add)
                        CustDate(Add) = .Range(Cells(k + 1, 2).Address).Value
                    ReDim Preserve CompDate(1 To Add)
                        CompDate(Add) = .Range(Cells(k + 1, 3).Address).Value
                    ReDim Preserve JobStat(1 To Add)
                        JobStat(Add) = .Range(Cells(k + 1, 1).Address).Value
                End With
                
            End If
            Next k
            
        StatusUpdate.Show
    End If
If copyTemp <> 0 Then
copyTemp = 0
GoTo EndSub
End If

wb.Close False 'does not save changes
GoTo EndSub

Errhandler:

wb.Close False 'does not save changes

'If errNum = 0 Then
'
'GoTo EndSub
'End If

EndSub:
End Function
