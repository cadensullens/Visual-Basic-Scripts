Attribute VB_Name = "MakerStatusFunction"

Public SONumber() As String
Public BuildQty() As Double
Public CustDate() As String
Public CompDate() As String
Public MakerTrue As Double
Public JobStat() As String
Public wb As Workbook

Function MakerStatus(PO As String)

Dim Add As Double
Dim table As ListObject
On Error GoTo Errhandler
'Opens BOM workbook
Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Planning/Maker Work Order Tracker 2023.xlsm")

Set table = Workbooks("Maker Work Order Tracker 2023.xlsm").Worksheets("WorkOrders").ListObjects("WorkOrders")
Sheets("WorkOrders").Select

For i = 1 To Len(PO)
Character = Mid(PO, i, 1)
If Character Like "[a-zA-Z]" Then GoTo StringCheck
Next i

'If the for loop does not find any letters then it skips to evaluating for the double
GoTo DoubleCheck


StringCheck:
With ActiveSheet
    CheckPO = .Evaluate(table.ListColumns(4).DataBodyRange.Address & "=""" & PO & """")
End With
GoTo BoolCheck

DoubleCheck:
With ActiveSheet
    CheckPO = .Evaluate(table.ListColumns(4).DataBodyRange.Address & "=" & CDbl(PO) & "")
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
            With Workbooks("Maker Work Order Tracker 2023.xlsm").Worksheets("WorkOrders")
                ReDim Preserve SONumber(1 To Add)
                    SONumber(Add) = .Range(Cells(k + 2, 6).Address).Value
                ReDim Preserve BuildQty(1 To Add)
                    BuildQty(Add) = .Range(Cells(k + 2, 9).Address).Value
                ReDim Preserve CustDate(1 To Add)
                    CustDate(Add) = .Range(Cells(k + 2, 18).Address).Value
                ReDim Preserve CompDate(1 To Add)
                    CompDate(Add) = .Range(Cells(k + 2, 22).Address).Value
                ReDim Preserve JobStat(1 To Add)
                    JobStat(Add) = .Range(Cells(k + 2, 3).Address).Value
            End With
        End If
        Next k
        
                            
        StatusUpdate.Show
    End If
'Reset Copy_AnotherSheet
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
