Attribute VB_Name = "DateEntryFunction"
Public LeadEntry As Variant

Function DateEntry()
On Error GoTo Errhandler

Start1:
errNum = 1
LeadEntry = Application.InputBox( _
Title:="Lead Time", _
Prompt:="Enter the Target Lead time in Weeks or Type 'All' for no Date Filter", _
Type:=1 + 2)

If VarType(LeadEntry) = 11 Then GoTo EndSub
If LCase(LeadEntry) = "all" Then
DueDate = "12/12/9999"
LeadEntry = ""
GoTo Skip
End If

ConvertDays = LeadEntry * 7
DueDate = Date + ConvertDays

Skip:
'Does not allow dates in the past for entry
If CDate(DueDate) > Date Then
ElseIf CDate(DueDate) = Date Then
GoTo Errhandler
End If

GoTo EndSub

Errhandler:
If errNum = 1 Then
MsgBox ("Date entered Incorrectly or Date is in the past.")
GoTo Start1
End If

EndSub:
End Function
