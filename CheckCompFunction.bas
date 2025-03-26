Attribute VB_Name = "CheckCompFunction"
Public PartErr As Double
Public CompNumb As Double

Function Check_Comp()
On Error GoTo Errhandler

PartErr = 0
Dim table As ListObject
Dim ws As Worksheet
Dim table1 As ListObject
Dim ws1 As Worksheet

Set ws = Worksheets("Qb inventory")
Set table = ws.ListObjects("Inventory")

Set ws1 = Worksheets("BOM Master")
Set table1 = ws1.ListObjects("BOMMaster")

Build = ws1.Evaluate(table1.ListColumns(1).DataBodyRange.Address & "=""" & hose & """")
     
     Dim Bool1() As Double
     Dim check1 As Double
     For j = LBound(Build) To UBound(Build)
        If Build(j, 1) = False Then
        check1 = 0
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        Else
        check1 = 1
        ReDim Preserve Bool1(1 To j)
        Bool1(j) = check1
        End If
        Next j
        
  errNum = 3
'Check for hose on BOM
If Application.WorksheetFunction.Sum(Bool1) > 0 Then GoTo Errhandler


errNum = 1
CompNumb = Application.InputBox( _
Title:="Components Count", _
Prompt:="How many Components are you entering for " & hose & " ?", _
Type:=1)
If CompNumb = False Then GoTo EndSub

errNum = 2
For i = 1 To CompNumb
part = Application.InputBox( _
Title:="Component Name" & " " & i, _
Prompt:="What is Component " & i & "'s name for " & hose & " ?", _
Type:=1 + 2)


If part = "0" Or part = "False" Then GoTo EndSub

errNum = 1
'If part has OPINV: included it will skip adding it
If Left(part, 6) <> "OPINV:" Then
Partqb = "OPINV:" & part
Else
Partqb = part
End If

PartCheck = ws.Evaluate(table.ListColumns(1).DataBodyRange.Address & "=""" & Partqb & """")
     
     Dim Bool() As Double
     Dim check As Double
     For j = LBound(PartCheck) To UBound(PartCheck)
        If PartCheck(j, 1) = False Then
        check = 0
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        Else
        check = 1
        ReDim Preserve Bool(1 To j)
        Bool(j) = check
        End If
        Next j
        
  errNum = 4
If Application.WorksheetFunction.Sum(Bool) = 0 Then GoTo Errhandler

Count = Count + 1
ReDim Preserve PartNames(1 To Count)
PartNames(Count) = Partqb

Start:
Next i
GoTo EndSub
Errhandler:


If errNum = 1 Then
MsgBox (errNum & " CheckComp")
GoTo EndSub
End If

If errNum = 2 Then
MsgBox (errNum & " CheckComp")
GoTo EndSub
End If

If errNum = 3 Then
MsgBox ("Hose is already on the Bom Master sheet, Use Look up part function to get Hose information.")
PartErr = 1
GoTo EndSub
End If

If errNum = 4 Then
If CompNumb > 1 Then
MsgBox ("Part, " & part & ", not found on QB Inventory List. Please check spelling of Component Name.")
CompNumb = CompNumb - 1
GoTo Start
Else
MsgBox ("Part, " & part & ", not found on QB Inventory List. Please check spelling of Component Name.")
PartErr = 1
GoTo EndSub
End If
End If


EndSub:
End Function
