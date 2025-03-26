VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveQuote 
   Caption         =   "Quote Information"
   ClientHeight    =   4020
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14880
   OleObjectBlob   =   "SaveQuote.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FBox() As New QuoteBox
Private FpBox() As New QuoteBox
Private FbBox() As New QuoteBox
Private Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NumberCheck
If MWrong > 0 Then GoTo EndSub


'Set wb2 = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/Quoting/Quotes_Dashboard.xlsx")
Workbooks("Quotes_Dashboard.xlsx").Sheets("RFQ").Activate
'Finds the last non-blank cell in a single row or column

Dim lRow As Long

    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    

For i = 1 To NumbHose.Value
With Workbooks("Quotes_Dashboard.xlsx").Worksheets("RFQ")
.Range(Cells(CDbl(lRow + i), 1).Address).Value = QuoteDateQuote.Value
.Range(Cells(CDbl(lRow + i), 2).Address).Value = SupplierNames.Value
.Range(Cells(CDbl(lRow + i), 3).Address).Value = CustRFQ.Value
.Range(Cells(CDbl(lRow + i), 4).Address).Value = PMNames.Value
.Range(Cells(CDbl(lRow + i), 6).Address).Value = i
.Range(Cells(CDbl(lRow + i), 7).Address).Value = SaveQuote.Controls("HoseNameQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 8).Address).Value = SaveQuote.Controls("CustNameQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 9).Address).Value = SaveQuote.Controls("QtyQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 10).Address).Value = SaveQuote.Controls("SellPriceQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 11).Address).Value = SaveQuote.Controls("SellPriceQuote" & i).Value * SaveQuote.Controls("QtyQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 12).Address).Value = SaveQuote.Controls("HoseCostQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 13).Address).Value = SaveQuote.Controls("MarginQuote" & i).Value / 100
.Range(Cells(CDbl(lRow + i), 14).Address).Value = SaveQuote.Controls("ProductCombo" & i).Value
.Range(Cells(CDbl(lRow + i), 15).Address).Value = SalesRep.Value
.Range(Cells(CDbl(lRow + i), 16).Address).Value = ApplicationCombo.Value
.Range(Cells(CDbl(lRow + i), 17).Address).Value = PlatformDrop.Value
.Range(Cells(CDbl(lRow + i), 18).Address).Value = SaveQuote.Controls("LeadtimeQuote" & i).Value
.Range(Cells(CDbl(lRow + i), 19).Address).Value = SaveQuote.Controls("LTCombo" & i).Value
End With

Next i


    wb.Close True 'save changes
    Unload SaveQuote
EndSub:
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True

End Sub

Private Sub NumbHose_Change()

If Not IsNumeric(NumbHose.Value) Then
NumbHose.Value = ""
Else
If NumbHose.Value = 0 Then
NumbHose.Value = 1
Else
If NumbHose.Value < OGBreak Then
RemoveBoxes
Else
FloatValue = NumbHose.Value
ReDim Preserve FBox(NumbHose.Value)
ReDim Preserve FpBox(NumbHose.Value)
ReDim Preserve FbBox(NumbHose.Value)
Update
OGBreak = NumbHose.Value
End If
End If
End If
EndSub:
End Sub

Public Sub UserForm_Initialize()
'Default values
ReDim FBox(1)
ReDim FpBox(1)
ReDim FbBox(1)
OGBreak = 0
FloatValue = 1
QuoteDateQuote.Value = Date
NumbHose.Value = 1

Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/Quoting/Quotes_Dashboard.xlsx")
Sheets("Lists").Select
'Adds all company names to the list
Dim lRow As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Worksheets("Lists").Cells(Rows.Count, 5).End(xlUp).Row

SupplierNames.List = Workbooks("Quotes_Dashboard.xlsx").Worksheets("Lists").Range("E2:E" & lRow).Value


With PMNames
    .AddItem "Caden Sullens"
    .AddItem "Elsa Siino"
    .AddItem "Sam Gust"
    .AddItem "Morgan Williams"
End With

With PlatformDrop
    .AddItem "Aircraft"
    .AddItem "Energy"
    .AddItem "Ground Support"
    .AddItem "Ground Vehicle"
    .AddItem "Helicopter"
    .AddItem "Ship"
    .AddItem "Space"
    .AddItem "Other"
End With

With ApplicationCombo
    .AddItem "Commercial"
    .AddItem "Military"
    .AddItem "Space"
    .AddItem "Other"
End With

With SalesRep
    .AddItem "Allied International"
    .AddItem "Cardavio 5.5%"
    .AddItem "Cardavio 8%"
    .AddItem "Jet Star 1.5%"
    .AddItem "Jet Star 5%"
    .AddItem "Mfg Conn 1.5%"
    .AddItem "Mfg Conn 5%"
    .AddItem "Ramani 2.5%"
    .AddItem "Ramani 5%"
    .AddItem "Tom Carmody 2%%"
    .AddItem "Tony Varnell 1.5%"
    .AddItem "Tony Varnell 5%"
End With
    
End Sub

Public Sub Update()
Dim i As Double

If FloatValue > OGBreak Then
StartValue = OGBreak + 1
End If

For i = StartValue To NumbHose.Value


With SaveQuote.Controls.Add("Forms.TextBox.1", "HoseNameQuote" & i)
    .Top = 118 + (i) * 20
    .Left = 24
    .Width = 150
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With
    
With SaveQuote.Controls.Add("Forms.TextBox.1", "CustNameQuote" & i)
    .Top = 118 + (i) * 20
    .Left = 176
    .Width = 150
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With

With SaveQuote.Controls.Add("Forms.TextBox.1", "QtyQuote" & i)
    .Top = 118 + (i) * 20
    .Left = 328
    .Width = 48
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With
    
Set HCBox = SaveQuote.Controls.Add("Forms.TextBox.1", "HoseCostQuote" & i)
With HCBox
    .Top = 118 + (i) * 20
    .Left = 378
    .Width = 54
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With
Set FbBox(i).newFbBox = HCBox
    
Set MarginBox = SaveQuote.Controls.Add("Forms.TextBox.1", "MarginQuote" & i)
With MarginBox
    .Top = 118 + (i) * 20
    .Left = 434
    .Width = 54
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .Value = 1
    End With
Set FBox(i).newFBox = MarginBox

Set SpBox = SaveQuote.Controls.Add("Forms.TextBox.1", "SellPriceQuote" & i)
    With SpBox
    .Top = 118 + (i) * 20
    .Left = 490
    .Width = 60
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With
Set FpBox(i).newFpBox = SpBox

With SaveQuote.Controls.Add("Forms.TextBox.1", "LeadtimeQuote" & i)
    .Top = 118 + (i) * 20
    .Left = 552
    .Width = 52
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    End With
    
With SaveQuote.Controls.Add("Forms.ComboBox.1", "LTCombo" & i)
    .Top = 118 + (i) * 20
    .Left = 606
    .Width = 60
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .ShowDropButtonWhen = 1
    .AddItem "Weeks"
    .AddItem "Days"
    .AddItem "In Stock"
    End With
    
With SaveQuote.Controls.Add("Forms.ComboBox.1", "ProductCombo" & i)
    .Top = 118 + (i) * 20
    .Left = 668
    .Width = 66
    .Height = 18
    .Font.Name = "Calibri"
    .Font.Size = 10
    .ForeColor = &H80000008
    .BorderStyle = 1
    .BorderColor = &HA9A9A9
    .SpecialEffect = 0
    .ShowDropButtonWhen = 1
    .AddItem "Maker"
    .AddItem "Bulk Hose"
    .AddItem "Buy/Sell"
    End With
    
Next i

SaveQuote.Height = (NumbHose.Value) * 18 + 229 + (2 * (NumbHose.Value))
ContinueInactive.Top = (NumbHose.Value) * 18 + 150 + (2 * (NumbHose.Value))
ContinueActive.Top = (NumbHose.Value) * 18 + 150 + (2 * (NumbHose.Value))
Skip:
End Sub
Sub RemoveBoxes()

StartV = CDbl(NumbHose.Value) + 1
For i = StartV To OGBreak
    SaveQuote.Controls.Remove ("LTCombo" & i)
    SaveQuote.Controls.Remove ("ProductCombo" & i)
    SaveQuote.Controls.Remove ("LeadtimeQuote" & i)
    SaveQuote.Controls.Remove ("SellPriceQuote" & i)
    SaveQuote.Controls.Remove ("MarginQuote" & i)
    SaveQuote.Controls.Remove ("HoseCostQuote" & i)
    SaveQuote.Controls.Remove ("QtyQuote" & i)
    SaveQuote.Controls.Remove ("HoseNameQuote" & i)
    SaveQuote.Controls.Remove ("CustNameQuote" & i)
    Next i

ReDim PreserveFBox(NumbHose.Value)
ReDim PreserveFpBox(NumbHose.Value)
ReDim PreserveFbBox(NumbHose.Value)

OGBreak = NumbHose.Value
SaveQuote.Height = (NumbHose.Value) * 18 + 229 + (2 * (NumbHose.Value))
ContinueInactive.Top = (NumbHose.Value) * 18 + 150 + (2 * (NumbHose.Value))
ContinueActive.Top = (NumbHose.Value) * 18 + 150 + (2 * (NumbHose.Value))


End Sub

Sub NumberCheck()
MWrong = 0
For i = 1 To NumbHose.Value
If Not IsNumeric(SaveQuote.Controls("MarginQuote" & i).Value) Then
SaveQuote.Controls("MarginQuote" & i).Value = ""
MWrong = MWrong + 1
End If
If Not IsNumeric(SaveQuote.Controls("QtyQuote" & i).Value) Then
SaveQuote.Controls("QtyQuote" & i).Value = ""
MWrong = MWrong + 1
End If
If Not IsNumeric(SaveQuote.Controls("HoseCostQuote" & i).Value) Then
SaveQuote.Controls("QtyQuote" & i).Value = ""
MWrong = MWrong + 1
End If
If Not IsNumeric(SaveQuote.Controls("LeadtimeQuote" & i).Value) Then
SaveQuote.Controls("QtyQuote" & i).Value = ""
MWrong = MWrong + 1
End If
Next i
If MWrong > 0 Then MsgBox ("Some entries have been erased. Entries can only be numbers.")

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = 0 Then
wb.Close False
End If

End Sub

