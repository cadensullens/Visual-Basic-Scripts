VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BuySellEntry 
   Caption         =   "Buy/Sell Infomation Entry"
   ClientHeight    =   3560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10584
   OleObjectBlob   =   "BuySellEntry.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BuySellEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ContinueActive_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SaveData
Unload BuySellEntry

Set wb = Workbooks.Open("https://futuremetals0.sharepoint.com/sites/Aero-HoseDCC/Shared Documents/From Sales/Quoting/Quote Sheet Files/BOMsForHoses.xlsx")
Sheets("Buy-Sell").Select
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
With Workbooks("BOMsForHoses.xlsx").Worksheets("Buy-Sell")
.Range(Cells(CDbl(lRow + 1), 1).Address).Value = hose
.Range(Cells(CDbl(lRow + 1), 2).Address).Value = Vendor
.Range(Cells(CDbl(lRow + 1), 3).Address).Value = PriceBS
.Range(Cells(CDbl(lRow + 1), 4).Address).Value = LeadtimeBS
.Range(Cells(CDbl(lRow + 1), 5).Address).Value = QuoteDate
.Range(Cells(CDbl(lRow + 1), 6).Address).Value = ValidFor
.Range(Cells(CDbl(lRow + 1), 7).Address).Value = MOQ
End With

    wb.Close True 'save changes
    ThisWorkbook.Connections("Query - Buy-Sell").Refresh
errNum = 7
Call Buy_Sell_Fill(hose)
BuySellInfo.Show
End Sub

Sub ContinueInactive_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Make Save Button Green when hovered on

ContinueInactive.Visible = False

End Sub

Private Sub LeadtimeEntry_Change()
If Not IsNumeric(LeadtimeEntry.Value) Then
LeadtimeEntry.Value = ""
End If
End Sub

Private Sub MOQEntry_Change()
If Not IsNumeric(MOQEntry.Value) Then
MOQEntry.Value = ""
End If
End Sub

Private Sub PriceEntry_Change()
If Not IsNumeric(PriceEntry.Value) Then
PriceEntry.Value = ""
End If
End Sub

Private Sub UserForm_Initialize()
HoseNameEntry.Value = hose
BuySellEntry.Caption = "Buy/Sell Information for " & hose
QuoteDateEntry.Value = Date
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'PURPOSE: Reset Userform buttons to Inactive Status

ContinueInactive.Visible = True

End Sub
Public Sub SaveData()

hose = HoseNameEntry.Value
Vendor = VendorEntry.Value
PriceBS = PriceEntry.Value
LeadtimeBS = LeadtimeEntry.Value
QuoteDate = QuoteDateEntry.Value
If ValidEntry.Value = "" Then
ValidFor = 0
Else
ValidFor = ValidEntry.Value
End If

MOQ = MOQEntry.Value
End Sub

Private Sub ValidEntry_Change()
If Not IsNumeric(ValidEntry.Value) Then
ValidEntry.Value = ""
End If
End Sub
