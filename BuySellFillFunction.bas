Attribute VB_Name = "BuySellFillFunction"
Function Buy_Sell_Fill(BuySellName As String)

Expire = CDate(CDbl(CDate(QuoteDate)) + ValidFor)

BuySellInfo.Valid.Caption = "Valid Until: " & Expire
BuySellInfo.Vendor.Caption = "Vendor: " & Vendor
BuySellInfo.HoseName.Caption = "Hose:" & " " & BuySellName
BuySellInfo.Price.Caption = "Price:" & " $" & PriceBS
BuySellInfo.Quoted.Caption = "Quote Date:" & " " & CDate(QuoteDate)
BuySellInfo.Leadtime.Caption = "Leadtime: " & LeadtimeBS & " Weeks"
BuySellInfo.MOQ.Caption = "Quantity Quoted: " & MOQ

End Function
