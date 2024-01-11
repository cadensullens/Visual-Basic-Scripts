Attribute VB_Name = "BuySellFillFunction"
Function Buy_Sell_Fill(hoseNames As String)

BuySellInfo.Valid.Caption = "Valid Until: " & Expire
BuySellInfo.Vendor.Caption = "Vendor: " & Vendor
BuySellInfo.Hosename.Caption = "Hose:" & " " & hoseNames
BuySellInfo.Price.Caption = "Price:" & " $" & PriceBS
BuySellInfo.Quoted.Caption = "Quote Date:" & " " & CDate(QuoteDate)
BuySellInfo.Leadtime.Caption = "Leadtime: " & LeadtimeBS & " Weeks"
BuySellInfo.MOQ.Caption = "Quantity Quoted: " & MOQ

End Function
