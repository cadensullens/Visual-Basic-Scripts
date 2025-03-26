Attribute VB_Name = "LookUpHose"
Public PartNames() As String
Public DueDate As String
Public partQty() As Double
Public Number As Double
Public breakCount As Double
Public hose As String
Public copyTemp As Double
Public compQTY() As Variant
Public Grand() As Double
Public PriceBreaks() As Double
Public LongLead() As Double
Public ShortPartList() As Double
Public BacklogList() As Double
Public onHandList() As Double
Public PriceList() As Double
Public LeadTimeList() As Double
Public max As Double
Public priceend As Double
Public iterate As Double
Public WireHole As Variant
Public Grandsum As Double
Public BarbRoy As Variant
Public SpecClean As String
Public OGBreak As Double
Public onIt As Double
Public FloatValue As Double
Public PriceWrong As Double
Public LeadEntry As String
Public MWrong As Double
Public CompGather As String
Public LiveLeadSkip As Boolean
Public OldPriceText As String
Public CleanCustomPrice As Double
Public PartInfoValue As Boolean
Public OGName As String



Public Sub Enter_Comp()

OGName = ActiveSheet.Name

Dim buyCount As Double
buyCount = 0
copyTemp = 0
OldPriceText = vbNullString

If onIt <> 1 Then
hose = ""
End If
HoseLookUp.Show vbModeless

onIt = 0
End Sub
