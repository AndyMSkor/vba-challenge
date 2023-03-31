Sub stocktickermoduletwo():
Dim stockticker As String
Dim stocktickertotal As Double
Dim summarytablerow As Integer
Dim rowcount As Long
Dim openingstockvalue As Double
Dim closingstockvalue As Double
Dim stockyearlychange As Double
Dim rownumber As Integer
Dim nextrow As Integer
Dim stockpercentagechange As Double
Dim greatesttotalvolume As Double


greatesttotalvolume = WorksheetFunction.Max(Range("L2:L3001").Value)


rowcount = Cells(Rows.Count, 1).End(xlUp).Row
stocktickertotal = 0
summarytablerow = 2
nextrow = 2

For e = 2 To rowcount

If Cells(e + 1, 1).Value <> Cells(e, 1).Value Then
stockticker = Cells(e, 1).Value
stocktickertotal = stocktickertotal + Cells(e, 7).Value
Range("I" & summarytablerow).Value = stockticker
Range("L" & summarytablerow).Value = stocktickertotal
summarytablerow = summarytablerow + 1
stocktickertotal = 0
Else
stocktickertotal = stocktickertotal + Cells(e, 7)
End If

Next e

For e = 2 To rowcount
If Cells(e + 1, 1).Value <> Cells(e, 1).Value Then
openingstockvalue = Cells(e, 3).Value
ElseIf Cells(e - 1, 1).Value <> Cells(e, 1).Value Then
closingstockvalue = Cells(e, 6).Value
Else
stockyearlychange = closingstockvalue - openingstockvalue
Cells(e + 1, "J").Value = stockyearlychange
End If

Next e

For e = 2 To rowcount

If Cells(e + 1, 1).Value <> Cells(e, 1).Value Then
stockpercentagechange = (closingstockvalue) - (openingstockvalue) / (openingstockvalue)
Range("k2:k3001").Value = stockpercentagechange
End If

Next e


For e = 2 To rowcount
If Cells(e, 12).Value = maxvalue Then

Range("N3").Value = Cells(e, 9).Value
End If



Next e





Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("N2").Value = "Greatest Total Volume"
Range("O2").Value = greatesttotalvolume


End Sub
