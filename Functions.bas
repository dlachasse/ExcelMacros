'''''''''''''''''''''''''''''''''''
'''''''''''' FUNCTIONS ''''''''''''
'''''''''''''''''''''''''''''''''''


' TO BUILD
'
' Public Function SizeColor()
'   FormulaR1C1 = "=IF(RC[-3]=1,""Single"",IF(RC[-3]=RC[-2],""Color"",IF(RC[-3]=RC[-1],""Size"",""SizeColor"")))"
' End Function
'

Public Function SKUBuild(VendorPrefix As String, Style As String, Color As String, Size As String, VariationType As String) As String

    Dim x As String

    SKUBuild = VendorPrefix + Style
    
    Select Case VariationType
        Case "SizeColor"
            SKUBuild = SKUBuild & "-" + Left(Color, 21) + "-" + Size
        Case "Size"
            SKUBuild = SKUBuild & "-" + Size
        Case "Color"
            SKUBuild = SKUBuild & "-" + Color
        Case "Single"
            SKUBuild = SKUBuild
    End Select

    SKUBuild = Replace(SKUBuild, " ", "")

End Function

Public Function NetProfit(Store As String, Weight As Double, Cost As Double, Price As Double) As Double

    Dim ShipCharge, ShipCost As Double
    Dim ShipBand, cell As Range
    Weight = Round(Weight * 16, 0)

    With Sheets("ProfitCalc")
        ShipCost = Range("A:A").Find(Weight).Offset(0, 1).Value

        Select Case Store
            Case "Blank"
                Set ShipBand = Range("D1:F18")
            Case "Hive"
                Set ShipBand = Range("H1:J14")
            Case "Combi"
                Set ShipBand = Range("L1:N4")
            Case Else
                Set ShipBand = Range("H1:J14")
        End Select
    
        For Each cell In ShipBand
            If Between(Price, cell.Value, cell.Offset(0, 1).Value) = True Then
                ShipCharge = cell.Offset(0, 2).Value
            End If
        Next cell
            
    End With

    NetProfit = Round((((Price + ShipCharge) * 0.85) - ShipCost) - Cost, 2)

End Function

Public Function Between(test As Variant, base As Variant, ceil As Variant) As Boolean

    If test >= base And test <= ceil Then
        Between = True
    Else
        Between = False
    End If

End Function

Public Function NetProfitBreakdown(Store As String, Weight As Double, Cost As Double, Price As Double, Calc As Integer) As Double

    ' 1 Price + Shipping
    ' 2 Amazon commission
    ' 3 Average shipping cost
    ' 4 Profit

    Dim ShipCharge, ShipCost As Double
    Dim ShipBand, cell As Range
    Weight = Round(Weight * 16, 0)
    Overhead = 0.2 ' Arbitrary number for shipment cost to us
    
    Windows("PERSONAL.XLSB").Activate

    With Sheets("ProfitCalc")
        ShipCost = Range("A:A").Find(Weight).Offset(0, 1).Value

        Select Case Store
            Case "Blank"
                Set ShipBand = Range("D1:F18")
            Case "Hive"
                Set ShipBand = Range("H1:J14")
            Case "Combi"
                Set ShipBand = Range("L1:N4")
            Case Else
                Set ShipBand = Range("H1:J14")
        End Select
    
        For Each cell In ShipBand
            If Between(Price, cell.Value, cell.Offset(0, 1).Value) = True Then
                ShipCharge = cell.Offset(0, 2).Value
            End If
        Next cell
            
    End With

    Select Case Calc
        Case 1 ' Price + Shipping
            NetProfitBreakdown = Round((Price + ShipCharge), 2)
        Case 2 ' Amazon commission
            NetProfitBreakdown = Round((Price + ShipCharge) * 0.85, 2)
        Case 3 ' Average shipping cost
            NetProfitBreakdown = Round(ShipCost, 2)
        Case 4 ' Profit
            NetProfitBreakdown = Round(((((Price + ShipCharge) * 0.85) - ShipCost) - Cost) - Overhead, 2)
    End Select

End Function

Public Function ArrayAdd(txt As String) As Long
    Dim arr As Variant
    Dim x As Integer
    Dim ct As Long
    
    arr = Split(txt, ",")
    
    For x = LBound(arr) To UBound(arr)
    
        ct = ct + arr(x)
         
    Next x

    ArrayAdd = ct

End Function

Public Function UPC(str As String) As Long

    str = Right(String(12, "0") & strOrdNo, 12)

End Function

Function StripHTML(cell As Range) As String
 Dim RegEx As Object
 Set RegEx = CreateObject("vbscript.regexp")

 Dim sInput As String
 Dim sOut As String
 sInput = cell.Text

 With RegEx
   .Global = True
   .IgnoreCase = True
   .MultiLine = True
.Pattern = "<[^>]+>" 'Regular Expression for HTML Tags.
 End With

 sOut = RegEx.Replace(sInput, "")
 StripHTML = sOut
 Set RegEx = Nothing
End Function


