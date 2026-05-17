Attribute VB_Name = "模块_抓台股利润表"
Option Explicit

' =================================================================
'  抓台股 Income Statement — FinMind
'  写入 Sheet: 台股_利润表
' =================================================================

Public Sub Main()
    RunTWStatement "Income", "台股_利润表", GetTWISConcepts(), 8
End Sub


Private Function GetTWISConcepts() As Variant
    Dim a(0 To 9) As Variant
    a(0) = Array("Revenue", "Revenue")
    a(1) = Array("Revenue", "Cost of goods & services sold")
    a(2) = Array("Profitability", "Gross profit")
    a(3) = Array("Operating Expenses", "Total operating expenses")
    a(4) = Array("Profitability", "Operating income")
    a(5) = Array("Profitability", "Pre-tax income")
    a(6) = Array("Tax", "Income tax expense")
    a(7) = Array("Profitability", "Net income")
    a(8) = Array("EPS", "Basic EPS")
    a(9) = Array("EPS", "Diluted EPS")
    GetTWISConcepts = a
End Function
