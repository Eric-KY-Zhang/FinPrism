Attribute VB_Name = "模块_抓台股指标表"
Option Explicit

' =================================================================
'  生成台股 18 项标准指标表
'  写入 Sheet: 台股_指标表
' =================================================================

Public Sub Main()
    BuildStandardIndicatorSheet "TW"
End Sub
