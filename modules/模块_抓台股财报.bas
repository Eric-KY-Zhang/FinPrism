Attribute VB_Name = "模块_抓台股财报"
Option Explicit

' =================================================================
'  抓台股财报 — 共享 helpers (RunTWStatement / FetchTWFromFinMind)
'
'  数据源: FinMind public API
'    https://api.finmindtrade.com/api/v4/data
'
'  约定:
'    - 第一版台股用 FinMind 三表 JSON 做主链路, 不引入登录/API Key。
'    - 金额写入值 = FinMind 原值 / 1,000,000 (百万 TWD); EPS 保持原值。
'    - 官方 TWSE/TPEx/MOPS Fin 后续可作为校验源和同行比较增强, 不阻塞三表抓数。
' =================================================================


Public Sub RunTWStatement(ByVal strKind As String, ByVal targetSheet As String, _
                          ByVal conceptMap As Variant, _
                          ByVal maxPeriods As Long)
    Dim wsPool As Worksheet, wsTarget As Worksheet
    Dim arrPool As Variant
    Dim i As Long, lngRow As Long, numCompanies As Long
    Dim intFailCnt As Long, strErrLog As String
    Dim strCodeRaw As String, strCode As String, strName As String
    Dim dtTime As Double: dtTime = Timer

    Dim dictData As Object: Set dictData = CreateObject("Scripting.Dictionary")
    Dim dictPeriodSet As Object: Set dictPeriodSet = CreateObject("Scripting.Dictionary")
    Dim dictIndicatorSet As Object: Set dictIndicatorSet = CreateObject("Scripting.Dictionary")
    Dim dictCategoryMap As Object: Set dictCategoryMap = CreateObject("Scripting.Dictionary")
    Dim dictCompanyName As Object: Set dictCompanyName = CreateObject("Scripting.Dictionary")
    Dim dictReportingCurrency As Object: Set dictReportingCurrency = CreateObject("Scripting.Dictionary")
    Dim collCodes As New Collection
    Dim collDiagRows As New Collection

    g_diagnosticSheetName = "台股_抓取诊断"

    Set wsPool = ThisWorkbook.Sheets("样本池")
    Set wsTarget = ThisWorkbook.Sheets(targetSheet)

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    lngRow = wsPool.Cells(wsPool.Rows.Count, POOL_TW_CODE_COL).End(xlUp).Row
    If lngRow < POOL_DATA_START_ROW Then
        intFailCnt = 1
        strErrLog = "样本池台股区为空"
        GoTo CleanUp
    End If

    arrPool = wsPool.Range(wsPool.Cells(POOL_DATA_START_ROW, POOL_TW_CODE_COL), _
                           wsPool.Cells(lngRow, POOL_TW_NAME_COL)).Value
    If Not IsArray(arrPool) Then
        Dim singleVal As Variant: singleVal = arrPool
        ReDim arrPool(1 To 1, 1 To 2)
        arrPool(1, 1) = singleVal
    End If
    numCompanies = UBound(arrPool, 1)

    Dim strQuarter As String: strQuarter = ReadQuarterSelection()
    Dim lngYear As Long: lngYear = ReadYearSelection()

    For i = 1 To numCompanies
        strCodeRaw = Trim$(CStr(arrPool(i, 1)))
        If Len(strCodeRaw) = 0 Then GoTo NextRow
        strName = Trim$(CStr(arrPool(i, 2)))
        strCode = NormalizeTWTicker(strCodeRaw)

        Application.StatusBar = "抓取中: " & targetSheet & " (" & i & "/" & numCompanies & ") " & strCode
        DoEvents

        Dim tempData As Object: Set tempData = CreateObject("Scripting.Dictionary")
        Dim tempPeriodSet As Object: Set tempPeriodSet = CreateObject("Scripting.Dictionary")
        Dim tempIndicatorSet As Object: Set tempIndicatorSet = CreateObject("Scripting.Dictionary")
        Dim tempCategoryMap As Object: Set tempCategoryMap = CreateObject("Scripting.Dictionary")

        On Error Resume Next
        Err.Clear
        FetchTWFromFinMind strCode, strKind, conceptMap, strQuarter, lngYear, maxPeriods, _
                           tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, collDiagRows
        Dim fetchErrNum As Long: fetchErrNum = Err.Number
        Dim fetchErrSource As String: fetchErrSource = Err.Source
        Dim fetchErrDesc As String: fetchErrDesc = Err.Description
        Err.Clear
        On Error GoTo CleanUp

        If fetchErrNum <> 0 Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": " & _
                        "错误号=" & fetchErrNum & "; 来源=" & fetchErrSource & "; " & fetchErrDesc
            AddMissingDiagnosticsForCompany strCode, strKind, conceptMap, collDiagRows, _
                                            "fetch_failed: " & fetchErrDesc
        ElseIf Not HasAnyCoreLabelTW(tempData, strCode, CoreLabelsForKind(strKind)) Then
            intFailCnt = intFailCnt + 1
            strErrLog = strErrLog & vbCrLf & strCode & " " & strName & ": 核心字段未命中"
        Else
            CommitTempTWData strCode, tempData, tempPeriodSet, tempIndicatorSet, tempCategoryMap, _
                             dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
            If Not dictCompanyName.Exists(strCode) Then dictCompanyName.Add strCode, strName
            dictReportingCurrency(strCode) = "TWD"
            collCodes.Add strCode
        End If

        Application.Wait Now + TimeSerial(0, 0, 1)

NextRow:
    Next i

    If collCodes.Count = 0 Then
        ClearTWWideTableOutput wsTarget
        GoTo CleanUp
    End If

    If dictPeriodSet.Count = 0 Or dictIndicatorSet.Count = 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[" & targetSheet & "] 无匹配数据" & _
                    "(期数=" & dictPeriodSet.Count & " / 指标=" & dictIndicatorSet.Count & ")"
        GoTo CleanUp
    End If

    Dim arrCodes() As String
    ReDim arrCodes(1 To collCodes.Count)
    For i = 1 To collCodes.Count
        arrCodes(i) = collCodes.Item(i)
    Next i

    Dim arrPeriods As Variant, arrIndicators As Variant
    arrPeriods = SortPeriodsDesc(dictPeriodSet)
    If IsArray(arrPeriods) Then
        On Error Resume Next
        Dim ubP As Long: ubP = -1
        ubP = UBound(arrPeriods)
        On Error GoTo CleanUp
        If lngYear = 0 And ubP - LBound(arrPeriods) + 1 > maxPeriods Then
            ReDim Preserve arrPeriods(LBound(arrPeriods) To LBound(arrPeriods) + maxPeriods - 1)
        End If
    End If

    arrIndicators = OrderIndicatorsByTWConceptMap(dictIndicatorSet, conceptMap)

    Dim hookKind As String
    Select Case True
        Case InStr(targetSheet, "资产负债") > 0: hookKind = "BalanceSheet"
        Case InStr(targetSheet, "利润") > 0:     hookKind = "Income"
        Case InStr(targetSheet, "现金流") > 0:   hookKind = "CashFlow"
        Case Else:                               hookKind = ""
    End Select

    Application.StatusBar = "写入: " & targetSheet
    DoEvents
    WriteWideTable wsTarget, arrCodes, dictCompanyName, dictData, _
                   arrPeriods, arrIndicators, dictCategoryMap, _
                   perCompanyPeriods:=True, _
                   dictReportingCurrency:=dictReportingCurrency, _
                   statementKind:=hookKind
    RefreshA1CurrencyComment wsTarget, targetSheet

CleanUp:
    Dim cleanErrNum As Long: cleanErrNum = Err.Number
    Dim cleanErrSource As String: cleanErrSource = Err.Source
    Dim cleanErrDesc As String: cleanErrDesc = Err.Description
    If cleanErrNum <> 0 Then
        intFailCnt = intFailCnt + 1
        strErrLog = strErrLog & vbCrLf & "[RunTWStatement] 错误号=" & cleanErrNum & _
                    "; 来源=" & cleanErrSource & "; " & cleanErrDesc
        Err.Clear
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    On Error Resume Next
    WriteDiagnosticForKind strKind, collDiagRows
    Err.Clear
    On Error GoTo 0

    g_globalFails = g_globalFails + intFailCnt
    If Len(strErrLog) > 0 Then _
        g_globalLog = g_globalLog & vbCrLf & "[" & targetSheet & "]" & strErrLog

    If Not g_silentMode Then
        Dim msg As String
        msg = targetSheet & " 抓取完成 (单位: 百万 TWD)" & vbCrLf & _
              "用时: " & Format(Timer - dtTime, "0.0 秒") & vbCrLf & _
              "公司数: " & collCodes.Count & " / 期数: " & dictPeriodSet.Count
        If intFailCnt > 0 Then msg = msg & vbCrLf & vbCrLf & "失败 " & intFailCnt & " 条:" & strErrLog

        Dim style As Long: style = vbInformation
        If intFailCnt > 0 Then style = vbExclamation
        MsgBox msg, style, "上市公司财务数据查询 (TW)"
    End If
End Sub


Private Sub FetchTWFromFinMind(ByVal strTicker As String, _
                               ByVal strKind As String, _
                               ByVal conceptMap As Variant, _
                               ByVal strQuarter As String, _
                               ByVal lngYear As Long, _
                               ByVal maxPeriods As Long, _
                               ByRef dictData As Object, _
                               ByRef dictPeriodSet As Object, _
                               ByRef dictIndicatorSet As Object, _
                               ByRef dictCategoryMap As Object, _
                               Optional ByVal collDiagRows As Collection = Nothing)
    Dim stage As String: stage = "init"
    On Error GoTo TwErr

    Dim mapFM As Object: Set mapFM = FinMindFieldMapForKindTW(strKind)
    Dim dictRecords As Object: Set dictRecords = CreateObject("Scripting.Dictionary")
    dictRecords.CompareMode = vbTextCompare
    Dim dictOrigin As Object: Set dictOrigin = CreateObject("Scripting.Dictionary")
    dictOrigin.CompareMode = vbTextCompare

    stage = "FetchJson"
    LoadTWFinMindRows strTicker, strKind, strQuarter, lngYear, dictRecords, dictOrigin
    If dictRecords.Count = 0 Then
        Err.Raise vbObjectError + 842, "FetchTWFromFinMind", _
            "FinMind TW 无匹配期间: " & strTicker & " / " & strQuarter & " / " & lngYear
    End If

    stage = "BuildCompanyDict"
    Dim dictCompany As Object
    If dictData.Exists(strTicker) Then
        Set dictCompany = dictData.Item(strTicker)
    Else
        Set dictCompany = CreateObject("Scripting.Dictionary")
        dictData.Add strTicker, dictCompany
    End If

    Dim dictMatchInfo As Object: Set dictMatchInfo = CreateObject("Scripting.Dictionary")
    dictMatchInfo.CompareMode = vbTextCompare

    Dim periodKey As Variant, dictTypes As Object, dictPer As Object
    Dim ci As Long, mapEntry As Variant, strCat As String, strLabel As String
    Dim candidates As Variant, cand As Variant, candIdx As Long, totalCand As Long
    Dim rawVal As Variant, scaledVal As Double, originName As String

    For Each periodKey In dictRecords.Keys
        ' Quarter filter moved here (was inside LoadTWFinMindRows). We need every
        ' loaded quarter visible to the cumulator above, then we drop ones the
        ' user did not ask for before writing.
        If Not MatchTWPeriod(CStr(periodKey), strQuarter, lngYear, strKind) Then GoTo NextPeriodKey

        Set dictTypes = dictRecords.Item(periodKey)
        If dictCompany.Exists(CStr(periodKey)) Then
            Set dictPer = dictCompany.Item(CStr(periodKey))
        Else
            Set dictPer = CreateObject("Scripting.Dictionary")
            dictCompany.Add CStr(periodKey), dictPer
        End If

        For ci = LBound(conceptMap) To UBound(conceptMap)
            mapEntry = conceptMap(ci)
            strCat = MapEntryCategory(mapEntry)
            strLabel = MapEntryLabel(mapEntry)
            If Not mapFM.Exists(strLabel) Then GoTo NextConcept

            candidates = Split(CStr(mapFM.Item(strLabel)), ",")
            totalCand = UBound(candidates) - LBound(candidates) + 1
            If StrComp(strLabel, "Depreciation & amortization", vbTextCompare) = 0 Then
                Dim sumFields As String, sumOrigin As String
                If TWTryScaledSum(dictTypes, candidates, strLabel, scaledVal, sumFields, sumOrigin, dictOrigin, CStr(periodKey)) Then
                    If Not dictPeriodSet.Exists(CStr(periodKey)) Then dictPeriodSet.Add CStr(periodKey), True
                    If Not dictIndicatorSet.Exists(strLabel) Then
                        dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                        If Not dictCategoryMap.Exists(strLabel) Then dictCategoryMap.Add strLabel, strCat
                    End If
                    If dictPer.Exists(strLabel) Then
                        dictPer.Item(strLabel) = scaledVal
                    Else
                        dictPer.Add strLabel, scaledVal
                    End If
                    AddOrUpdateTWMatch dictMatchInfo, strLabel, sumFields, sumOrigin, 1, 1, CStr(periodKey)
                End If
                GoTo NextConcept
            End If

            candIdx = 0
            For Each cand In candidates
                candIdx = candIdx + 1
                Dim candKey As String: candKey = Trim$(CStr(cand))
                If dictTypes.Exists(candKey) Then
                    rawVal = dictTypes.Item(candKey)
                    If TWTryScaledValue(rawVal, scaledVal, strLabel) Then
                        If Not dictPeriodSet.Exists(CStr(periodKey)) Then dictPeriodSet.Add CStr(periodKey), True
                        If Not dictIndicatorSet.Exists(strLabel) Then
                            dictIndicatorSet.Add strLabel, dictIndicatorSet.Count
                            If Not dictCategoryMap.Exists(strLabel) Then dictCategoryMap.Add strLabel, strCat
                        End If
                        If dictPer.Exists(strLabel) Then
                            dictPer.Item(strLabel) = scaledVal
                        Else
                            dictPer.Add strLabel, scaledVal
                        End If
                        originName = ""
                        If dictOrigin.Exists(CStr(periodKey) & "|" & candKey) Then _
                            originName = CStr(dictOrigin.Item(CStr(periodKey) & "|" & candKey))
                        AddOrUpdateTWMatch dictMatchInfo, strLabel, candKey, originName, candIdx, totalCand, CStr(periodKey)
                        Exit For
                    End If
                End If
            Next cand
NextConcept:
        Next ci
NextPeriodKey:
    Next periodKey

    stage = "WriteDiagnostics"
    AppendTWDiagnosticsForConceptMap strTicker, strKind, conceptMap, dictMatchInfo, collDiagRows
    Exit Sub

TwErr:
    Dim origNum As Long: origNum = Err.Number
    Dim origSource As String: origSource = Err.Source
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise vbObjectError + 890, "FetchTWFromFinMind", _
        "[stage=" & stage & "] 原始错误号=" & origNum & _
        "; 原始来源=" & origSource & "; " & origDesc
End Sub


Private Sub LoadTWFinMindRows(ByVal strTicker As String, _
                              ByVal strKind As String, _
                              ByVal strQuarter As String, _
                              ByVal lngYear As Long, _
                              ByRef dictRecords As Object, _
                              ByRef dictOrigin As Object)
    Dim datasetName As String: datasetName = FinMindDatasetForKindTW(strKind)
    Dim startDate As String, endDate As String
    If lngYear > 0 Then
        If strKind = "BalanceSheet" Or strKind = "Income" Then
            startDate = CStr(lngYear - 1) & "-01-01"
        Else
            startDate = CStr(lngYear) & "-01-01"
        End If
        endDate = CStr(lngYear) & "-12-31"
    Else
        startDate = Format$(DateAdd("yyyy", -8, Date), "yyyy-mm-dd")
        endDate = Format$(Date, "yyyy-mm-dd")
    End If

    Dim url As String
    url = "https://api.finmindtrade.com/api/v4/data?dataset=" & datasetName & _
          "&data_id=" & NormalizeTWTicker(strTicker) & _
          "&start_date=" & startDate & _
          "&end_date=" & endDate

    Dim result As THttpResult
    Dim raw As String
    raw = RunCachedHttpGet(url, "finmind_TW_" & NormalizeTWTicker(strTicker) & "_" & strKind & "_" & startDate & "_" & endDate, _
                           "FINMIND", GetTtlHoursForSource("FINMIND"), result)

    Dim parsed As Object: Set parsed = JsonConverter.ParseJson(raw)
    If parsed.Exists("status") Then
        If CStr(parsed.Item("status")) <> "200" Then
            Err.Raise vbObjectError + 843, "LoadTWFinMindRows", _
                "FinMind status=" & CStr(parsed.Item("status")) & "; " & TW_NzStr(parsed, "msg")
        End If
    End If
    If Not parsed.Exists("data") Then
        Err.Raise vbObjectError + 844, "LoadTWFinMindRows", "FinMind 响应缺 data"
    End If

    Dim rows As Object: Set rows = parsed.Item("data")
    Dim i As Long
    For i = 1 To rows.Count
        Dim row As Object: Set row = rows.Item(i)
        Dim periodEnd As String: periodEnd = TW_NzStr(row, "date")
        If Len(periodEnd) < 10 Then GoTo NextRow
        ' Quarter filter is applied later in FetchTWFromFinMind — we need every
        ' quarter within the year range so cumulation below has complete YTD slices.
        If lngYear > 0 Then
            Dim periodYr As Long: periodYr = CLng(Left$(periodEnd, 4))
            If strKind = "BalanceSheet" Or strKind = "Income" Then
                If periodYr < lngYear - 1 Or periodYr > lngYear Then GoTo NextRow
            ElseIf periodYr < lngYear Or periodYr > lngYear Then
                GoTo NextRow
            End If
        End If
        Dim typeName As String: typeName = TW_NzStr(row, "type")
        If Len(typeName) = 0 Then GoTo NextRow

        Dim dictTypes As Object
        If dictRecords.Exists(periodEnd) Then
            Set dictTypes = dictRecords.Item(periodEnd)
        Else
            Set dictTypes = CreateObject("Scripting.Dictionary")
            dictTypes.CompareMode = vbTextCompare
            dictRecords.Add periodEnd, dictTypes
        End If
        dictTypes.Item(typeName) = row.Item("value")
        dictOrigin.Item(periodEnd & "|" & typeName) = TW_NzStr(row, "origin_name")
NextRow:
    Next i

    ' FinMind TaiwanStockFinancialStatements (Income) values are single-quarter
    ' as reported per quarter (Q1=3M, Q2=3M only, Q3=3M only, Q4=3M only). We
    ' cumulate them into year-to-date so the standard indicator formulas
    ' (which assume days=90/180/270/365 mirror cumulative revenue / cost) compute
    ' correct annual ratios for DIO / DSO / DPO / ROA / NPM / etc.
    '
    ' TaiwanStockCashFlowsStatement values are already YTD cumulative as
    ' published on MOPS (Q1=3M YTD, Q2=6M YTD, Q3=9M YTD, Q4=12M YTD), so we
    ' DO NOT cumulate them again — that would double-count.
    '
    ' TaiwanStockBalanceSheet values are point-in-time and not cumulated.
    If strKind = "Income" Then
        CumulateTWPeriodicValues dictRecords
    End If
End Sub


Private Sub CumulateTWPeriodicValues(ByRef dictRecords As Object)
    If dictRecords Is Nothing Then Exit Sub
    If dictRecords.Count = 0 Then Exit Sub

    ' Sort period keys ascending so each quarter can absorb the prior YTD value.
    Dim n As Long: n = dictRecords.Count
    Dim periods() As String
    ReDim periods(0 To n - 1)
    Dim idx As Long: idx = 0
    Dim k As Variant
    For Each k In dictRecords.Keys
        periods(idx) = CStr(k)
        idx = idx + 1
    Next k

    Dim i As Long, j As Long, tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If periods(j) < periods(i) Then
                tmp = periods(i)
                periods(i) = periods(j)
                periods(j) = tmp
            End If
        Next j
    Next i

    Dim prevYear As String: prevYear = ""
    Dim acc As Object: Set acc = CreateObject("Scripting.Dictionary")
    acc.CompareMode = vbTextCompare

    Dim p As Long
    For p = 0 To n - 1
        Dim periodKey As String: periodKey = periods(p)
        If Len(periodKey) < 4 Then GoTo NextPeriod
        Dim yearStr As String: yearStr = Left$(periodKey, 4)
        If yearStr <> prevYear Then
            Set acc = CreateObject("Scripting.Dictionary")
            acc.CompareMode = vbTextCompare
            prevYear = yearStr
        End If

        Dim dictTypes As Object: Set dictTypes = dictRecords.Item(periodKey)
        Dim typeKey As Variant
        For Each typeKey In dictTypes.Keys
            Dim typeName As String: typeName = CStr(typeKey)
            If IsTWPointInTimeField(typeName) Then GoTo NextType
            Dim raw As Variant: raw = dictTypes.Item(typeKey)
            If IsNull(raw) Or IsEmpty(raw) Then GoTo NextType
            If Not IsNumeric(raw) Then GoTo NextType

            Dim runningTotal As Double
            runningTotal = CDbl(raw)
            If acc.Exists(typeName) Then runningTotal = runningTotal + CDbl(acc.Item(typeName))
            dictTypes.Item(typeKey) = runningTotal
            acc.Item(typeName) = runningTotal
NextType:
        Next typeKey
NextPeriod:
    Next p
End Sub


Private Function IsTWPointInTimeField(ByVal typeName As String) As Boolean
    Select Case typeName
        Case "CashBalancesBeginningOfPeriod", "CashBalancesEndOfPeriod"
            IsTWPointInTimeField = True
        Case Else
            IsTWPointInTimeField = False
    End Select
End Function


Private Function FinMindDatasetForKindTW(ByVal strKind As String) As String
    Select Case strKind
        Case "BalanceSheet": FinMindDatasetForKindTW = "TaiwanStockBalanceSheet"
        Case "Income":       FinMindDatasetForKindTW = "TaiwanStockFinancialStatements"
        Case "CashFlow":     FinMindDatasetForKindTW = "TaiwanStockCashFlowsStatement"
        Case Else
            Err.Raise vbObjectError + 845, "FinMindDatasetForKindTW", _
                "FinMind TW 不支持报表类型: " & strKind
    End Select
End Function


Private Function FinMindFieldMapForKindTW(ByVal strKind As String) As Object
    Dim mapFM As Object: Set mapFM = CreateObject("Scripting.Dictionary")
    mapFM.CompareMode = vbTextCompare

    Select Case strKind
        Case "BalanceSheet"
            mapFM.Add "Cash & equivalents", "CashAndCashEquivalents"
            mapFM.Add "Accounts receivable, net", "AccountsReceivableNet,AccountsReceivable"
            mapFM.Add "Inventory", "Inventories"
            mapFM.Add "Total current assets", "CurrentAssets"
            mapFM.Add "Accounts payable", "AccountsPayable"
            mapFM.Add "Property, plant & equipment, net", "PropertyPlantAndEquipment"
            mapFM.Add "Investments", "InvestmentAccountedForUsingEquityMethod,FinancialAssetsAtFairvalueThroughOtherComprehensiveIncome,NonCurrentFinancialAssetsAtFairvalueThroughProfitOrLoss"
            mapFM.Add "Total non-current assets", "NoncurrentAssets"
            mapFM.Add "Total assets", "TotalAssets"
            mapFM.Add "Short-term debt", "ShortTermBorrowings,CurrentBorrowings"
            mapFM.Add "Total current liabilities", "CurrentLiabilities"
            mapFM.Add "Long-term debt", "LongtermBorrowings,BondsPayable"
            mapFM.Add "Total non-current liabilities", "NoncurrentLiabilities"
            mapFM.Add "Total liabilities", "Liabilities"
            mapFM.Add "Minority interests", "NoncontrollingInterests"
            mapFM.Add "Total equity", "Equity"
            mapFM.Add "Total stockholders' equity", "EquityAttributableToOwnersOfParent,Equity"
            mapFM.Add "Total liabilities & equity", "TotalLiabilitiesEquity,TotalAssets"

        Case "Income"
            mapFM.Add "Revenue", "Revenue"
            mapFM.Add "Cost of goods & services sold", "CostOfGoodsSold"
            mapFM.Add "Gross profit", "GrossProfit"
            mapFM.Add "Total operating expenses", "OperatingExpenses"
            mapFM.Add "Operating income", "OperatingIncome"
            mapFM.Add "Pre-tax income", "PreTaxIncome"
            mapFM.Add "Income tax expense", "TAX"
            mapFM.Add "Net income", "EquityAttributableToOwnersOfParent,IncomeAfterTaxes"
            mapFM.Add "Basic EPS", "EPS"
            mapFM.Add "Diluted EPS", "EPS"

        Case "CashFlow"
            mapFM.Add "Cash from operations", "CashFlowsFromOperatingActivities,NetCashInflowFromOperatingActivities,CashReceivedThroughOperations"
            mapFM.Add "Depreciation & amortization", "Depreciation,AmortizationExpense"
            mapFM.Add "Cash from investing", "CashProvidedByInvestingActivities"
            mapFM.Add "Capex", "PropertyAndPlantAndEquipment"
            mapFM.Add "Cash from financing", "CashFlowsProvidedFromFinancingActivities"
            mapFM.Add "Dividends paid", "CashDividendsPaid,DistributionOfCashDividends"
            mapFM.Add "Interest paid", "PayTheInterest,InterestExpense"
            mapFM.Add "Interest received", "InterestIncome"
            mapFM.Add "FX effect on cash", "EffectsOfExchangeRateChangesOnCashAndCashEquivalents"
            mapFM.Add "Cash at beginning of period", "CashBalancesBeginningOfPeriod"
            mapFM.Add "Cash at end of period", "CashBalancesEndOfPeriod"

        Case Else
            Err.Raise vbObjectError + 846, "FinMindFieldMapForKindTW", _
                "FinMind TW 不支持报表类型: " & strKind
    End Select

    Set FinMindFieldMapForKindTW = mapFM
End Function


Private Function MatchTWPeriod(ByVal periodEnd As String, _
                               ByVal strQuarter As String, _
                               ByVal lngYear As Long, _
                               ByVal strKind As String) As Boolean
    MatchTWPeriod = False
    If Len(periodEnd) < 10 Then Exit Function

    If lngYear > 0 Then
        Dim periodYear As Long: periodYear = CLng(Left$(periodEnd, 4))
        If (strKind = "BalanceSheet" Or strKind = "Income") Then
            If periodYear <> lngYear And periodYear <> lngYear - 1 Then Exit Function
        ElseIf periodYear <> lngYear Then
            Exit Function
        End If
    End If

    Select Case UCase$(Trim$(strQuarter))
        Case "全部", ""
            MatchTWPeriod = True
        Case "Q1"
            MatchTWPeriod = (Right$(periodEnd, 5) = "03-31")
        Case "Q2"
            MatchTWPeriod = (Right$(periodEnd, 5) = "06-30")
        Case "Q3"
            MatchTWPeriod = (Right$(periodEnd, 5) = "09-30")
        Case "Q4"
            MatchTWPeriod = (Right$(periodEnd, 5) = "12-31")
        Case Else
            MatchTWPeriod = True
    End Select
End Function


Private Function NormalizeTWTicker(ByVal rawCode As String) As String
    Dim s As String: s = UCase$(Trim$(rawCode))
    s = Replace(s, ".TW", "", 1, -1, vbTextCompare)
    s = Replace(s, ".TWO", "", 1, -1, vbTextCompare)
    s = Replace(s, "TWSE:", "", 1, -1, vbTextCompare)
    s = Replace(s, "TPEX:", "", 1, -1, vbTextCompare)
    s = Replace(s, "TW", "", 1, -1, vbTextCompare)
    s = Replace(s, " ", "")
    NormalizeTWTicker = s
End Function


Private Function TWTryScaledValue(ByVal rawValue As Variant, _
                                  ByRef scaledValue As Double, _
                                  Optional ByVal strLabel As String = "") As Boolean
    If IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    Dim normalized As Variant
    normalized = NormalizeValue(CStr(rawValue))
    If IsEmpty(normalized) Or IsNull(normalized) Then Exit Function
    If Not IsNumeric(normalized) Then Exit Function

    If TWIsPerShareLabel(strLabel) Then
        scaledValue = CDbl(normalized)
    Else
        scaledValue = CDbl(normalized) / 1000000#
    End If
    TWTryScaledValue = True
End Function


Private Function TWTryScaledSum(ByVal dictTypes As Object, _
                                ByVal candidates As Variant, _
                                ByVal strLabel As String, _
                                ByRef scaledValue As Double, _
                                ByRef fieldText As String, _
                                ByRef originText As String, _
                                ByVal dictOrigin As Object, _
                                ByVal periodKey As String) As Boolean
    Dim cand As Variant, candKey As String, oneVal As Double
    scaledValue = 0#
    fieldText = ""
    originText = ""

    For Each cand In candidates
        candKey = Trim$(CStr(cand))
        If Len(candKey) > 0 And dictTypes.Exists(candKey) Then
            If TWTryScaledValue(dictTypes.Item(candKey), oneVal, strLabel) Then
                scaledValue = scaledValue + oneVal
                If Len(fieldText) > 0 Then fieldText = fieldText & "+"
                fieldText = fieldText & candKey
                If Not dictOrigin Is Nothing Then
                    If dictOrigin.Exists(periodKey & "|" & candKey) Then
                        If Len(originText) > 0 Then originText = originText & "+"
                        originText = originText & CStr(dictOrigin.Item(periodKey & "|" & candKey))
                    End If
                End If
                TWTryScaledSum = True
            End If
        End If
    Next cand
End Function


Private Function TWIsPerShareLabel(ByVal strLabel As String) As Boolean
    TWIsPerShareLabel = (InStr(1, strLabel, "EPS", vbTextCompare) > 0)
End Function


Private Sub AddOrUpdateTWMatch(ByVal dictMatchInfo As Object, _
                               ByVal label As String, _
                               ByVal fieldName As String, _
                               ByVal originName As String, _
                               ByVal candIdx As Long, _
                               ByVal totalCand As Long, _
                               ByVal periodEnd As String)
    Dim scoreText As String: scoreText = CStr(candIdx) & "/" & CStr(totalCand)
    Dim fieldText As String
    fieldText = fieldName
    If Len(originName) > 0 Then fieldText = fieldText & " / " & originName

    If dictMatchInfo.Exists(label) Then
        Dim oldInfo As Variant: oldInfo = dictMatchInfo.Item(label)
        Dim oldIdx As Long: oldIdx = CLng(Split(CStr(oldInfo(2)), "/")(0))
        If candIdx > oldIdx Then Exit Sub
        If candIdx = oldIdx Then
            If periodEnd > CStr(oldInfo(3)) Then oldInfo(3) = periodEnd
            dictMatchInfo.Item(label) = oldInfo
            Exit Sub
        End If
        dictMatchInfo.Item(label) = Array(fieldText, "TWD millions", scoreText, periodEnd)
    Else
        dictMatchInfo.Add label, Array(fieldText, "TWD millions", scoreText, periodEnd)
    End If
End Sub


Private Sub AppendTWDiagnosticsForConceptMap(ByVal ticker As String, _
                                             ByVal strKind As String, _
                                             ByVal conceptMap As Variant, _
                                             ByVal dictMatchInfo As Object, _
                                             ByVal collRows As Collection)
    If collRows Is Nothing Then Exit Sub

    Dim i As Long, mapEntry As Variant, label As String, info As Variant
    For i = LBound(conceptMap) To UBound(conceptMap)
        mapEntry = conceptMap(i)
        label = MapEntryLabel(mapEntry)
        If dictMatchInfo.Exists(label) Then
            info = dictMatchInfo.Item(label)
            Dim fxText As String: fxText = FxRateTextForDiagnostic("TWD", CStr(info(3)), strKind)
            AddDiagnosticRow collRows, ticker, strKind, label, "OK_FINMIND", _
                             "FinMind", "JSON", CStr(info(0)), _
                             CStr(info(1)), CStr(info(2)), "exact type match via FinMind JSON", fxText
        Else
            AddDiagnosticRow collRows, ticker, strKind, label, "MISSING", _
                             "FinMind", "JSON", "—", "TWD millions", "—", _
                             "FinMind dataset did not contain mapped type"
        End If
    Next i
End Sub


Private Function HasAnyCoreLabelTW(ByVal dictData As Object, _
                                   ByVal ticker As String, _
                                   ByVal coreLabels As Variant) As Boolean
    On Error Resume Next
    If Not dictData.Exists(ticker) Then GoTo CleanExit
    Dim dictCompany As Object: Set dictCompany = dictData.Item(ticker)
    Dim p As Variant, i As Long, dictPer As Object
    For Each p In dictCompany.Keys
        Set dictPer = dictCompany.Item(p)
        For i = LBound(coreLabels) To UBound(coreLabels)
            If dictPer.Exists(CStr(coreLabels(i))) Then
                HasAnyCoreLabelTW = True
                GoTo CleanExit
            End If
        Next i
    Next p
CleanExit:
    Err.Clear
    On Error GoTo 0
End Function


Private Sub CommitTempTWData(ByVal strCode As String, _
                             ByVal tempData As Object, _
                             ByVal tempPeriodSet As Object, _
                             ByVal tempIndicatorSet As Object, _
                             ByVal tempCategoryMap As Object, _
                             ByRef dictData As Object, _
                             ByRef dictPeriodSet As Object, _
                             ByRef dictIndicatorSet As Object, _
                             ByRef dictCategoryMap As Object)
    If tempData.Exists(strCode) Then
        If dictData.Exists(strCode) Then dictData.Remove strCode
        dictData.Add strCode, tempData.Item(strCode)
    End If

    Dim k As Variant
    For Each k In tempPeriodSet.Keys
        If Not dictPeriodSet.Exists(CStr(k)) Then dictPeriodSet.Add CStr(k), True
    Next k
    For Each k In tempIndicatorSet.Keys
        If Not dictIndicatorSet.Exists(CStr(k)) Then dictIndicatorSet.Add CStr(k), dictIndicatorSet.Count
    Next k
    For Each k In tempCategoryMap.Keys
        If Not dictCategoryMap.Exists(CStr(k)) Then dictCategoryMap.Add CStr(k), tempCategoryMap.Item(k)
    Next k
End Sub


Private Function OrderIndicatorsByTWConceptMap(ByVal dictIndicatorSet As Object, _
                                               ByVal conceptMap As Variant) As Variant
    Dim coll As New Collection
    Dim i As Long, label As String
    For i = LBound(conceptMap) To UBound(conceptMap)
        label = MapEntryLabel(conceptMap(i))
        If dictIndicatorSet.Exists(label) Then coll.Add label
    Next i

    If coll.Count = 0 Then
        OrderIndicatorsByTWConceptMap = Array()
        Exit Function
    End If

    Dim arr() As String
    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count
        arr(i) = CStr(coll.Item(i))
    Next i
    OrderIndicatorsByTWConceptMap = arr
End Function


Private Sub ClearTWWideTableOutput(ByVal wsTarget As Worksheet)
    On Error Resume Next
    wsTarget.UsedRange.UnMerge
    Dim lastRow As Long, lastCol As Long
    lastRow = wsTarget.UsedRange.Rows.Count + wsTarget.UsedRange.Row - 1
    lastCol = wsTarget.UsedRange.Columns.Count + wsTarget.UsedRange.Column - 1
    If lastRow < 2 Then lastRow = 2
    If lastCol < 3 Then lastCol = 3
    wsTarget.Range(wsTarget.Cells(1, 3), wsTarget.Cells(lastRow, lastCol)).Clear
    If lastRow >= 2 Then wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(lastRow, 2)).Clear
    wsTarget.Range("A1").Value = "大类"
    wsTarget.Range("B1").Value = "指标名称"
    With wsTarget.Range("A1:B1")
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Err.Clear
    On Error GoTo 0
End Sub


Private Function TW_NzStr(ByVal dict As Object, ByVal key As String) As String
    If dict Is Nothing Then TW_NzStr = "": Exit Function
    If Not dict.Exists(key) Then TW_NzStr = "": Exit Function
    Dim v As Variant: v = dict.Item(key)
    If IsNull(v) Or IsEmpty(v) Then
        TW_NzStr = ""
    Else
        TW_NzStr = CStr(v)
    End If
End Function
