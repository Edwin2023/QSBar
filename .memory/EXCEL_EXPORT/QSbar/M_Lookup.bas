Attribute VB_Name = "M_Lookup"
Option Explicit

'======================================================================
' 超高性能VLOOKUP函数模块
' 专门优化整列引用和大数据量场景
' 版本：3.0 (极速版)
'======================================================================

Public Enum matchType
    ExactMatch = 0
    FuzzyMatch = 1
    ContainsMatch = 2
    StartsWithMatch = 3
    EndsWithMatch = 4
End Enum

'======================================================================
' 超快速查找函数 - 专门优化整列引用
'======================================================================
Public Function SmartLookup(lookupValue As Variant, _
                           tableArray As Range, _
                           returnRange As Range, _
                           Optional similarity As Double = 0.6) As Variant
    
    On Error GoTo ErrorHandler
    
    ' 快速获取查找值
    Dim searchValue As String
    searchValue = GetLookupValueFast(lookupValue)
    If searchValue = "" Then
        SmartLookup = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' 预处理查找值
    searchValue = UCase(Trim(searchValue))
    
    ' 获取实际数据范围（跳过空行）
    Dim actualTableRange As Range
    Dim actualReturnRange As Range
    
    Set actualTableRange = GetUsedRange(tableArray)
    Set actualReturnRange = GetCorrespondingRange(returnRange, tableArray, actualTableRange)
    
    If actualTableRange Is Nothing Or actualReturnRange Is Nothing Then
        SmartLookup = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' 使用数组一次性读取数据，避免逐个单元格访问
    Dim searchArray As Variant
    Dim returnArray As Variant
    
    searchArray = actualTableRange.Value
    returnArray = actualReturnRange.Value
    
    ' 超快速数组遍历
    Dim i As Long
    Dim currentValue As String
    Dim bestMatch As Long
    Dim bestScore As Double
    
    bestMatch = -1
    bestScore = 0
    
    ' 确保是二维数组
    If Not IsArray(searchArray) Then
        ' 单个值情况
        currentValue = UCase(Trim(CStr(searchArray)))
        Dim singleScore As Double
        singleScore = CalculateSmartSimilarity(searchValue, currentValue)
        If singleScore >= similarity Then
            SmartLookup = returnArray
        Else
            SmartLookup = CVErr(xlErrNA)
        End If
        Exit Function
    End If
    
    ' 处理一维数组
    Dim rowCount As Long
    If UBound(searchArray, 2) = 1 Then
        rowCount = UBound(searchArray, 1)
    Else
        rowCount = UBound(searchArray, 1)
    End If
    
    ' 提取搜索关键词
    Dim searchKeywords As Variant
    searchKeywords = ExtractKeywords(searchValue)
    
    ' 快速遍历数组
    For i = 1 To rowCount
        ' 跳过空值
        If Not IsEmpty(searchArray(i, 1)) Then
            currentValue = UCase(Trim(CStr(searchArray(i, 1))))
            
            ' 精确匹配优先
            If currentValue = searchValue Then
                SmartLookup = returnArray(i, 1)
                Exit Function
            End If
            
            ' 智能相似度计算
            Dim score As Double
            score = CalculateSmartSimilarity(searchValue, currentValue)
            
            If score > bestScore And score >= similarity Then
                bestScore = score
                bestMatch = i
                
                ' 如果匹配度很高，直接返回
                If score > 0.9 Then
                    SmartLookup = returnArray(bestMatch, 1)
                    Exit Function
                End If
            End If
        End If
        
        ' 每10000行检查一次，避免Excel无响应
        If i Mod 10000 = 0 Then DoEvents
    Next i
    
    ' 返回最佳匹配
    If bestMatch > 0 Then
        SmartLookup = returnArray(bestMatch, 1)
    Else
        SmartLookup = CVErr(xlErrNA)
    End If
    
    Exit Function
    
ErrorHandler:
    SmartLookup = CVErr(xlErrValue)
End Function

'======================================================================
' 智能相似度计算函数
'======================================================================
Private Function CalculateSmartSimilarity(searchValue As String, currentValue As String) As Double
    
    ' 精确匹配
    If searchValue = currentValue Then
        CalculateSmartSimilarity = 1
        Exit Function
    End If
    
    ' 简单包含匹配
    If InStr(currentValue, searchValue) > 0 Then
        CalculateSmartSimilarity = CDbl(Len(searchValue)) / CDbl(Len(currentValue))
        Exit Function
    End If
    
    If InStr(searchValue, currentValue) > 0 Then
        CalculateSmartSimilarity = CDbl(Len(currentValue)) / CDbl(Len(searchValue)) * 0.9
        Exit Function
    End If
    
    ' 提取关键词进行匹配
    Dim searchKeywords As Variant
    Dim currentKeywords As Variant
    
    searchKeywords = ExtractKeywords(searchValue)
    currentKeywords = ExtractKeywords(currentValue)
    
    ' 计算关键词匹配度
    Dim keywordScore As Double
    keywordScore = CalculateKeywordSimilarity(searchKeywords, currentKeywords)
    
    ' 计算字符级相似度（简化版）
    Dim charScore As Double
    charScore = CalculateCharSimilarity(searchValue, currentValue)
    
    ' 综合评分：关键词匹配更重要
    CalculateSmartSimilarity = keywordScore * 0.7 + charScore * 0.3
    
End Function

'======================================================================
' 提取关键词函数
'======================================================================
Private Function ExtractKeywords(text As String) As Variant
    
    ' 移除常见停用词和标点符号
    Dim cleanText As String
    cleanText = text
    
    ' 替换常见分隔符为空格
    cleanText = Replace(cleanText, ";", " ")
    cleanText = Replace(cleanText, ",", " ")
    cleanText = Replace(cleanText, ".", " ")
    cleanText = Replace(cleanText, "-", " ")
    cleanText = Replace(cleanText, "_", " ")
    cleanText = Replace(cleanText, "/", " ")
    cleanText = Replace(cleanText, "(", " ")
    cleanText = Replace(cleanText, ")", " ")
    
    ' 移除常见停用词
    Dim stopWords As Variant
    stopWords = Array("THE", "AND", "OR", "TO", "FROM", "WITH", "AS", "PER", "FOR", "OF", "IN", "ON", "AT", "BY", "IS", "ARE", "WAS", "WERE", "BE", "BEEN", "HAVE", "HAS", "HAD", "A", "AN", "MM", "THICK")
    
    Dim i As Integer
    For i = 0 To UBound(stopWords)
        cleanText = Replace(cleanText, " " & stopWords(i) & " ", " ")
    Next i
    
    ' 移除数字和单位（保留重要数字）
    cleanText = RemoveUnimportantNumbers(cleanText)
    
    ' 分割为单词数组
    Dim words As Variant
    words = Split(Trim(cleanText), " ")
    
    ' 过滤空字符串和短词
    Dim result() As String
    Dim resultCount As Integer
    resultCount = 0
    
    For i = 0 To UBound(words)
        If Len(Trim(words(i))) > 2 Then
            ReDim Preserve result(resultCount)
            result(resultCount) = Trim(words(i))
            resultCount = resultCount + 1
        End If
    Next i
    
    If resultCount > 0 Then
        ExtractKeywords = result
    Else
        ExtractKeywords = Array()
    End If
    
End Function

'======================================================================
' 移除不重要数字函数
'======================================================================
Private Function RemoveUnimportantNumbers(text As String) As String
    
    Dim result As String
    result = text
    
    ' 移除常见单位和数字组合
    result = RegexReplace(result, "\d+(\.\d+)?\s*(MM|CM|M|KG|G|%)", "")
    result = RegexReplace(result, "\d+(\.\d+)?\s*THICK", "")
    result = RegexReplace(result, "\d+(ST|ND|RD|TH)\s*FLOOR", "FLOOR")
    
    ' 清理多余空格
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    RemoveUnimportantNumbers = Trim(result)
    
End Function

'======================================================================
' 简化的正则替换函数
'======================================================================
Private Function RegexReplace(text As String, pattern As String, replacement As String) As String
    
    ' 简化实现，处理常见模式
    Dim result As String
    result = text
    
    ' 处理数字+单位模式
    If InStr(pattern, "MM|CM|M|KG|G|%") > 0 Then
        Dim i As Integer
        Dim char As String
        Dim inNumber As Boolean
        Dim numberStart As Integer
        
        For i = 1 To Len(text)
            char = Mid(text, i, 1)
            If IsNumeric(char) Or char = "." Then
                If Not inNumber Then
                    inNumber = True
                    numberStart = i
                End If
            ElseIf inNumber And (char = " " Or char = "M" Or char = "C" Or char = "K" Or char = "G" Or char = "%") Then
                ' 检查是否是单位
                Dim nextChars As String
                nextChars = UCase(Mid(text, i, 4))
                If InStr(nextChars, "MM") = 1 Or InStr(nextChars, "CM") = 1 Or InStr(nextChars, "KG") = 1 Or _
                   Left(nextChars, 1) = "M" Or Left(nextChars, 1) = "G" Or Left(nextChars, 1) = "%" Then
                    ' 移除数字和单位
                    result = Left(result, numberStart - 1) & " " & Mid(result, i + 3)
                    Exit For
                End If
                inNumber = False
            Else
                inNumber = False
            End If
        Next i
    End If
    
    RegexReplace = result
    
End Function

'======================================================================
' 关键词相似度计算
'======================================================================
Private Function CalculateKeywordSimilarity(keywords1 As Variant, keywords2 As Variant) As Double
    
    If Not IsArray(keywords1) Or Not IsArray(keywords2) Then
        CalculateKeywordSimilarity = 0
        Exit Function
    End If
    
    If UBound(keywords1) < 0 Or UBound(keywords2) < 0 Then
        CalculateKeywordSimilarity = 0
        Exit Function
    End If
    
    Dim matchCount As Integer
    Dim i As Integer, j As Integer
    matchCount = 0
    
    ' 计算匹配的关键词数量
    For i = 0 To UBound(keywords1)
        For j = 0 To UBound(keywords2)
            ' 完全匹配
            If keywords1(i) = keywords2(j) Then
                matchCount = matchCount + 1
                Exit For
            End If
            ' 部分匹配（一个词包含另一个）
            If Len(keywords1(i)) > 3 And Len(keywords2(j)) > 3 Then
                If InStr(keywords1(i), keywords2(j)) > 0 Or InStr(keywords2(j), keywords1(i)) > 0 Then
                    matchCount = matchCount + 0.7
                    Exit For
                End If
            End If
        Next j
    Next i
    
    ' 计算相似度
    Dim totalKeywords As Integer
    totalKeywords = UBound(keywords1) + 1 + UBound(keywords2) + 1
    
    If totalKeywords > 0 Then
        CalculateKeywordSimilarity = (2 * matchCount) / totalKeywords
    Else
        CalculateKeywordSimilarity = 0
    End If
    
    ' 确保不超过1
    If CalculateKeywordSimilarity > 1 Then CalculateKeywordSimilarity = 1
    
End Function

'======================================================================
' 字符相似度计算（简化版）
'======================================================================
Private Function CalculateCharSimilarity(str1 As String, str2 As String) As Double
    
    Dim len1 As Integer, len2 As Integer
    Dim minLen As Integer, maxLen As Integer
    Dim commonChars As Integer
    Dim i As Integer
    
    len1 = Len(str1)
    len2 = Len(str2)
    
    If len1 = 0 And len2 = 0 Then
        CalculateCharSimilarity = 1
        Exit Function
    End If
    
    If len1 = 0 Or len2 = 0 Then
        CalculateCharSimilarity = 0
        Exit Function
    End If
    
    minLen = IIf(len1 < len2, len1, len2)
    maxLen = IIf(len1 > len2, len1, len2)
    
    ' 长度差异过大，返回低分
    If maxLen > minLen * 2 Then
        CalculateCharSimilarity = 0.1
        Exit Function
    End If
    
    commonChars = 0
    
    ' 计算相同位置的字符数
    For i = 1 To minLen
        If Mid(str1, i, 1) = Mid(str2, i, 1) Then
            commonChars = commonChars + 1
        End If
    Next i
    
    CalculateCharSimilarity = CDbl(commonChars) / CDbl(maxLen)
    
End Function

'======================================================================
' 高性能版AdvancedVLookup - 支持多种匹配模式
'======================================================================
Public Function AdvancedVLookup(lookupValue As Variant, _
                                tableArray As Range, _
                                returnRange As Range, _
                                Optional matchType As matchType = ContainsMatch, _
                                Optional similarity As Double = 0.7) As Variant
    
    On Error GoTo ErrorHandler
    
    ' 快速获取查找值
    Dim searchValue As String
    searchValue = GetLookupValueFast(lookupValue)
    If searchValue = "" Then
        AdvancedVLookup = CVErr(xlErrNA)
        Exit Function
    End If
    
    searchValue = UCase(Trim(searchValue))
    
    ' 获取实际数据范围
    Dim actualTableRange As Range
    Dim actualReturnRange As Range
    
    Set actualTableRange = GetUsedRange(tableArray)
    Set actualReturnRange = GetCorrespondingRange(returnRange, tableArray, actualTableRange)
    
    If actualTableRange Is Nothing Or actualReturnRange Is Nothing Then
        AdvancedVLookup = CVErr(xlErrNA)
        Exit Function
    End If
    
    ' 使用数组读取数据
    Dim searchArray As Variant
    Dim returnArray As Variant
    
    searchArray = actualTableRange.Value
    returnArray = actualReturnRange.Value
    
    ' 处理单值情况
    If Not IsArray(searchArray) Then
        AdvancedVLookup = ProcessSingleValue(searchValue, CStr(searchArray), returnArray, matchType)
        Exit Function
    End If
    
    ' 数组遍历查找
    Dim i As Long
    Dim currentValue As String
    Dim bestMatch As Long
    Dim bestScore As Double
    Dim currentScore As Double
    
    bestMatch = -1
    bestScore = 0
    
    Dim rowCount As Long
    rowCount = UBound(searchArray, 1)
    
    For i = 1 To rowCount
        If Not IsEmpty(searchArray(i, 1)) Then
            currentValue = UCase(Trim(CStr(searchArray(i, 1))))
            
            ' 根据匹配类型计算得分
            currentScore = CalculateScore(searchValue, currentValue, matchType)
            
            If currentScore > bestScore And currentScore >= similarity Then
                bestScore = currentScore
                bestMatch = i
                
                ' 完全匹配直接返回
                If currentScore = 1 Then
                    AdvancedVLookup = returnArray(bestMatch, 1)
                    Exit Function
                End If
            End If
        End If
        
        If i Mod 5000 = 0 Then DoEvents
    Next i
    
    If bestMatch > 0 Then
        AdvancedVLookup = returnArray(bestMatch, 1)
    Else
        AdvancedVLookup = CVErr(xlErrNA)
    End If
    
    Exit Function
    
ErrorHandler:
    AdvancedVLookup = CVErr(xlErrValue)
End Function

'======================================================================
' 极速查找函数 - 仅精确和包含匹配
'======================================================================
Public Function FastLookup(lookupValue As Variant, _
                          tableArray As Range, _
                          returnRange As Range) As Variant
    
    On Error GoTo ErrorHandler
    
    Dim searchValue As String
    searchValue = GetLookupValueFast(lookupValue)
    If searchValue = "" Then
        FastLookup = CVErr(xlErrNA)
        Exit Function
    End If
    
    searchValue = UCase(Trim(searchValue))
    
    ' 使用Find方法进行超快速查找
    Dim foundCell As Range
    Set foundCell = tableArray.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        ' 精确匹配找到
        Dim rowOffset As Long
        rowOffset = foundCell.Row - tableArray.Row
        FastLookup = returnRange.Cells(rowOffset + 1, 1).Value
        Exit Function
    End If
    
    ' 如果精确匹配失败，使用包含匹配
    Set foundCell = tableArray.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not foundCell Is Nothing Then
        rowOffset = foundCell.Row - tableArray.Row
        FastLookup = returnRange.Cells(rowOffset + 1, 1).Value
    Else
        FastLookup = CVErr(xlErrNA)
    End If
    
    Exit Function
    
ErrorHandler:
    FastLookup = CVErr(xlErrValue)
End Function

'======================================================================
' 辅助函数：快速获取查找值
'======================================================================
Private Function GetLookupValueFast(lookupValue As Variant) As String
    On Error GoTo ErrorHandler
    
    If VarType(lookupValue) = vbString Then
        GetLookupValueFast = lookupValue
    ElseIf IsNumeric(lookupValue) Then
        GetLookupValueFast = CStr(lookupValue)
    ElseIf TypeName(lookupValue) = "Range" Then
        If Not Intersect(lookupValue, Application.Caller) Is Nothing Then
            GetLookupValueFast = ""
        Else
            GetLookupValueFast = CStr(lookupValue.Value)
        End If
    Else
        GetLookupValueFast = CStr(lookupValue)
    End If
    Exit Function
    
ErrorHandler:
    GetLookupValueFast = ""
End Function

'======================================================================
' 辅助函数：获取实际使用的范围（排除空行）
'======================================================================
Private Function GetUsedRange(inputRange As Range) As Range
    On Error GoTo ErrorHandler
    
    ' 如果是整列引用，获取实际使用的范围
    If inputRange.Rows.Count > 50000 Then
        Dim ws As Worksheet
        Set ws = inputRange.Worksheet
        
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, inputRange.Column).End(xlUp).Row
        
        If lastRow >= inputRange.Row Then
            Set GetUsedRange = ws.Range(ws.Cells(inputRange.Row, inputRange.Column), _
                                       ws.Cells(lastRow, inputRange.Column))
        Else
            Set GetUsedRange = Nothing
        End If
    Else
        Set GetUsedRange = inputRange
    End If
    
    Exit Function
    
ErrorHandler:
    Set GetUsedRange = inputRange
End Function

'======================================================================
' 辅助函数：获取对应的返回范围
'======================================================================
Private Function GetCorrespondingRange(returnRange As Range, _
                                     originalTableRange As Range, _
                                     actualTableRange As Range) As Range
    On Error GoTo ErrorHandler
    
    If actualTableRange Is Nothing Then
        Set GetCorrespondingRange = Nothing
        Exit Function
    End If
    
    ' 如果是整列引用，调整返回范围
    If originalTableRange.Rows.Count > 50000 Then
        Dim ws As Worksheet
        Set ws = returnRange.Worksheet
        
        Dim startRow As Long
        Dim endRow As Long
        startRow = actualTableRange.Row
        endRow = actualTableRange.Row + actualTableRange.Rows.Count - 1
        
        Set GetCorrespondingRange = ws.Range(ws.Cells(startRow, returnRange.Column), _
                                           ws.Cells(endRow, returnRange.Column))
    Else
        Set GetCorrespondingRange = returnRange
    End If
    
    Exit Function
    
ErrorHandler:
    Set GetCorrespondingRange = returnRange
End Function

'======================================================================
' 辅助函数：处理单值情况
'======================================================================
Private Function ProcessSingleValue(searchValue As String, _
                                   currentValue As String, _
                                   returnValue As Variant, _
                                   matchType As matchType) As Variant
    
    currentValue = UCase(Trim(currentValue))
    
    Select Case matchType
        Case ExactMatch
            If searchValue = currentValue Then
                ProcessSingleValue = returnValue
            Else
                ProcessSingleValue = CVErr(xlErrNA)
            End If
        Case ContainsMatch
            If InStr(currentValue, searchValue) > 0 Or InStr(searchValue, currentValue) > 0 Then
                ProcessSingleValue = returnValue
            Else
                ProcessSingleValue = CVErr(xlErrNA)
            End If
        Case Else
            If searchValue = currentValue Or InStr(currentValue, searchValue) > 0 Then
                ProcessSingleValue = returnValue
            Else
                ProcessSingleValue = CVErr(xlErrNA)
            End If
    End Select
    
End Function

'======================================================================
' 辅助函数：快速计算匹配得分
'======================================================================
Private Function CalculateScore(searchValue As String, _
                               currentValue As String, _
                               matchType As matchType) As Double
    
    Select Case matchType
        Case ExactMatch
            CalculateScore = IIf(searchValue = currentValue, 1, 0)
            
        Case ContainsMatch
            If InStr(currentValue, searchValue) > 0 Then
                CalculateScore = CDbl(Len(searchValue)) / CDbl(Len(currentValue))
            ElseIf InStr(searchValue, currentValue) > 0 Then
                CalculateScore = CDbl(Len(currentValue)) / CDbl(Len(searchValue)) * 0.8
            Else
                CalculateScore = 0
            End If
            
        Case StartsWithMatch
            If Len(currentValue) >= Len(searchValue) And _
               Left(currentValue, Len(searchValue)) = searchValue Then
                CalculateScore = CDbl(Len(searchValue)) / CDbl(Len(currentValue))
            Else
                CalculateScore = 0
            End If
            
        Case EndsWithMatch
            If Len(currentValue) >= Len(searchValue) And _
               Right(currentValue, Len(searchValue)) = searchValue Then
                CalculateScore = CDbl(Len(searchValue)) / CDbl(Len(currentValue))
            Else
                CalculateScore = 0
            End If
            
        Case FuzzyMatch
            ' 简化的模糊匹配，避免复杂算法
            If searchValue = currentValue Then
                CalculateScore = 1
            ElseIf InStr(currentValue, searchValue) > 0 Then
                CalculateScore = CDbl(Len(searchValue)) / CDbl(Len(currentValue)) * 0.9
            ElseIf InStr(searchValue, currentValue) > 0 Then
                CalculateScore = CDbl(Len(currentValue)) / CDbl(Len(searchValue)) * 0.7
            Else
                CalculateScore = 0
            End If
            
        Case Else
            CalculateScore = IIf(searchValue = currentValue, 1, 0)
    End Select
    
End Function

