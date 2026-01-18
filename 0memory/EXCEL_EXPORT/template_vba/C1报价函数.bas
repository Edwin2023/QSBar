Attribute VB_Name = "C1报价函数"
'
''---------------根据首列获取单元格------------------

'扩展版：可从 "1000mm宽 × 2400mm高的甲级防火门" 中提取 1000*2400
Function ExtExp(txt As String) As String
    Dim regEx As Object
    Dim matches As Object
    Dim m As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    
    With regEx
        ' 改良后的规格匹配：
        ' 数字 + 可选单位 + 可选中文（宽/高/厚/长等）
        ' + 乘号
        ' + 数字 + 可选单位 + 可选中文
        .Pattern = "(\d+(\.\d+)?)[ ]*(mm|cm|m)?[ ]*(宽|高|厚|长)?[ ]*[*x×][ ]*(\d+(\.\d+)?)[ ]*(mm|cm|m)?[ ]*(宽|高|厚|长)?"
        .IgnoreCase = True
        .Global = False
    End With
    
    If regEx.Test(txt) Then
        Set matches = regEx.Execute(txt)
        Set m = matches(0)
        
        ' m.SubMatches:
        ' 0 = 第一个数字
        ' 4 = 第二个数字（注意分组顺序）
        ExtExp = m.SubMatches(0) & "*" & m.SubMatches(4)
    Else
        ExtExp = ""
    End If
End Function



'提取数字
Function ExtNum(rng As Range) As String
    Dim s As String, i As Integer, ch As String
    s = rng.Value
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Or ch = "," Then
            ExtNum = ExtNum & ch
        End If
    Next i
End Function

' 公共函数 JS：计算Excel单元格区域中的表达式，并返回结果
Public Function JS(C) As Double

    ' 错误处理：遇到错误时跳过至下一行代码
    On Error Resume Next

    ' 遍历传入的单元格区域
    For k = 1 To C.Count
        ' 获取当前单元格的值
        Debug.Print C.item(k)
        P = C.item(k)

        ' 如果单元格为空，则将函数返回值设为0并退出循环
        If P = "" Then
            JS = ""
            Exit For
        Else
            ' 替换掉非标准符号，以便于计算
            ' A：包含非标准符号的数组
            ' B：包含对应标准符号的数组
            a = Array("【", "】", "＋", "－", "×", "÷", "（", "）", "，", "。", Chr(10), " ")
            B = Array("[", "]", "+", "-", "*", "/", "(", ")", ",", ".", "", "")

            ' 用标准符号替换非标准符号
            For n = 0 To UBound(a)
                P = Replace(P, a(n), B(n))
            Next n

            ' 移除用中括号包围的注释部分
            For i = 1 To Len(P) - Len(Replace(P, "[", ""))
                P = Left$(P, InStr(P, "[") - 1) & Right$(P, Len(P) - InStr(P, "]"))
            Next i

            ' 计算处理后的表达式，并累加到结果变量 JS
            JS = JS + Application.Evaluate(P)
        End If
    Next k

    ' 返回计算结果
     Application.Volatile 可以开启以确保函数响应所有环境变化
    ' Application.Volatile '如果使用自定义变量引用则需要开启
End Function



Function FORMULAR(rng As Range)

     Application.Volatile 可以开启以确保函数响应所有环境变化
    ' Application.Volatile '如果使用自定义变量引用则需要开启
    
    F = Application.Evaluate(WorksheetFunction.Substitute(WorksheetFunction.Substitute(rng, "[", "*ISTEXT(""["), "]", "]"")"))

    If IsError(F) Then
        FORMULAR = ""
    Else
        FORMULAR = F
    End If
    ' 返回计算结果


End Function
'
'
'
'返回日期格式：2018-1-1
'Function myDate(rng As Range) As String
'
'    On Error GoTo Err:
'    Dim DayDate As Date
'    DayDate = CDate(rng.value)
''    Debug.Print Replace(DayDate, "/", "-")
''    Debug.Print Application.WorksheetFunction.Year(DayDate)
'    myDate = Replace(DayDate, "/", "-")
'Err:
'    myDate = rng.value
'
'End Function
'
''在第一列找到最后一个对应值，返回单元格对象
'Function ENDCELL(rng As Range, Optional myCriteria As String, Optional myRow As Integer, Optional myCol As Integer) As Range
'
''On Error Resume Next
'
''第一列
'Set rng = Intersect(rng.Columns(1), rng.Parent.UsedRange)
'If (IsNull(myRow)) Then myRow = 0
'If (IsNull(myCol)) Then myCol = 0
'
''1、先看有没有判断值，有就按判断值，没有就取第一个不为空不出错的
'If myCriteria = "" Then
'    For i = rng.Rows.Count To 1 Step -1
'        If rng(i, 1).value <> "" And Not IsError(rng(i, 1)) Then
'            Set ENDCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'Else
'    For i = rng.Rows.Count To 1 Step -1
'        If (rng(i, 1).value Like myCriteria) Then
'            Set ENDCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'End If
'
''2、如果最后都没有返回值（即有判断值，到没找到对应单元格），那就取第一个不为空不出错的
'If ENDCELL Is Nothing Then
'    For i = rng.Rows.Count To 1 Step -1
'        If rng(i, 1).value <> "" And Not IsError(rng(i, 1)) Then
'            Set ENDCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'End If
'
''3、实在什么都没有的话，就去第一格吧
'Set ENDCELL = rng(i + myRow, 1 + myCol)
'
'
'End Function
'
''在第一列找到第一个对应值，返回单元格对象
'Function FIRSTCELL(rng As Range, Optional myCriteria As String, Optional myRow As Integer, Optional myCol As Integer) As Range
'
''On Error Resume Next
'
''第一列
'Set rng = Intersect(rng.Columns(1), rng.Parent.UsedRange)
'If (IsNull(myRow)) Then myRow = 0
'If (IsNull(myCol)) Then myCol = 0
'
''1、先看有没有判断值，有就按判断值，没有就取第一个不为空不出错的
'If myCriteria = "" Then
'    For i = 1 To rng.Rows.Count
'        If rng(i, 1).value <> "" And Not IsError(rng(i, 1)) Then
'            Set FIRSTCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'Else
'    For i = 1 To rng.Rows.Count
'        If (rng(i, 1).value Like myCriteria) Then
'            Set FIRSTCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'End If
'
''2、如果最后都没有返回值（即有判断值，到没找到对应单元格），那就取第一个不为空不出错的
'If FIRSTCELL Is Nothing Then
'    For i = 1 To rng.Rows.Count
'        If rng(i, 1).value <> "" And Not IsError(rng(i, 1)) Then
'            Set FIRSTCELL = rng(i + myRow, 1 + myCol)
'            Exit Function
'        End If
'    Next
'End If
'
''3、实在什么都没有的话，就去第一格吧
'
'Set FIRSTCELL = rng(i + myRow, 1 + myCol)
'
'
'End Function
'
''取本单元格以上全部单元格
'Function UPRANGE() As Range
'    Application.Volatile '声明易失性
'    On Error GoTo Err:
'
'    Dim rng As Range
'
'    '取范围
'    Set rng = Application.ThisCell  '本单元格
'    Hight = rng.Row - 1             '高度
'    Set rng = rng.Offset(-Hight).Resize(Hight)  '偏移调整
'
'    '返回
'    Set UPRANGE = rng
'
'Err:
'    Set UPRANGE = rng
'
'End Function
'
'
''rng是判断列，creteria是判断值
''从函数所在单元格往下，取有值且不报错的范围作为高度，符合判断列中判断值的
'
'Function SUMDOWN(Optional rng As Range, Optional creteria As String) As Variant
'Application.Volatile '声明易失性
'On Error GoTo Err:
'
'Dim LimitsNum As Integer
'LimitsNum = 500
'
'' 如果没有指定判断列，就取单元格自身
'If rng Is Nothing Then Set rng = Application.ThisCell
''Debug.Print rng.Address
'
''高度大于500按大的，否则都调成500
'If rng.Rows.Count > LimitsNum Then
'    Set rng = rng.Resize(rng.Rows.Count, 1).Offset(Application.ThisCell.Row - rng.Row + 1, 0)
'Else
'    Set rng = rng.Resize(LimitsNum, 1).Offset(Application.ThisCell.Row - rng.Row + 1, 0)
'End If
'Set rng = Intersect(rng.Columns(1), rng.Parent.UsedRange)
'
'If creteria = "" Then
'    '1无判断值，取有值且不报错的部分确定高度
'    For i = 1 To rng.Rows.Count
'        If rng(i, 1).value = "" And Not IsError(rng(i, 1)) Then   '发现空就停
'            Set rng = Application.ThisCell.Offset(1, 0).Resize(i - 1, 1)
'            SUMDOWN = WorksheetFunction.Sum(rng)
'            Exit Function
'        End If
'    Next
'    '如果上面没有找到合适条件，则取区域高
'    Set rng = Application.ThisCell.Offset(1, 0).Resize(i - 1, 1)
'    SUMDOWN = WorksheetFunction.Sum(rng)
'
'Else
'    '2有判断值，则找判断值所在位置确定高度。
'    For i = 1 To rng.Rows.Count
'         If (rng(i, 1).value Like "*" & creteria & "*") Then
'            Set rng = Application.ThisCell.Offset(1, 0).Resize(i - 1, 1)
'            SUMDOWN = WorksheetFunction.Sum(rng)
'            Exit Function
'        End If
'    Next
'    '3有判断值，但找不到，确定不了高度，则按所选列往下，有值且不报错的部分确定高度（同前面）
'    For i = 1 To rng.Rows.Count
'        If rng(i, 1).value = "" And Not IsError(rng(i, 1)) Then
'            Set rng = Application.ThisCell.Offset(1, 0).Resize(i - 1, 1)
'            SUMDOWN = WorksheetFunction.Sum(rng)
'            Exit Function
'        End If
'    Next
'
'End If
'
'Err:
'
'SUMDOWN = 0
'
'End Function
'
'
'
'
'
'
'
'
''排除值相同值
''例EXCLUDE(A1:A10,A5:A10)
'
'Function EXCLUDE(ByVal OriginalArr As Variant, ByVal FunctionalArr As Variant) As Variant
'    Dim arr() As String
'    Dim i, j, k As Integer
'    k = 0
'
'    'Range对象都转成数组
''    If TypeName(OriginalArr) = "Range" Then OArr = WorksheetFunction.Transpose(OriginalArr)
''    If TypeName(FunctionalArr) = "Range" Then FArr = WorksheetFunction.Transpose(FunctionalArr)
'    OArr = WorksheetFunction.Transpose(OriginalArr)
'    FArr = WorksheetFunction.Transpose(FunctionalArr)
'
'    '所有原数组中的相同值变为0
'    For i = 1 To UBound(OArr)
'        For j = 1 To UBound(FArr)
'            If OArr(i) = FArr(j) And FArr(j) <> 0 And FArr(j) <> "" Then
'                OArr(i) = 0
'                k = k + 1   '记录排除的相同值数量
'            End If
'        Next
'    Next
'
'    '得到排除相同值的新数组Arr
'    ReDim arr(1 To (UBound(OArr) - k)) '重新定义新数组大小为原数量-排除的数量
'    j = 1   'j用来索引新数组
'    For i = 1 To UBound(OArr)
'        '排除空值和0值
'        If OArr(i) <> "" And OArr(i) <> 0 Then
'            arr(j) = OArr(i)
'            j = j + 1
'        End If
'    Next
'
'    EXCLUDE = arr
'
'End Function
'
''rng上全部单元格按CountIf大于0的计数
'Function UpCount() As Integer
'    Application.Volatile '声明易失性
'    Dim UpSection As Range
'
'    '选择从本单元格往上全部列
'    Set MeCell = Application.ThisCell
'    Hight = MeCell.Row - 1
'    Set UpSection = MeCell.Offset(-Hight).Resize(Hight) '移到最顶部取高度
'
'    UpCount = WorksheetFunction.CountIf(UpSection, ">0")
'
'End Function
'
'
'
'
''返回自身单元格下部区域
'Function RANGEDOWN(Optional rng As Range, Optional cond As String) As Range
'Application.Volatile '声明易失性
'On Error Resume Next
'
''1、确认区域
'If rng Is Nothing Then  '如果rng不存在
'    Set rng = Application.ThisCell.Offset(1, 0)
'Else
'    Set rng = Intersect(Application.ThisCell.EntireRow, rng.EntireColumn).Offset(1, 0)
'End If
'
''2、确认高度
'Dim LimitsNum As Integer
'Dim BelongHere As Boolean   '在范围内
'LimitsNum = 500             '向下查找的上限
'i = 0
'Do
'    i = i + 1
'
'    If cond = "" Then       '如果cond未定义
'        BelongHere = rng(i, 1).value <> ""
'    Else
'        BelongHere = rng(i, 1).value <> "" And rng(i, 1).value Like "*" & cond & "*"
'    End If
'Loop While BelongHere
'
''3、根据确认的i调整
'Set RANGEDOWN = rng.Resize(i - 1, 1)
''Debug.Print rng.Address
'
'End Function
'
''正则Like
''例RLike("[GS]", rng)为True
''例RLike("[GS]", "GS2775[2775*1050,-650]")为True
'Function RLIKE(ByVal pt As String, cond) As Boolean
'    Dim Reg As Object   '正则表达式用于判断工程量标题QtyRange
'    Set Reg = CreateObject("VBScript.Regexp")
'    Dim str As String
'
'    If TypeName(cond) = "Range" Then
'        str = cond.value
'    Else
'        str = cond
'    End If
'
'    Reg.Pattern = pt      '正则判断值
'    RLIKE = Reg.test(str)
'
'End Function
'
''
'
''获取最后一个字符串
'
''举例 B7有值等于"Long Yang-Lyod2019/011 25/3/2019 Maintenance work(labour cost) 25/1-15/3/2019 No.1 FC P26-3 GL-2019-019-FCP26-LW"
''ENDSTRING(B7,"/")得到 2019 No.1 FC P26-3 GL-2019-019-FCP26-LW"
'
'Function ENDSTRING(rng As Range, V As String, Optional k As Integer) As String
'
'    Dim str, TotalArr() As String
'
'    str = rng(1, 1).value
'    TotalArr = Split(str, V)
'
'    '最后一个值。K调整位置
''    Debug.Print TotalArr(UBound(TotalArr))
'
'    ENDSTRING = TotalArr(UBound(TotalArr) + k)
'
'End Function
'
'
'
''获取第一个字符串
'
''举例 B7有值等于"Long Yang-Lyod2019/011 25/3/2019 Maintenance work(labour cost) 25/1-15/3/2019 No.1 FC P26-3 GL-2019-019-FCP26-LW"
''STARTSTRING(B7,"/")得到 2019 No.1 FC P26-3 GL-2019-019-FCP26-LW"
'
'Function STARTSTRING(rng As Range, V As String, Optional k As Integer) As String
'
'    Dim str, TotalArr() As String
'
'    str = rng(1, 1).value
'    TotalArr = Split(str, V)
'
'    '最后一个值。K调整位置
''    Debug.Print TotalArr(UBound(TotalArr))
'
'    STARTSTRING = TotalArr(0 + k)
'
'End Function
'

'
'Function toString(rng As Range) As String
'
'    Application.Volatile '声明易失性
'    On Error Resume Next
'
'    toString = CStr(rng.value)
'
'End Function

