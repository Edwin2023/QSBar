VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 全局变量用于存储当前双击的单元格
Public gTargetCell As Range

' 缓存相关变量
Private m_CachedData As Variant  ' 缓存的数据数组
Private m_CacheValid As Boolean  ' 缓存是否有效
Private m_LastWorksheetName As String  ' 上次缓存的工作表名称

' 双击事件：在A列双击非空单元格时触发
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 1 Then
        If Not IsEmpty(Target) Then
            On Error GoTo ErrorHandler
            
            Debug.Print "双击单元格: " & Target.Address & ", 值: " & Target.Value
            
            Dim ws As Worksheet
            Dim uf As rateLibrary
                  
            Set ws = GetSheet("RATE")
            Set uf = New rateLibrary
            
  ' =========== 【在这里插入代码】 ===========
            ' 修复 BUG：强制刷新缓存，确保每次都从 Excel 表格读取最新数据
            Call RefreshCache
            
            ' 智能加载数据：使用缓存或重新加载
            Call LoadAllDataWithCache(uf.ListBox1, ws)
            
            ' 查找并选择指定值
            Dim foundIndex As Long
            foundIndex = FindItemInListBox(uf.ListBox1, Target.Value)
            
            If foundIndex >= 0 Then
                uf.ListBox1.ListIndex = foundIndex
                ' 选中项居中显示
                Call EnsureItemCentered(uf, foundIndex)
                Debug.Print "找到并选择了: " & Target.Value & " (索引: " & foundIndex & ")，已居中显示"
            Else
                Debug.Print "未找到: " & Target.Value & "，选择第一项"
                If uf.ListBox1.ListCount > 0 Then
                    uf.ListBox1.ListIndex = 0
                    Call EnsureItemCentered(uf, 0)
                End If
            End If
            
            Set gTargetCell = Target
            Cancel = True
            uf.Show vbModal
            
            Debug.Print "窗体操作完成"
        End If
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "双击事件错误: " & Err.Description & " (错误号: " & Err.Number & ")"
    If Err.Number = 9 Then
        MsgBox "RATE表不存在！", vbCritical
    Else
        MsgBox "发生错误: " & Err.Description, vbCritical, "错误"
    End If
End Sub

' 通过CodeName获取工作表对象
Private Function GetSheet(cName As String) As Worksheet
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Sheets
        If sht.CodeName = cName Then
            Set GetSheet = sht
            Exit Function
        End If
    Next sht
    Err.Raise 9 ' 触发错误号为9的错误（下标越界）
End Function

' 智能加载所有数据（使用缓存机制）
Private Sub LoadAllDataWithCache(listBox As Object, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' 检查是否需要重新加载缓存
    If Not m_CacheValid Or m_LastWorksheetName <> ws.Name Then
        Debug.Print "=== 重新加载数据缓存 ==="
        
        ' 重新构建缓存
        Call BuildDataCache(ws)
        m_LastWorksheetName = ws.Name
        m_CacheValid = True
        
        Debug.Print "缓存构建完成，共" & UBound(m_CachedData, 1) + 1 & "行数据"
    Else
        Debug.Print "=== 使用现有缓存 ==="
    End If
    
    ' 从缓存加载数据到ListBox
    Call LoadCacheToListBox(listBox)
    
    Debug.Print "ListBox加载完成，项目数: " & listBox.ListCount
    Exit Sub
    
ErrorHandler:
    Debug.Print "数据加载失败: " & Err.Description & " (错误号: " & Err.Number & ")"
    m_CacheValid = False  ' 缓存失效
End Sub

' 构建数据缓存
Private Sub BuildDataCache(ws As Worksheet)
    Dim r As Long
    Dim lastRow As Long
    Dim rowData As Variant
    Dim cacheIndex As Long
    Dim tempCache As Collection  ' 使用集合临时存储，避免数组越界
    
    ' 获取数据范围
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' 创建临时集合
    Set tempCache = New Collection
    
    ' 从第2行开始读取所有数据（跳过表头）
    For r = 2 To lastRow
        On Error Resume Next  ' 避免单行错误影响整体加载
        rowData = GetRowValues(ws, r)
        
        ' 只缓存非空行且编号不为空的行
        If Not IsZeroRow(rowData) And Trim(CStr(rowData(0))) <> "" And Trim(CStr(rowData(0))) <> "编号" Then
            ' 将行数据添加到临时集合
            tempCache.Add Array(rowData(0), rowData(1), rowData(2), rowData(3))
            
            ' 缓存项目统计（已清理详细输出）
        End If
        On Error GoTo 0
    Next r
    
    Debug.Print "有效数据行数: " & tempCache.Count
    
    ' 将集合数据转换为数组
    If tempCache.Count > 0 Then
        ReDim m_CachedData(0 To tempCache.Count - 1, 0 To 3)
        
        Dim i As Long
        Dim item As Variant
        For i = 1 To tempCache.Count
            item = tempCache(i)
            m_CachedData(i - 1, 0) = item(0)  ' 编号
            m_CachedData(i - 1, 1) = item(1)  ' 名称
            m_CachedData(i - 1, 2) = item(2)  ' 单位
            m_CachedData(i - 1, 3) = item(3)  ' 单价
        Next i
    Else
        ' 创建空缓存
        ReDim m_CachedData(0 To 0, 0 To 3)
        m_CachedData(0, 0) = ""
        m_CachedData(0, 1) = ""
        m_CachedData(0, 2) = ""
        m_CachedData(0, 3) = ""
        Debug.Print "警告：未找到有效数据"
    End If
End Sub

' 从缓存加载数据到ListBox（高性能版本）
Private Sub LoadCacheToListBox(listBox As Object)
    Dim i As Long, validCount As Long
    
    ' 清空ListBox
    listBox.Clear
    
    ' 检查缓存是否为空
    If Not IsArray(m_CachedData) Then
        Debug.Print "缓存数组未初始化"
        Exit Sub
    End If
    
    ' 检查缓存是否有有效数据
    If UBound(m_CachedData, 1) = 0 And m_CachedData(0, 0) = "" Then
        Debug.Print "缓存为空，没有有效数据"
        Exit Sub
    End If
    
    ' 先统计有效行数
    validCount = 0
    For i = 0 To UBound(m_CachedData, 1)
        If Trim(CStr(m_CachedData(i, 0))) <> "" Then
            validCount = validCount + 1
        End If
    Next i
    
    If validCount = 0 Then
        Debug.Print "没有有效数据行"
        Exit Sub
    End If
    
    ' 创建临时数组用于一次性赋值
    Dim tempArray() As String
    ReDim tempArray(0 To validCount - 1, 0 To 3)
    
    ' 填充临时数组
    Dim arrayIndex As Long
    arrayIndex = 0
    For i = 0 To UBound(m_CachedData, 1)
        If Trim(CStr(m_CachedData(i, 0))) <> "" Then
            tempArray(arrayIndex, 0) = CStr(m_CachedData(i, 0))  ' 编号
            tempArray(arrayIndex, 1) = CStr(m_CachedData(i, 1))  ' 名称
            tempArray(arrayIndex, 2) = CStr(m_CachedData(i, 2))  ' 单位
            
            ' 安全处理单价格式化
            On Error Resume Next
            If IsNumeric(m_CachedData(i, 3)) Then
                tempArray(arrayIndex, 3) = Format(m_CachedData(i, 3), "#,##0.00")
            Else
                tempArray(arrayIndex, 3) = CStr(m_CachedData(i, 3))
            End If
            On Error GoTo 0
            
            arrayIndex = arrayIndex + 1
        End If
    Next i
    
    ' 一次性赋值给ListBox（高性能）
    listBox.List = tempArray
    
    ' 加载完成
End Sub

' 在ListBox中查找指定值
Private Function FindItemInListBox(listBox As Object, searchValue As String) As Long
    Dim i As Long
    
    For i = 0 To listBox.ListCount - 1
        If CStr(listBox.List(i, 0)) = CStr(searchValue) Then
            FindItemInListBox = i
            Exit Function
        End If
    Next i
    
    FindItemInListBox = -1  ' 未找到
End Function

' 确保指定项在ListBox中居中显示
Private Sub EnsureItemCentered(uf As rateLibrary, itemIndex As Long)
    On Error Resume Next
    
    If itemIndex < 0 Or itemIndex >= uf.ListBox1.ListCount Then Exit Sub
    
    ' 调用窗体的居中显示方法
    uf.EnsureItemVisibleInForm itemIndex
    
    Debug.Print "项目已居中显示: 索引 " & itemIndex
End Sub

' 获取指定行的1,2,3,10列数据（自动格式化数值）
Private Function GetRowValues(ws As Worksheet, rowIndex As Long) As Variant
    Dim result(0 To 3) As Variant
    Dim columns(0 To 3) As Long
    
    ' 定义要获取的列号
    columns(0) = 1  ' 编号
    columns(1) = 2  ' 名称
    columns(2) = 3  ' 单位
    columns(3) = 10  ' 单价 (J列)
    
    Dim i As Long
    Dim cellValue As Variant
    
    For i = 0 To 3
        On Error Resume Next
        cellValue = ws.Cells(rowIndex, columns(i)).Value
        On Error GoTo 0
        
        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            result(i) = ""
        ElseIf IsNumeric(cellValue) And i = 3 Then  ' 只对单价列（第7列）进行格式化
            result(i) = CDbl(cellValue)  ' 保持数值类型，在显示时再格式化
        Else
            result(i) = CStr(cellValue)  ' 转换为字符串
        End If
    Next i
    
    ' 调试信息已清理
    
    GetRowValues = result
End Function

' 判断行是否为空行（前三列都为空/0时返回True）
Private Function IsZeroRow(rowData As Variant) As Boolean
    If Not IsArray(rowData) Then
        IsZeroRow = False
        Exit Function
    End If
    
    Dim j As Integer
    ' 检查前三列（编号、名称、单位）
    For j = 0 To 2
        If IsEmpty(rowData(j)) Or Trim(CStr(rowData(j))) = "" Or Trim(CStr(rowData(j))) = "0" Then
            ' 继续检查下一列
        Else
            ' 如果有一列不为空，说明不是空行
            IsZeroRow = False
            Exit Function
        End If
    Next j
    ' 如果前三列都为空，返回True
    IsZeroRow = True
End Function

' 强制刷新缓存
Public Sub RefreshCache()
    Debug.Print "=== 强制刷新数据缓存 ==="
    m_CacheValid = False
    m_LastWorksheetName = ""
    Debug.Print "缓存已清除，下次使用时将重新加载"
End Sub

' 显示缓存状态
Public Sub ShowCacheStatus()
    Debug.Print "=== 缓存状态信息 ==="
    Debug.Print "缓存有效: " & m_CacheValid
    Debug.Print "工作表名称: " & m_LastWorksheetName
    
    If IsArray(m_CachedData) Then
        Debug.Print "缓存数据行数: " & UBound(m_CachedData, 1) + 1
        If UBound(m_CachedData, 1) >= 0 Then
            Debug.Print "第一项: " & m_CachedData(0, 0) & " - " & m_CachedData(0, 1)
        End If
    Else
        Debug.Print "缓存数据: 未初始化"
    End If
    Debug.Print "=== 状态信息完成 ==="
End Sub

' 简化的测试方法
Public Sub TestRate(Optional ByVal testValue As String = "B.71")
    On Error GoTo ErrorHandler
    
    Debug.Print "=== 测试RATE窗体 ==="
    Debug.Print "测试值: " & testValue
    
    ' 模拟双击事件
    Dim testCell As Range
    Set testCell = Application.ActiveSheet.Range("A1")
    testCell.Value = testValue
    
    ' 触发双击事件
    Call Worksheet_BeforeDoubleClick(testCell, False)
    
    Debug.Print "=== 测试完成 ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "测试失败: " & Err.Description & " (错误号: " & Err.Number & ")"
    MsgBox "测试失败: " & Err.Description, vbCritical, "错误"
End Sub





