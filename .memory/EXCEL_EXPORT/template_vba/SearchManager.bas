Attribute VB_Name = "SearchManager"
' 搜索管理模块 - 处理ListBox搜索相关功能
' 将搜索逻辑从窗体中分离，提高代码可维护性

Option Explicit

' 搜索结果管理
Private m_SearchResults() As Long     ' 存储匹配的索引
Private m_CurrentSearchIndex As Long  ' 当前搜索结果中的位置
Private m_SearchResultCount As Long   ' 搜索结果总数
Private m_ActiveListBox As Object     ' 当前操作的ListBox
Private m_StatusLabel As Object       ' 显示搜索状态的Label

' ====================================================================
' 公共接口方法
' ====================================================================

' 初始化搜索管理器
Public Sub InitializeSearchManager(ByRef listBox As Object, ByRef statusLabel As Object)
    Set m_ActiveListBox = listBox
    Set m_StatusLabel = statusLabel
    Call ClearSearchResults
End Sub

' 执行新搜索（收集所有匹配项）
Public Function PerformNewSearch(ByVal searchText As String) As Long
    If Trim(searchText) = "" Then
        PerformNewSearch = -1  ' 返回-1表示搜索文本为空
        Exit Function
    End If
    
    If m_ActiveListBox Is Nothing Then
        PerformNewSearch = -2  ' 返回-2表示ListBox未设置
        Exit Function
    End If
    
    Dim upperSearchText As String
    upperSearchText = UCase(Trim(searchText))  ' 转为大写进行不区分大小写搜索
    
    ' 初始化搜索结果 - 安全的数组初始化
    If m_ActiveListBox.ListCount > 0 Then
        ReDim m_SearchResults(0 To m_ActiveListBox.ListCount - 1)
    Else
        ReDim m_SearchResults(0 To 0)
    End If
    m_SearchResultCount = 0
    m_CurrentSearchIndex = 0
    
    Dim i As Long
    
    ' 在ListBox中搜索所有匹配项
    For i = 0 To m_ActiveListBox.ListCount - 1
        ' 在编号或名称中搜索
        If InStr(1, UCase(m_ActiveListBox.List(i, 0)), upperSearchText) > 0 Or _
           InStr(1, UCase(m_ActiveListBox.List(i, 1)), upperSearchText) > 0 Then
            m_SearchResults(m_SearchResultCount) = i
            m_SearchResultCount = m_SearchResultCount + 1
        End If
    Next i
    
    ' 如果找到匹配项
    If m_SearchResultCount > 0 Then
        ' 重新调整数组大小
        ReDim Preserve m_SearchResults(0 To m_SearchResultCount - 1)
        
        ' 定位到第一个匹配项
        m_CurrentSearchIndex = 0
        Call GotoCurrentResult
        
        Debug.Print "搜索成功：找到 " & m_SearchResultCount & " 个匹配项"
        PerformNewSearch = m_SearchResultCount  ' 返回找到的结果数量
    Else
        ' 没有找到匹配项
        Call ClearSearchResults
        Debug.Print "搜索失败：未找到匹配项"
        PerformNewSearch = 0  ' 返回0表示未找到
    End If
End Function

' 跳转到上一条搜索结果
Public Function GotoPreviousResult() As Boolean
    If m_SearchResultCount = 0 Then
        GotoPreviousResult = False
        Exit Function
    End If
    
    If m_CurrentSearchIndex > 0 Then
        m_CurrentSearchIndex = m_CurrentSearchIndex - 1
    Else
        ' 循环到最后一条
        m_CurrentSearchIndex = m_SearchResultCount - 1
    End If
    
    Call GotoCurrentResult
    GotoPreviousResult = True
End Function

' 跳转到下一条搜索结果
Public Function GotoNextResult() As Boolean
    If m_SearchResultCount = 0 Then
        GotoNextResult = False
        Exit Function
    End If
    
    If m_CurrentSearchIndex < m_SearchResultCount - 1 Then
        m_CurrentSearchIndex = m_CurrentSearchIndex + 1
    Else
        ' 循环到第一条
        m_CurrentSearchIndex = 0
    End If
    
    Call GotoCurrentResult
    GotoNextResult = True
End Function

' 清除搜索结果
Public Sub ClearSearchResults()
    m_SearchResultCount = 0
    m_CurrentSearchIndex = 0
    If Not m_StatusLabel Is Nothing Then
        m_StatusLabel.Caption = "0/0"
    End If
    Debug.Print "搜索结果已清除"
End Sub

' 获取搜索统计信息
Public Function GetSearchStatus() As String
    If m_SearchResultCount > 0 Then
        GetSearchStatus = (m_CurrentSearchIndex + 1) & "/" & m_SearchResultCount
    Else
        GetSearchStatus = "0/0"
    End If
End Function

' 获取当前搜索结果数量
Public Function GetSearchResultCount() As Long
    GetSearchResultCount = m_SearchResultCount
End Function

' 获取当前索引
Public Function GetCurrentSearchIndex() As Long
    GetCurrentSearchIndex = m_CurrentSearchIndex
End Function

' ====================================================================
' 私有辅助方法
' ====================================================================

' 跳转到当前搜索结果
Private Sub GotoCurrentResult()
    If m_SearchResultCount > 0 And m_CurrentSearchIndex >= 0 And m_CurrentSearchIndex < m_SearchResultCount Then
        Dim targetIndex As Long
        targetIndex = m_SearchResults(m_CurrentSearchIndex)
        
        ' 设置ListBox选中项
        m_ActiveListBox.ListIndex = targetIndex
        
        ' 确保项目可见（调用窗体的方法）
        If TypeName(m_ActiveListBox.Parent) = "UserForm" Then
            On Error Resume Next
            m_ActiveListBox.Parent.EnsureItemVisibleInForm targetIndex
            On Error GoTo 0
        End If
        
        ' 更新状态显示
        If Not m_StatusLabel Is Nothing Then
            m_StatusLabel.Caption = (m_CurrentSearchIndex + 1) & "/" & m_SearchResultCount
        End If
        
        Debug.Print "跳转到搜索结果 " & (m_CurrentSearchIndex + 1) & "/" & m_SearchResultCount & " (行" & (targetIndex + 1) & ")"
    End If
End Sub

' 清理资源
Public Sub CleanupSearchManager()
    Set m_ActiveListBox = Nothing
    Set m_StatusLabel = Nothing
    m_SearchResultCount = 0
    m_CurrentSearchIndex = 0
End Sub

