Attribute VB_Name = "SearchManager"
' 搜索功能模块 - 管理ListBox搜索相关功能
' 将搜索逻辑从窗体中分离，提高代码可维护性

Option Explicit

' 搜索结果相关
Private m_SearchResults() As Long     ' 存储匹配结果的行索引
Private m_CurrentSearchIndex As Long  ' 当前在搜索结果中的位置
Private m_SearchResultCount As Long   ' 搜索结果数量
Private m_ActiveListBox As Object     ' 当前操作的ListBox
Private m_StatusLabel As Object       ' 显示搜索状态的Label

' ====================================================================
' 公开接口方法
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
        PerformNewSearch = -1
        Exit Function
    End If

    If m_ActiveListBox Is Nothing Then
        PerformNewSearch = -2
        Exit Function
    End If

    Dim upperSearchText As String
    upperSearchText = UCase(Trim(searchText))

    ' 初始化结果数组
    If m_ActiveListBox.ListCount > 0 Then
        ReDim m_SearchResults(0 To m_ActiveListBox.ListCount - 1)
    Else
        ReDim m_SearchResults(0 To 0)
    End If
    m_SearchResultCount = 0
    m_CurrentSearchIndex = 0

    Dim i As Long
    For i = 0 To m_ActiveListBox.ListCount - 1
        If InStr(1, UCase(m_ActiveListBox.List(i, 0)), upperSearchText) > 0 Or _
           InStr(1, UCase(m_ActiveListBox.List(i, 1)), upperSearchText) > 0 Then
            m_SearchResults(m_SearchResultCount) = i
            m_SearchResultCount = m_SearchResultCount + 1
        End If
    Next i

    If m_SearchResultCount > 0 Then
        ReDim Preserve m_SearchResults(0 To m_SearchResultCount - 1)
        m_CurrentSearchIndex = 0
        Call GotoCurrentResult
        PerformNewSearch = m_SearchResultCount
    Else
        Call ClearSearchResults
        PerformNewSearch = 0
    End If
End Function

' 转到上一个搜索结果
Public Function GotoPreviousResult() As Boolean
    If m_SearchResultCount = 0 Then
        GotoPreviousResult = False
        Exit Function
    End If

    If m_CurrentSearchIndex > 0 Then
        m_CurrentSearchIndex = m_CurrentSearchIndex - 1
    Else
        m_CurrentSearchIndex = m_SearchResultCount - 1
    End If

    Call GotoCurrentResult
    GotoPreviousResult = True
End Function

' 转到下一个搜索结果
Public Function GotoNextResult() As Boolean
    If m_SearchResultCount = 0 Then
        GotoNextResult = False
        Exit Function
    End If

    If m_CurrentSearchIndex < m_SearchResultCount - 1 Then
        m_CurrentSearchIndex = m_CurrentSearchIndex + 1
    Else
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
End Sub

' 获取搜索状态
Public Function GetSearchResultCount() As Long
    GetSearchResultCount = m_SearchResultCount
End Function

' 清理资源
Public Sub CleanupSearchManager()
    Set m_ActiveListBox = Nothing
    Set m_StatusLabel = Nothing
    m_SearchResultCount = 0
    m_CurrentSearchIndex = 0
End Sub

' ====================================================================
' 显示价格库窗体（被 RateDoubleClickHandler 调用）
' ====================================================================

Public Sub ShowRateLibrary(ByVal sourceSheet As Object, ByVal TargetCell As Range, Optional ByVal initialValue As String = "")
    On Error GoTo ErrorHandler

    Dim uf As rateLibrary
    Set uf = New rateLibrary

    ' 如果提供初始值，查找并定位
    If initialValue <> "" Then
        Dim i As Long
        For i = 0 To uf.ListBox1.ListCount - 1
            If Trim(CStr(uf.ListBox1.List(i, 0))) = Trim(initialValue) Then
                uf.ListBox1.ListIndex = i
                uf.EnsureItemVisibleInForm i
                Exit For
            End If
        Next i
    End If

    uf.Show vbModeless
    Exit Sub

ErrorHandler:
    MsgBox "显示价格库窗体时发生错误: " & Err.Description, vbCritical, "错误"
End Sub

' ====================================================================
' 私有辅助方法
' ====================================================================

' 转到当前搜索结果
Private Sub GotoCurrentResult()
    If m_SearchResultCount > 0 And m_CurrentSearchIndex >= 0 And m_CurrentSearchIndex < m_SearchResultCount Then
        Dim targetIndex As Long
        targetIndex = m_SearchResults(m_CurrentSearchIndex)

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
    End If
End Sub
