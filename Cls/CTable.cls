VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим

Option Explicit

Private m_wsSheet        As Worksheet

Private m_rngUpperBound  As Range
Private m_rngLowerBound  As Range

'Инициализация класса
Public Sub InitClass(wsSheet As Worksheet, _
                     Optional vNameRange As Variant)
                    
    Set m_wsSheet = wsSheet
    
    With m_wsSheet

        If TypeName(vNameRange) = "String" Then
            
            On Error Resume Next
            
            Set m_rngUpperBound = .Range(vNameRange)
            
            
        ElseIf TypeName(vNameRange) = "Range" Then
            Set m_rngUpperBound = vNameRange
        End If
    
    End With
      
End Sub

'Информация таблицы
Public Sub ToString()
    
    If Not IsAccesTable Then
        
        Debug.Print "-----------------------------------------------------------------------------------------"
        Debug.Print "Table not found"
        
    Else
    
        Debug.Print "-----------------------------------------------------------------------------------------"
        Debug.Print "Sheet name: " & m_wsSheet.Name & Chr(10) & _
                    "First_row: " & GetFirstRow & "; Last_row: " & GetLastRow & Chr(10) & _
                    "First_column: "; GetFirstColumn & "; Last_column: " & GetLastColumn & Chr(10) & _
                    "Range: " & GetRangeTable.Address(False, False)
                    
    End If

End Sub

'Снятие фильтра
Public Sub DisableFilters()
                
    With m_wsSheet
        If .FilterMode Then  'Если нет фильтра
            .ShowAllData
        End If
    End With
    
End Sub

'Получаем диапозон таблицы, начало и конца
Public Property Get GetRangeTable(Optional ByVal isStandartRow As Boolean = True, _
                                  Optional ByVal isStandartCol As Boolean = True) As Range

    If Not IsAccesTable Then
        GetRangeTable = Nothing
    End If

    With m_wsSheet
        Set GetRangeTable = .Range( _
                                    .Cells(GetFirstRow + 1, GetFirstColumn), _
                                    .Cells(GetLastRow, GetLastColumn) _
                                  )
    End With

End Property

'Первая строка
Public Property Get GetFirstRow() As Long

    GetFirstRow = m_rngUpperBound.Row
    
End Property

'Первый столбец
Public Property Get GetFirstColumn() As Long
    
    GetFirstColumn = m_rngUpperBound.Column
    
End Property

'Получаем диапозон верхниций границы
Public Property Get GetRangeUpperBound() As Range

    Set GetRangeUpperBound = m_rngUpperBound
    
End Property

'Изменяем верхнию границу c помощью диапозона
Public Sub SetRangeUpperBound(vBound As Variant)
    
    If TypeName(vBound) = "String" Then
        Set m_rngUpperBound = m_wsSheet.Range(vBound)
    ElseIf TypeName(vBound) = "Range" Then
        Set m_rngUpperBound = vBound
    End If

End Sub

'Изменяем верхнию границу c помощью координат
Public Sub SetCellUpperBound(ByVal lRow As Long, _
                             ByVal lColumn As Long)
    
    Set m_rngUpperBound = m_wsSheet.Cells(lRow, lColumn)
    
End Sub

'Последняя строка
Public Property Get GetLastRow(Optional lColumn As Long = 0) As Long
    
    If lColumn = 0 Then
        lColumn = GetFirstColumn
    End If
    
    If m_rngLowerBound Is Nothing Then
        GetLastRow = m_wsSheet.Cells(m_wsSheet.Rows.Count, lColumn).End(xlUp).Row
    Else
        GetLastRow = m_rngLowerBound.Row
    End If

End Property

'Последний столбец
Public Property Get GetLastColumn(Optional lRow As Long = 0) As Long
    
    If lRow = 0 Then
        lRow = GetFirstRow
    End If
    
    If m_rngLowerBound Is Nothing Then
        GetLastColumn = m_wsSheet.Cells(lRow, m_wsSheet.Columns.Count).End(xlToLeft).Column
    Else
        GetLastColumn = m_rngLowerBound.Column
    End If

End Property

'Получаем диапозон верхниций границы
Public Property Get GetRangeLowerBound() As Range

    Set GetRangeUpperBound = m_rngLowerBound
    
End Property

'Изменяем верхнию границу c помощью диапозона
Public Sub SetRangeLowerBound(vBound As Variant)
    
    If TypeName(vBound) = "String" Then
        Set m_rngLowerBound = m_wsSheet.Range(vBound)
    ElseIf TypeName(vBound) = "Range" Then
        Set m_rngLowerBound = vBound
    End If
   
End Sub

'Изменяем верхнию границу c помощью координат
Public Sub SetCellLowerBound(ByVal lRow As Long, _
                             ByVal lColumn As Long)
    
    Set m_rngLowerBound = m_wsSheet.Cells(lRow, lColumn)
    
End Sub

'Ищем строку по метке, если не задан параметр After: = по умолчанию берем верхнию границу
Public Property Get GetFindRow(Optional ByVal lColumn As Long = 0, _
                               Optional ByVal sFind As String = vbNullString, _
                               Optional rngAfter As Range = Nothing) As Long
    If IsAccesTable Then
        
        If lColumn = 0 Then
            lColumn = GetFirstColumn
        End If
        
        On Error Resume Next
        
        If rngAfter Is Nothing Then
            GetFindRow = m_wsSheet.Columns(lColumn).Find(What:=sFind, SearchOrder:=xlByRows).Row
        Else
            GetFindRow = m_wsSheet.Columns(lColumn).Find(What:=sFind, SearchOrder:=xlByRows, After:=rngAfter).Row
        End If

        Err.Clear
        On Error GoTo 0
        
    End If
    
End Property

'Ищем столбец по метке, если не задан параметр After: = по умолчанию берем верхнию границу
Public Property Get GetFindColumn(Optional ByVal lRow As Long = 0, _
                                  Optional ByVal sFind As String = vbNullString, _
                                  Optional rngAfter As Range = Nothing) As Long
    
    If IsAccesTable Then
        
        If lRow = 0 Then
            lRow = GetFirstRow
        End If
        
        On Error Resume Next
        
        If rngAfter Is Nothing Then
            GetFindColumn = m_wsSheet.Rows(lRow).Find(What:=sFind, SearchOrder:=xlByColumns).Column
        Else
            GetFindColumn = m_wsSheet.Rows(lRow).Find(What:=sFind, SearchOrder:=xlByColumns, After:=rngAfter).Column
        End If
        
    End If
     
End Property

'Сбрасываем последнию строку
Public Sub ClearLowerBound()
    
    Set m_rngLowerBound = Nothing
    
End Sub

'Ищем диапозон таблицы по именованию столбца
Public Sub FindRangeInTable(Optional ByVal sNameColumn As String = vbNullString)
    Dim rngTable    As Range
    
    If LenB(sNameColumn) = 0 Then
        Exit Sub
    End If
    
    With m_wsSheet
        'Если необходимо найти диапозон
        Set rngTable = .UsedRange.Find(sNameColumn, LookAt:=xlWhole)
        If Not rngTable Is Nothing Then
            Set m_rngUpperBound = rngTable
        End If
    End With
    
End Sub

'Получаем список колонок
Public Property Get GetColumns(Optional ByVal lRowLables As Long = 0) As Object

    Dim oColumns    As Object
    Dim i           As Long
    
    Set oColumns = CreateObject("Scripting.Dictionary")
    
    If lRowLables = 0 Then 'Если строка не задана, то рассматриваем метки по первой строке
        lRowLables = GetFirstRow
    End If

    With m_wsSheet

        For i = GetFirstColumn To GetLastColumn
            With .Cells(lRowLables, i)
                If LenB(.Value) > 0 And Not oColumns.exists(.Value) Then
                    oColumns.Add .Value, i
                End If
            End With
        Next i
        
    End With
    
    Set GetColumns = oColumns
    
End Property

' наименование колонки c помощью формуляра
Public Sub UpdateKeyColumn(ByRef oColumns As Object, _
                           ByVal sName As String, _
                           ByVal sNewName As String)
                           
    If oColumns.exists(sName) Then
        oColumns.Key(sName) = sNewName
    End If
    
End Sub

'Меняем наименование стандартных колонок c помощью формуляров
Public Sub UpdateKeysColumns(ByRef oColumns As Object, _
                                   oFormular As Object)
    Dim vKey As Variant

    For Each vKey In oFormular
        If oColumns.exists(vKey) Then
            oColumns.Key(vKey) = oFormular(vKey)
        End If
    Next vKey
    
End Sub

'Определяем таблицу по верхнему флагу
Public Property Get IsAccesTable() As Boolean
    
    IsAccesTable = Not m_rngUpperBound Is Nothing
    
End Property

'Если пустая таблица
Public Property Get IsEmptyTable() As Boolean
    
    If IsAccesTable Then
        IsEmptyTable = Not GetFirstRow < GetLastRow
    End If

End Property

'Если диапозон находится в таблицы
Public Property Get IsRangeInTableExists(rngValue As Range, _
                                         Optional rngTable As Range = Nothing) As Boolean
    
    If Not IsEmptyTable Then
    
        If rngTable Is Nothing Then
            Set rngTable = GetRangeTable
        End If

        IsRangeInTableExists = Not Application.Intersect(rngValue, rngTable) Is Nothing
        
    End If
    
End Property
