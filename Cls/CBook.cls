VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим

Option Explicit

Private m_wbBook    As Workbook
Private m_isSource  As Boolean

'Инициализация класса
Private Sub Class_Initialize()

    Set m_wbBook = ThisWorkbook
    
End Sub

'Инициализация класса
Public Sub SetSourceWorkbook(wbBook As Workbook)

    m_wbBook = wbBook
    m_isSource = True

End Sub

'Стиль ссылок
Public Sub SetApplicationWorkbook(Optional ByVal isEnable As Boolean = True, _
                               Optional ByVal isCalculate As Boolean = True)

    With Application
        .DisplayAlerts = isEnable
        .ScreenUpdating = isEnable
        .EnableEvents = isEnable
        .Calculation = IIf(isCalculate, xlCalculationAutomatic, xlCalculationManual)
    End With

End Sub

'Ставим стиль ссылок буквенный
Public Sub SetApplicationReference()
    
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        End If
    End With

End Sub

'Удаляем листы
Public Sub DeleteSheets(vNames As Variant)
    Dim vKey    As Variant

    Application.DisplayAlerts = False
    
    If IsArray(vNames) Then
        For Each vKey In vNames
            DeleteSheets vKey
        Next vKey
    Else
        If IsSheetExists(vNames) Then
            m_wbBook.Sheets(vNames).Delete
        End If
    End If
    
    Application.DisplayAlerts = True

End Sub

'Включение общего доступа
Public Sub EnableGeneralAccess(Optional ByVal IsKeepChange As Boolean = False, _
                               Optional ByVal iHistoryDuration As Integer = 30)

    'Если общий доступ выходим
    If IsMultiUser Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
        
    With m_wbBook
        .SaveAs Filename:=wbBook.FullName, accessMode:=xlShared 'Паратер фалй
        .KeepChangeHistory = IsKeepChange
        .ChangeHistoryDuration = iHistoryDuration
    End With
    
    Application.DisplayAlerts = True
    
End Sub

'Отключение общего доступа
Public Sub DisableGeneralAccess()

    'Если общий доступ выходим
    If Not IsMultiUser Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    m_wbBook.ExclusiveAccess
    
    Application.DisplayAlerts = True
    
End Sub

'Создать макетный лист
Public Property Get GetCopySheet(ByVal sNameMaket As String, _
                                 ByVal sNewName As String, _
                                 Optional ByVal isAfterCopy As Boolean = False) As Worksheet
    
    Dim wsMaket         As Worksheet, wsNewSheet    As Worksheet
    Dim isVisibleBook   As Boolean, isVisibleSheet  As Boolean
    
    'Если книга скрыта, то раскрываем
    If m_isSource And Not m_wbBook.Windows(1).Visible Then
        m_wbBook.Windows(1).Visible = True
        isVisibleBook = True
    End If
    
    Set wsMaket = m_wbBook.Sheets(sNameMaket)
    
    With wsMaket
    
        'Если лист скрыт

        If .Visible = xlSheetVeryHidden Then
            .Visible = xlSheetVisible
            isVisibleSheet = True
        End If

        DeleteSheets sNewName
        
        .Copy After:=m_wbBook.Sheets(IIf(isAfterCopy, sNameMaket, m_wbBook.Sheets.Count))   'Копируем лист
        
        Set wsNewSheet = ActiveSheet
        wsNewSheet.Name = sNewName
        
        .Visible = IIf(isVisibleSheet, xlSheetVeryHidden, xlSheetVeryHidden)
        
        If isVisibleBook Then
            m_wbBook.Windows(1).Visible = Not isVisibleBook
        End If

        Set GetCopySheet = wsNewSheet
        
    End With
    
End Property

'Если общий доступ
Public Property Get IsMultiUser() As Boolean
    
    IsMultiUser = m_wbBook.MultiUserEditing
    
End Property

'Проверка существующего листа
Public Property Get IsSheetExists(ByVal sNameSheet As String) As Boolean
    
    Dim wsSheet As Worksheet

    On Error Resume Next
    Set wsSheet = m_wbBook.Sheets(sNameSheet)
    
    IsSheetExists = Not wsSheet Is Nothing
    
End Property
