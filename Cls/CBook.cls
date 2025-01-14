VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����� �� �������

Option Explicit

Private m_wbBook    As Workbook
Private m_isSource  As Boolean

'������������� ������
Private Sub Class_Initialize()

    Set m_wbBook = ThisWorkbook
    
End Sub

'������������� ������
Public Sub SetSourceWorkbook(wbBook As Workbook)

    m_wbBook = wbBook
    m_isSource = True

End Sub

'����� ������
Public Sub SetApplicationWorkbook(Optional ByVal isEnable As Boolean = True, _
                               Optional ByVal isCalculate As Boolean = True)

    With Application
        .DisplayAlerts = isEnable
        .ScreenUpdating = isEnable
        .EnableEvents = isEnable
        .Calculation = IIf(isCalculate, xlCalculationAutomatic, xlCalculationManual)
    End With

End Sub

'������ ����� ������ ���������
Public Sub SetApplicationReference()
    
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        End If
    End With

End Sub

'������� �����
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

'��������� ������ �������
Public Sub EnableGeneralAccess(Optional ByVal IsKeepChange As Boolean = False, _
                               Optional ByVal iHistoryDuration As Integer = 30)

    '���� ����� ������ �������
    If IsMultiUser Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
        
    With m_wbBook
        .SaveAs Filename:=wbBook.FullName, accessMode:=xlShared '������� ����
        .KeepChangeHistory = IsKeepChange
        .ChangeHistoryDuration = iHistoryDuration
    End With
    
    Application.DisplayAlerts = True
    
End Sub

'���������� ������ �������
Public Sub DisableGeneralAccess()

    '���� ����� ������ �������
    If Not IsMultiUser Then
        Exit Sub
    End If
    
    Application.DisplayAlerts = False
    
    m_wbBook.ExclusiveAccess
    
    Application.DisplayAlerts = True
    
End Sub

'������� �������� ����
Public Property Get GetCopySheet(ByVal sNameMaket As String, _
                                 ByVal sNewName As String, _
                                 Optional ByVal isAfterCopy As Boolean = False) As Worksheet
    
    Dim wsMaket         As Worksheet, wsNewSheet    As Worksheet
    Dim isVisibleBook   As Boolean, isVisibleSheet  As Boolean
    
    '���� ����� ������, �� ����������
    If m_isSource And Not m_wbBook.Windows(1).Visible Then
        m_wbBook.Windows(1).Visible = True
        isVisibleBook = True
    End If
    
    Set wsMaket = m_wbBook.Sheets(sNameMaket)
    
    With wsMaket
    
        '���� ���� �����

        If .Visible = xlSheetVeryHidden Then
            .Visible = xlSheetVisible
            isVisibleSheet = True
        End If

        DeleteSheets sNewName
        
        .Copy After:=m_wbBook.Sheets(IIf(isAfterCopy, sNameMaket, m_wbBook.Sheets.Count))   '�������� ����
        
        Set wsNewSheet = ActiveSheet
        wsNewSheet.Name = sNewName
        
        .Visible = IIf(isVisibleSheet, xlSheetVeryHidden, xlSheetVeryHidden)
        
        If isVisibleBook Then
            m_wbBook.Windows(1).Visible = Not isVisibleBook
        End If

        Set GetCopySheet = wsNewSheet
        
    End With
    
End Property

'���� ����� ������
Public Property Get IsMultiUser() As Boolean
    
    IsMultiUser = m_wbBook.MultiUserEditing
    
End Property

'�������� ������������� �����
Public Property Get IsSheetExists(ByVal sNameSheet As String) As Boolean
    
    Dim wsSheet As Worksheet

    On Error Resume Next
    Set wsSheet = m_wbBook.Sheets(sNameSheet)
    
    IsSheetExists = Not wsSheet Is Nothing
    
End Property
