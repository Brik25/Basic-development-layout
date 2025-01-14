VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����� �������� ��������� �� ���������� ��������� � ������ ��������� ����������� ���������
'����������� �� �������: CTable, CSheet, CFunctional

Option Explicit

Private cm_ws       As CSheet
Private cm_func     As CFunctional
Private cm_table    As CTable

Private m_wsSheet   As Worksheet
Private m_sLog      As String

'������������� �������
Public Sub InitClass(Optional wbSetting As Workbook = Nothing, _
                     Optional wsSetting As Worksheet = Nothing)

    
    If wbSetting Is Nothing Then
        Set wbSetting = ThisWorkbook
    End If

    If wsSetting Is Nothing Then
        Set m_wsSheet = wbSetting.Sheets(MConstants.SHEET_SETTING)
    Else
        Set m_wsSheet = wsSetting
    End If
    
    '�������������� ��� ����������� �����
    Set cm_ws = New CSheet
    cm_ws.InitClass m_wsSheet
    
    '�������������� ��� ����������
    Set cm_func = New CFunctional
    
End Sub

'��������� ������������ � ������� ����������
Private Property Get IsExistsInfo(sNameRange As String, _
                                  Optional IsTable�ontainsObject As Boolean = False) As Boolean
                                 
    Dim sLog        As String
    Dim IsCorrect   As Boolean
    
    '��������� ������� ��������� �� �����
    If Not cm_ws.IsRangeExists(sNameRange) Then
        IsExistsInfo = False
        Exit Property
    End If
    
    IsExistsInfo = True
  
    If IsTable�ontainsObject Then
    
        Set cm_table = New CTable
        
        With cm_table
            .InitClass m_wsSheet, sNameRange
            
            If Not .IsAccesTable Or .IsEmptyTable Then
                Set cm_table = Nothing
                IsExistsInfo = False
                Exit Property
            End If
            
            .SetCellLowerBound .GetLastRow, _
                               .GetFindColumn(, , .GetRangeUpperBound) - 1

            '���� �� ��������� ��������
            IsExistsInfo = True
            
        End With

    End If
    
End Property

'�������� �������� �� ����� ��������
Public Property Get GetRange(ByVal sNameRange As String) As Range
    
    If IsExistsInfo(sNameRange) Then
        Set GetRange = m_wsSheet.Range(sNameRange)
    End If
    
End Property

'�������� �������� �� ����� ��������
Public Property Get GetValue(ByVal sNameRange As String) As Variant
    
    If IsExistsInfo(sNameRange) Then
        GetValue = m_wsSheet.Range(sNameRange).Value
    End If

End Property

'�������� ������ �� ����� ��������
Public Property Get GetArrayData(ByVal sNameRange As String, _
                              Optional sDataOption As String = vbNullString) As Object
     
    If IsExistsInfo(sNameRange, True) Then
        'pass
 
    End If
    
End Property


