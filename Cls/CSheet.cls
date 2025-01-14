VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����� �� �������

Option Explicit

Private m_wsSheet   As Worksheet

Private Const MAX_LEVEL_OUTLINE As Byte = 8

'������������� ������
Public Sub InitClass(wsSheet As Worksheet)

    Set m_wsSheet = wsSheet
End Sub

'���������� ������� � ���.���� � �����
Public Sub DestroyerFormats()
    
    With m_wsSheet
        .UsedRange.Validation.Delete
        .UsedRange.FormatConditions.Delete
    End With
    
End Sub

'���������� ������� � �����
Public Sub DestroyerFormullas()

    With m_wsSheet
        .UsedRange.Copy
        .UsedRange.PasteSpecial Paste:=xlPasteValues
    End With
    
End Sub

'���������� ��� ������ � ��������
Public Sub DestroyerShapes()
                        
    Dim vShape As Variant

    For Each vShape In m_wsSheet.Shapes
        vShape.Delete
    Next vShapes
    
End Sub

'����� ����������� (����� � ��������)
Public Sub ShowOutline()
    Dim i As Long
    
    For i = 1 To MAX_LEVEL_OUTLINE
        m_wsSheet.Outline.ShowLevels i, i
    Next i
    
End Sub

'������ ����������� �����
Public Sub HiddenOutlineRows()
    Dim i As Long

    For i = MAX_LEVEL_OUTLINE To 1 Step -1
        m_wsSheet.Outline.ShowLevels RowLevels:=i
    Next i

End Sub

'������ ����������� ��������
Public Sub HiddenOutlineColumns()
    Dim i As Long
    
    For i = MAX_LEVEL_OUTLINE To 1 Step -1
        m_wsSheet.Outline.ShowLevels ColumnLevels:=i
    Next i
    
End Sub

'��������������
Public Sub UnGroupOutline(ByVal lRow As Long, _
                          Optional lCountRows As Long = 0)

    With m_wsSheet
    
        If .Rows(lRow).OutlineLevel > 1 Then
            .Rows(lRow).Resize(lCountRows).Ungroup
        End If
        
    End With
    
End Sub

'�������� �����
Public Sub ProtectSheet(Optional sPassword As String = vbNullString)

    m_wsSheet.Protect PASSWORD:=IIf(sPassword = vbNullString, MConstants.PASSWORD, sPassword), _
    DrawingObjects:=True, UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    m_wsSheet.EnableOutlining = True
    
End Sub

'������������� �����
Public Sub UnProtectSheet(Optional sPassword As String = vbNullString)

    m_wsSheet.Unprotect IIf(sPassword = vbNullString, MConstants.PASSWORD, sPassword)
    
End Sub

'�������� ������
Public Property Get GetAddress�olumn(ByVal lColumn As Long) As String

    GetAddress�olumn = Split(Cells(1, lCol).Address(True, False), "$")(0)
    
End Property

'�������� ��������
Public Property Get GetAddressRange(lFirstRow As Long, _
                                    lFirstCol As Long, _
                                    Optional lLastRow As Long = 0, _
                                    Optional lLastCol As Long = 0)
                                    
    
    lLastRow = IIf(lLastRow = 0, lFirstRow, lLastRow)
    lLastCol = IIf(lLastCol = 0, lFirstCol, lLastCol)
                     
    With m_wsSheet
        
        GetAddressRange = .Range( _
                                    .Cells(lFirstRow, lFirstCol), _
                                    .Cells(lLastRow, lLastCol) _
                                ).Address(False, True, external:=True)
    
    End With

End Property

'�������� ��������� �� �����
Public Property Get IsRangeExists(ByVal sNameRange As String) As Boolean
                                
    Dim sAddress As String
    
    On Error Resume Next
    
    sAddress = m_wsSheet.Range(sNameRange).Address
    Err.Clear
    On Error GoTo 0
    
    IsRangeExists = LenB(sAddress) > 0
    
End Property
