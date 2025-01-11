VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим

Option Explicit

Private m_oRegExp As Object

Private Sub Class_Initialize()
    Set m_oRegExp = CreateObject("VBScript.RegExp")
End Sub

'Получаем маску
Private Property Get GetRegExpMask(ByVal sMaskName As String) As String
    
    Select Case sMaskName
        
        Case "Guid":          GetRegExpMask = "^(\{{0,1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{0,1})$"
        Case "INN":           GetRegExpMask = "(ИНН[\s\.:]*|)(\d{12}|\d{10})"
        Case "Participant":   GetRegExpMask = "(Участник|Уч)[\s\.:]*№*[\s\.:]*[0-9]{0,2}\b[\s\.:]*|^(№)*[\s]*[0-9]{0,2}\b[\s\.:]*"
        Case "Object":        GetRegExpMask = "^[А-я]*[0-9]*\b"
        Case "Symbols":       GetRegExpMask = "^[;_,/\.…]+|[;_,/\.…]+$"
        Case "Column":        GetRegExpMask = "(C|B)[0-9]{2}"
        Case "Label":         GetRegExpMask = "tmp[0-9]{2}"
        Case "NumberDog":     GetRegExpMask = "[A-Za-z]{0,1}\d{6,7}[/-]{1}\d{4,5}[СсCcДд]\d{0,3}|(100-){0,1}\d{4}[-]\d{4}[-]\d{1,2}|[0-9]{1}[А-я]{1,2}[0-9]{0,2}.\d{2}[А-я]{0,3}|\d{1,2}.\d{1,2}.\d{1,2}[А-я]{1,3}\d{0,2}"
        Case "Month":         GetRegExpMask = "(янв|февр|март|апрел|май|мая|июн|июл|авгу|сент|октя|нояб|дека)[А-я]*([\s]*(20)[0-9]{2}){0,1}([\s]*(г)[\s\.]*)*" & _
                                              "[\s]*(-|по){0,1}[\s]*" & _
                                              "((янв|февр|март|апрел|май|мая|июн|июл|авгу|сент|октя|нояб|дека)[А-я]*([\s]*(20)[0-9]{2}){0,1}){0,1}"
    End Select

End Property

'Получаем адресс
Private Property Get GetAddress(ByVal lCol As Long) As String

    GetAddress = Split(Cells(1, lCol).Address(True, False), "$")(0)
    
End Property

'Заменяем значение
Public Property Get GetReplaceValue(ByVal sValue As String, _
                                    ByVal sPattern As String, _
                                    Optional ByVal sNewValue As String = vbNullString, _
                                    Optional ByVal isIgnoreCase As Boolean = False) As String
    
    With m_oRegExp
        .Global = True
        .Pattern = GetRegExpMask(sPattern)      'Маска
        .IgnoreCase = isIgnoreCase              'Игнор регистра
        GetReplaceValue = .Replace(sValue, sNewValue)
    End With
    
End Property

'Заменяем временные метки на значения -> tmp10 -> 2
Public Property Get GetReplaceFormulaForLables(ByVal sValue As String, _
                                               oColumns As Object) As String
    
    With m_oRegExp
        .Global = True
        .Pattern = GetRegExpMask("Label")      'Маска
        
        For Each vFormula In .Execute(sFormula)
            sValue = Replace(sValue, vFormula.Value, oColumns(vFormula.Value))
        Next vFormula
        
    End With
    
    GetReplaceFormulaForLables = sValue
    
End Property

'Заменяем временные метки на значения : C10 - > A
Public Property Get GetReplaceFormulaForColumns(ByVal sValue As String, _
                                               oColumns As Object) As String
    
    With m_oRegExp
       .Global = True
       .Pattern = GetRegExpMask("Column")      'Маска
       
       For Each vFormula In .Execute(sFormula)
           sValue = Replace(sValue, vFormula.Value, GetAddress(oColumns(vFormula.Value)))
       Next vFormula
       
    End With
    
    GetReplaceFormulaForColumns = sValue

End Property

'Получаем значение благодаря маске
Public Property Get GetValue(ByVal sValue As String, _
                             ByVal sPattern As String, _
                             Optional ByVal isIgnoreCase As Boolean = False) As String
    
    Dim oMatch As Object
    
    With m_oRegExp
        .Global = True
        .Pattern = sPattern
        .IgnoreCase = isIgnoreCase
        
        If .Test(sValue) Then
            
            Set oMatch = .Execute(sValue)
            If oMatch.Count > 0 Then
                GetValue = oMatch.Item(0).Value
            End If
            
        End If
    End With
  
End Property
