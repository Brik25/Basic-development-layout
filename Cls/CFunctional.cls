VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFunctional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим
'Класс являем вспомогательным, в котором находятся вспомогательные механизмы
'Префикса типа String_ и.т.д служат обозоначением к той или иной принадлежности и нужны для избавления лишних классов

Option Explicit


'Определяем значение на основе формуляра
Public Property Get String_ValueDictionary(sValue As String, _
                                           oDictionary As Object) As String

    If oDictionary.exists(sValue) Then
        String_ValueFormular = oDictionary(sValue)
    Else
        String_ValueFormular = sValue
    End If

End Property

'Получаем корекктные номера без ведущего нуля ( 0000123 -> 123 )
Public Property Get String_RemoveLeadingNull(ByVal sText As String, _
                                             Optional ByVal sDelimter As String = vbNullString) As String
                                    
    Dim sResult As String
    Dim vArray  As Variant
    Dim i       As Long
    
    On Error GoTo ErrorHandler
    
    vArray = Split(sText, sDelimter)
    For i = 0 To UBound(vArray) - LBound(vArray)
        If LenB(vArray(i)) > 0 And vArray(i) <> " " Then
            sResult = sResult & CDbl(vArray(i)) & sDelimter
        End If
    Next i
    
    String_RemoveLeadingNull = String_RStrip(sResult, sDelimter)
    
    Exit Property
    
ErrorHandler:

    String_RemoveLeadingNull = sText
    
End Property

'Получаем список из уникальных значений
Public Property Get String_JoinUniq(vArgs As Variant, _
                            Optional ByVal sDelimter As String = vbNullString) As String

    String_JoinUniq = Join( _
                            Dictionary_SplitValue(Join(vArgs, sDelimter), sDelimter).Keys, _
                            sDelimter _
                           )
    
End Property

'Убираем символ с двух сторон
Public Property Get String_Strip(ByVal sText As String, _
                                 Optional ByVal sDelimter As String = vbNullString) As String

    If LenB(sDelimter) > 0 Then
        While Right(sText, Len(sDelimter)) = sDelimter
            sText = Left(sText, Len(sText) - Len(sDelimter))
        Wend

        While Left(sText, Len(sDelimter)) = sDelimter
            sText = Right(sText, Len(sText) - Len(sDelimter))
        Wend
    End If
    
    String_Strip = sText
    
End Property

'Убираем символ справа
Public Property Get String_RStrip(ByVal sText As String, _
                                  Optional ByVal sDelimter As String = vbNullString) As String
    
    If LenB(sDelimter) > 0 Then
        While Right(sText, Len(sDelimter)) = sDelimter
            sText = Left(sText, Len(sText) - Len(sDelimter))
        Wend
    End If
    
    String_RStrip = sText
    
End Property

'Убираем символ слева
Public Property Get String_LStrip(ByVal sText As String, _
                                  Optional ByVal sDelimter As String = vbNullString) As String
    
    If LenB(sDelimter) > 0 Then
        While Left(sText, Len(sDelimter)) = sDelimter
            sText = Right(sText, Len(sText) - Len(sDelimter))
        Wend
    End If
    
    String_LStrip = sText
    
End Property

'Преобразования из полного имени в короткий
Public Property Get String_GetShortName(ByVal sFullName As String) As String
    
    On Error GoTo ErrorHandler
    
    String_GetShortName = Split(sFullName)(0) & " " & Left(Split(sFullName)(1), 1) & "." & Left(Split(sFullName)(2), 1) & "."
    
    Exit Property
    
ErrorHandler:

    String_GetShortName = sFullName
    
End Property

'Получаем имя именованой ячейки
Public Property Get String_GetNameRange(rngTmp As Range) As String

    String_GetNameRange = Split(rngTmp.Name.Name, "!")(1)
    
End Property

'Проверка пусты объекта
Public Property Get Dictionary_IsEmpty(oData As Object) As Boolean
    
    If oData Is Nothing Then
        IsEmptyData = True
        Exit Property
    End If
    
    Dinctionary_IsEmpty = oData.Count = 0
    
End Property

'Получаем список колонок из преобразования массива и ложим как словарь
Public Property Get Dictionary_SplitValue(ByVal sText As String, _
                                          Optional ByVal sDelimter As String = vbNullString) As Object
                                
    Dim oData   As Object
    Dim vArray  As Variant
    Dim i       As Long
    
    vArray = Split(sText, sDelimter)
    
    Set oData = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(vArray) - LBound(vArray)
        If Not oData.exists(vArray(i)) And LenB(vArray(i)) > 0 Then
            oData.Add vArray(i), vbNullString
        End If
    Next i

    Set Dictionary_SplitValue = oData
    
End Property

'Переварачиваем словарь
Public Property Get Dictionary_Reverse(oData As Object) As Object
    Dim oTmpData    As Object
    Dim i           As Long
    
    Set oTmpData = CreateObject("Scripting.Dictionary")
    For i = oData.Count - 1 To 0 Step -1
        oTmpData.Add oData.Keys()(i), oData.Items()(i)
    Next i

    Set Dictionary_Reverse = oTmpData
    
End Property

'Меняем местами текстовые ключ со значениями
'Key, Value не должны быть типа Object

Public Property Get Dictionary_SwapKeyAndValue(oData As Object) As Object
    Dim vKey        As Variant
    Dim oTmpData    As Object
    
    Set oTmpData = CreateObject("Scripting.Dictionary")
    
    For Each vKey In oData
        If Not oTmpData.exists(oData(vKey)) Then
        
            If IsObject(vKey) Then 'Объекты не меняем
                Exit Property
            End If
        
            oTmpData.Add oData(vKey), vKey
        End If
    Next vKey
    
    Set Dictionary_SwapKeyAndValue = oTmpData
    
End Property

'Сортируем ключи
Public Sub Dictionary_Sort(ByRef oData As Object)
    Dim oList   As Object, oSortedData  As Object
    Dim vKey    As Variant

    If Not oData Is Nothing Or oData.Count = 0 Then
        Exit Sub
    End If
    
    Set oList = CreateObject("System.Collections.arrayList")
    For Each vKey In oData.Keys
        oList.Add vKey
    Next vKey
    
    oList.Sort
    
    Set oSortedData = CreateObject("Scripting.Dictionary")
    For Each vKey In oList
        oSortedData.Add vKey, oData(vKey)
    Next vKey
    
    Set oData = oSortedData

End Sub

'Объединяем словари
Public Sub Dictionary_Merge(ByRef oData As Object, _
                            oTmpData As Object)
                    
    Dim vKey As Variant
    
    For Each vKey In oTmpData
        If Not oData.exists(vKey) Then
            oData.Add vKey, oTmpData(vKey)
        End If
    Next vKey

End Sub

Public Property Get Dictionary_Clone(oData As Object) As Object
    Dim oNewData As Object
    Dim vKey As Variant
    
    Set oNewData = CreateObject("Scripting.Dictionary")
    For Each vKey In oData
        If IsObject(oData(vKey)) Then
            oNewData.Add vKey, Dictionary_Clone(oData(vKey))
        Else
            oNewData.Add vKey, oData(vKey)
        End If
    Next vKey

    Set Dictionary_Clone = oNewData
    
End Property

'Есть ли значение в массиве
Public Property Get Array_IsExistInArray(vValue As Variant, _
                                         vArray As Variant) As Boolean
    
    On Error Resume Next
    
    'IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
    
    Array_IsExistInArray = WorksheetFunction.Match(vValue, vArray, 0) > 0
    
End Property

'Считаем количество вхождение в массив (можно указать диапозон, тогда будет учитваться общая сумма)
Public Property Get Array_CountValue(vValues As Variant, _
                                 vArray As Variant) As Long
                                
    Dim check   As New CCheck
    Dim vValue    As Variant
    Dim n       As Long

    For Each vArg In vValues
        If IsExistInArray(vValue, vArray) Then
            n = n + 1
        End If
    Next vArg
    
    Array_CountValue = n
    
End Property
