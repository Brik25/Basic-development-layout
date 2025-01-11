VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим

Option Explicit

Private m_Connect   As Object
Private m_Request   As Object

'Подключение к базе
Public Sub Connect(Optional wbBook As Workbook = Nothing)
    Dim oConnect As Object
    
    If wbBook Is Nothing Then
        Set wbBook = ThisWorkbook
    End If
    
    Set oConnect = CreateObject("ADODB.Connection")
    With oConn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionString = "Data Source=" & wbBook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        .Open
    End With
    
    Set m_Connect = oConnect
    
End Sub

'Выполнение запроса
Public Sub Execute(sRequest As String)
    Set m_oData = m_Connect.Execute(sRequest)
End Sub

'Получаем данные запроса
Public Property Get GetDataReques() As Collection
    Dim oData           As Object
    Dim oCollection     As Collection
    Dim i               As Long
    
    Set oCollection = New Collection
    
    With m_Request
        Do While Not .EOF
            
            Set oData = CreateObject("Scripting.Dictionary")
            
            For i = 0 To .Fields.Count - 1
                oData.Add .Fields(i).Name, .Fields(i).Value
            Next i
            
            .MoveNext
            oCollection.Add oData
            
        Loop
    End With
    
    Set GetDataReques = oCollection
    
End Property

'Закрыть соединение
Public Sub CloseConnect()
    Set m_Connect = Nothing
    Set m_oData = Nothing
End Sub


