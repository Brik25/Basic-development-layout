VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс не зависим

Option Explicit

Private Const PATH_TMP_FOLDER As String = "C:\TemporaryFolder"

'Создание папки
Public Sub CreateFolder(ByVal sPath As String)

    Application.Wait (Now + TimeValue("0:00:02"))
    MkDir sPath
    Application.Wait (Now + TimeValue("0:00:02"))
    
End Sub

'Удаление папки
Public Sub DeleteFolder(ByVal sPath As String)

    If IsPathExists(sPath) Then
        Shell "cmd /c rd /S/Q """ & sPath
    End If
    
End Sub

'Закрытие книги
Public Sub CloseWorkbook(ByRef oFile As Object, _
                        Optional ByVal isSave As Boolean = False)
    
    oFile.Windows(1).Visible = True
    oFile.Close savechanges:=isSave
    Set oFile = Nothing
    
End Sub

'Получить путь к файлу
Public Property Get GetPathFile() As String
    
    With Application.FileDialog(msoFileDialogOpen)
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFile = .SelectedItems(1)
        End If
    End With
    
End Property

'Выбор файлов с указанием папки
Public Property Get GetPathsFiles(Optional ByVal isExcel As Boolean = False, _
                                  Optional ByVal sPath As String = vbNullString) As Object
                                      
    Dim fso     As Object, oFiles As Object
    Dim vFile   As Variant
    Dim lFile   As Long

    If LenB(sPath) = 0 Then 'Если не передан путь
        sPath = GetPathFile()
    End If
    
    If LenB(sPath) > 0 Then
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set oFiles = CreateObject("Scripting.Dictionary")
    
        For Each vFile In fso.GetFolder(sPath).Files
            
            If Not InStr(vFile.Name, "*~*") <> 0 Then
                
                If isExcel Then
                    If InStr(Right$(LCase(vFile.Name), 4), xls) Then
                        oFiles.Add vFile.Name, vFile.Path
                    End If
                    
                Else
                    oFiles.Add vFile.Name, vFile.Path
                    
                End If

            End If
            
        Next vFile
        
    End If
    
    Set SelectFileSeveral = oFiles
    
End Property

'Получить путь к папке
Public Property Get GetPathFolder() As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems(1)
        End If
    End With
    
End Property

'Открытие книги
Public Property Get GetSourceBook(ByVal sPath As String, _
                                  Optional ByVal isRead As Boolean = False) As Workbook
                                
    Dim oFile As Object
    
    Application.ScreenUpdating = False
    Set oFile = Workbooks.Open(Filename:=sPath, UpdateLinks:=False, ReadOnly:=isRead)
    oFile.Windows(1).Visible = False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True

    Set GetSourceBook = oFile
    
End Property

'Переопределение именна файлов под временную папку
Public Property Get GetTemplateFiles(ByRef oPaths As Object) As Object
    Dim func        As New CFunctional
    Dim fso         As Object, oFile As Object
    Dim vKey        As Variant
    Dim sNewName    As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    DeleteFolder PATH_TMP_FOLDER    'Удаляем папку, если она есть
    CreateFolder PATH_TMP_FOLDER    'Создаем временную папку
    
    For Each vKey In oPaths
        Set oFile = fso.GetFile(vKey)                           'указываем путь к файле
        sNewName = MGuid.Generate & ".xlsb"                     'Генерируем уникальное имя
        fso.CopyFile oFile, PATH_TMP_FOLDER & "\" & sNewName    'Копируем файлы
        oPaths.Key(vKey) = PATH_TMP_FOLDER & "\" & sNewName     'Копируем во временную папку
    Next vKey
    
    Set GetTemplateFiles = oPaths
    
End Property

'Проверка на Открытость книги
Public Property Get IsBookOpenClosed(ByVal sPath As String) As Boolean
    Dim ff As Integer
    
    ff = FreeFile
    
    On Error Resume Next
    Open sPath For Random Access Read Write Lock Read Write As #ff
    Close #ff
    
    IsBookOpenClosed = (Err.Number <> 0)
    
End Property

'Проверка листа
Public Property Get IsPathExists(ByVal sPath As String) As Boolean
    
    If LenB(sPath) > 0 Then
        On Error Resume Next
        IsPathExists = LenB(Dir(sPath, vbDirectory)) > 0
    End If
    
End Property
