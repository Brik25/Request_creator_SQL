VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Выбор файла
Public Property Get SelectFile() As String
    With Application.FileDialog(msoFileDialogOpen)
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFile = .SelectedItems(1)
        End If
    End With
End Property

'Открытие книги
Public Property Get GetSourceBook(sPath As String, Optional fReadOnly As Boolean = False) As Workbook
    Dim oFile As Object
    
    Application.ScreenUpdating = False
    Set oFile = Workbooks.Open(Filename:=sPath, UpdateLinks:=False, ReadOnly:=fReadOnly)
    oFile.Windows(1).Visible = False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True

    Set GetSourceBook = oFile
End Property

'Закрытие книги
Public Sub CloseWorkbook(oFile As Object, Optional fSave As Boolean = False)
    oFile.Windows(1).Visible = True
    oFile.Close savechanges:=fSave
    Set oFile = Nothing
End Sub

'Получаем лист с помощью условия
Public Property Get GetSheet(wbTmp As Workbook, sNameFind As String) As Worksheet
    Dim tblSource As New CTableParametrs
    
    Dim rngFind As Range
    Dim wsTmp As Worksheet
    
    For Each wsTmp In wbTmp.Worksheets
        With wsTmp
            Set rngFind = .UsedRange.Find(sNameFind, LookAt:=xlWhole)
            If Not rngFind Is Nothing Then
                tblSource.InitClass wsTmp, rngFind
                If tblSource.FirstRow < tblSource.LastRow Then
                    Set GetSheet = wsTmp
                    Exit Property
                End If
            End If
        End With
    Next wsTmp
End Property

'Переопределение пути файла под временную папку
Public Property Get GetTemplateFile(oPaths As Object) As Object
    Dim func As New CFunctional
    
    Dim fso As Object, oFile As Object
    Dim vKey As Variant
    Dim sNewName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CreateTemplateFolder True
    
    For Each vKey In oPaths
        Set oFile = fso.GetFile(vKey)                                      'указываем путь к файле
        sNewName = func.CreateUniqNumbers() & ".xlsb"                         'Генерируем уникальное имя
        fso.CopyFile oFile, MConstants.PATH_TMP_FOLDER & "\" & sNewName    'Копируем файлы
        oPaths.Key(vKey) = MConstants.PATH_TMP_FOLDER & "\" & sNewName     'Копируем во временную папку
    Next vKey
    
    Set GetTemplateFile = oPaths
End Property

'Работа с временной папкой
Public Sub CreateTemplateFolder(Optional fCreateFolder As Boolean = False)
    Dim check As New CCheck
    
    If check.IsPathExists(MConstants.PATH_TMP_FOLDER) Then
        Shell "cmd /c rd /S/Q """ & MConstants.PATH_TMP_FOLDER
    End If
    
    If fCreateFolder Then
        Application.Wait (Now + TimeValue("0:00:02"))
        MkDir MConstants.PATH_TMP_FOLDER
        Application.Wait (Now + TimeValue("0:00:02"))
    End If

End Sub

'Получаем книгу на основке поиска листа
Public Property Get GetBook(sNameSheet As String) As Workbook
    Dim setting As New CSetting
    
    Dim check As New CCheck
    
    Dim wbTmp As Workbook
    
    For Each wbTmp In Application.Workbooks
        If wbTmp.Name <> ThisWorkbook.Name And check.IsSheetExists(sNameSheet, wbTmp) Then
            Set GetBook = wbTmp
            Exit Property
        End If
    Next wbTmp
End Property

'Получаем новый путь
Public Property Get CreateFolder() As String
    
    Dim check As New CCheck
    
    Dim fso As Object, oFile As Object
    Dim sPath As String
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    sPath = ThisWorkbook.Path & "\КД\"

    If check.IsPathExists(sPath) Then
        Set oFile = fso.GetFolder(sPath)
        Shell "cmd /c rd /S/Q """ & sPath
        Application.Wait (Now + TimeValue("0:00:02"))
        MkDir sPath
    Else
        MkDir sPath
    End If

    CreateFolder = sPath
End Property

Public Property Get SQLConnect(Optional wbBook As Workbook) As Object
    Dim oConn As Object
    
    If wbBook Is Nothing Then
        Set wbBook = ThisWorkbook
    End If
    
    Set oConn = CreateObject("ADODB.Connection")
    With oConn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionString = "Data Source=" & wbBook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        .Open
    End With
    
    Set SQLConnect = oConn
    
End Property


