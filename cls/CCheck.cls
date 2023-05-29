VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Проверка существующего листа
Public Property Get IsSheetExists(sName As String, Optional wbTmp As Workbook) As Boolean
    Dim wsTmp As Worksheet
    
    If wbTmp Is Nothing Then
        Set wbTmp = ThisWorkbook
    End If
        
    On Error Resume Next
    Set wsTmp = wbTmp.Sheets(sName)
    
    IsSheetExists = Not wsTmp Is Nothing
End Property

'Проверка пусты объекта
Public Property Get IsEmptyData(oData As Object) As Boolean

    If oData Is Nothing Then
        IsEmptyData = True
        Exit Property
    End If
    
    IsEmptyData = oData.Count = 0
End Property

'Проверка существует ли путь
Public Property Get IsPathExists(vFile As Variant) As Boolean
    If vFile <> "" Then
        On Error Resume Next
        IsPathExists = Dir(vFile, vbDirectory) <> ""
    End If
End Property

'Проверка диапозона
Public Property Get IsRangeExists(sName As String, wsTmp As Worksheet) As Boolean
    Dim sAddress As String
    On Error Resume Next
    
    sAddress = wsTmp.Range(sName).Address
    Err.Clear
    On Error GoTo 0
    
    IsRangeExists = sAddress <> ""
End Property

'Проверка на Открытость книги
Public Property Get IsBookOpenClosed(fPath As String) As Boolean
    Dim ff As Integer
    ff = FreeFile
    
    On Error Resume Next
    Open fPath For Random Access Read Write Lock Read Write As #ff
    Close #ff
    
    IsBookOpenClosed = (Err.Number <> 0)
End Property

'Проверяем есть ли данные в таблице
Public Property Get IsRowsExists(wsSheet As Worksheet, Optional rngSheet As Range) As Boolean
    Dim tblSheet As New CTableParametrs
    
    tblSheet.InitClass wsSheet, rngSheet
    
    IsRowsExists = tblSheet.FirstRow < tblSheet.LastRow
End Property

'Общий доступ
Public Property Get IsMultiUser(Optional wbTmp As Workbook) As Boolean
    If wbTmp Is Nothing Then
        Set wbTmp = ThisWorkbook
    End If
    
    IsMultiUser = wbTmp.MultiUserEditing
End Property

'Есть ли значение в массиве
Public Property Get IsExistInArray(vArg As Variant, vArray As Variant) As Boolean
    
    On Error Resume Next
    IsExistInArray = Application.WorksheetFunction.Match(vArg, vArray, 0) > 0
    
End Property
