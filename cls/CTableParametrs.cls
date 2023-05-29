VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTableParametrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DataSheet
Option Explicit

Private m_wsTmp As Worksheet
Private m_rngMain As Range

Public Sub InitClass(wsTmp As Worksheet, Optional rngTable As Range)
    
    Set m_wsTmp = wsTmp
    
    If rngTable Is Nothing Then
        Set m_rngMain = m_wsTmp.Range("headerCell")
    Else
        Set m_rngMain = rngTable
    End If
    
End Sub

'Первая строка
Public Property Get FirstRow() As Long
    FirstRow = m_rngMain.Row
End Property

'Первый столбец
Public Property Get FirstColumn() As Long
    FirstColumn = m_rngMain.Column
End Property

'Последняя строка
Public Property Get LastRow(Optional fStandartValue As Boolean = True) As Long
    Dim setting As New CSetting
    Dim check As New CCheck

    If fStandartValue Then
        LastRow = m_wsTmp.Cells(m_wsTmp.Rows.Count, m_rngMain.Column).End(xlUp).Row
    Else
        LastRow = m_wsTmp.Columns(FirstColumn).Find(What:="", SearchOrder:=xlByRows, After:=m_wsTmp.Cells(FirstRow, FirstColumn)).Row - 1
    End If

End Property

'Последний столбец
Public Property Get LastColumn(Optional fStandartValue As Boolean = True) As Long
    Dim setting As New CSetting
    Dim check As New CCheck
    
    If fStandartValue Then
        LastColumn = m_wsTmp.Cells(m_rngMain.Row, m_wsTmp.Columns.Count).End(xlToLeft).Column
    Else
        LastColumn = m_wsTmp.Rows(FirstRow).Find(What:="", SearchOrder:=xlByColumns, After:=m_wsTmp.Cells(FirstRow, FirstColumn)).Column - 1
    End If
    
End Property

'Колонки
Public Property Get TableColumn(Optional fStandartValue As Boolean = True) As Object
    Dim oCollumns As Object
    Dim i As Long
    
    With m_wsTmp
        Set oCollumns = CreateObject("Scripting.Dictionary")
        
        For i = FirstColumn To LastColumn(fStandartValue)
            With .Cells(FirstRow, i)
                If .Value <> "" And Not oCollumns.exists(.Value) Then
                    oCollumns.Add .Value, i
                End If
            End With
        Next i
    End With
    
    Set TableColumn = oCollumns
End Property

'Получаем диапозон таблицы
Public Property Get GetRange(sNameFind As String, Optional wsTmp As Worksheet) As Range
    Dim rngFind As Range

    If Not wsTmp Is Nothing Then
        Set m_wsTmp = wsTmp
    End If

    Set rngFind = m_wsTmp.UsedRange.Find(sNameFind, LookAt:=xlWhole)
    If Not rngFind Is Nothing Then
        Set GetRange = rngFind
        Exit Property
    End If
   
End Property
