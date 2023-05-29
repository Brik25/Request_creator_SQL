VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_wsSetting As Worksheet

Private Sub Class_Initialize()
    Set m_wsSetting = ThisWorkbook.Sheets(MConstants.SHEET_SETTING)
End Sub

'Таблица мост
Public Property Get SQLBridge() As Object
    Set SQLBridge = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLBridge")
End Property

'Таблица select
Public Property Get SQLSelect() As Object
    Set SQLSelect = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLSelect")
End Property

'Агрегатные функции
Public Property Get SQLAgregation() As Object
    Set SQLAgregation = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLAgregation")
End Property

'Сортировка
Public Property Get SQLOrder() As Object
    Set SQLOrder = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLOrder")
End Property

'Заменитель значений
Public Property Get SQLLabel() As Object
    Set SQLLabel = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLLabel")
End Property

Public Property Get SQLCondition() As Object
    Set SQLCondition = GetTableValues(m_wsSetting.Range("tableBridge"), "SQLCondition")
End Property

'Таблица заменитель
Public Property Get TableReplaceble() As Object
    Set TableReplaceble = GetTableValues(m_wsSetting.Range("tableReplaceble"))
End Property

'Таблица обработки данных
Public Property Get TableProcessing() As Object
    Set TableProcessing = GetTableValues(m_wsSetting.Range("tableProcessing"))
End Property

'Получаем значения
Private Function GetTableValues(rngTable As Range, Optional sType = "") As Object
    Dim tblSetting As New CTableParametrs
    
    Dim func As New CFunctional
    
    Dim oData As Object, oColumns As Object
    Dim sKey As String
    Dim i As Long

    tblSetting.InitClass m_wsSetting, rngTable
    Set oColumns = tblSetting.TableColumn(False)
    
    Set oData = CreateObject("Scripting.Dictionary")
    
    With m_wsSetting
        
        Select Case func.GetNameRange(rngTable)
            Case "tableBridge"
                Select Case sType
                    Case "SQLBridge" 'Мост связи шаблона с внешней книгой
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C12")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C11")).Value <> "" Then
                                oData.Add sKey, .Cells(i, oColumns("C11")).Value
                            End If
                        Next i
                        
                    Case "SQLSelect" 'Selects для запроса
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C12")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C13")).Value = "" And .Cells(i, oColumns("C11")).Value <> "" Then
                                oData.Add sKey, ""
                            End If
                        Next i
                    
                    Case "SQLAgregation" 'Агрегатные функции
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C12")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C13")).Value <> "" Then
                                oData.Add sKey, .Cells(i, oColumns("C13")).Value
                            End If
                        Next i
                    
                    Case "SQLOrder" 'Сортировка
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C12")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C14")).Value <> "" Then
                                oData.Add sKey, .Cells(i, oColumns("C14")).Value
                            End If
                        Next i
                        
                    Case "SQLLabel" 'Заменитель меток
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C16")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C12")).Value <> "" Then
                                oData.Add sKey, .Cells(i, oColumns("C12")).Value
                            End If
                        Next i
                        
                    Case "SQLCondition" 'Условный оператор
                        For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                            sKey = CStr(.Cells(i, oColumns("C12")).Value)
                            
                            If sKey <> "" And .Cells(i, oColumns("C15")).Value <> "" Then
                                oData.Add sKey, .Cells(i, oColumns("C15")).Value
                            End If
                        Next i
                        
                End Select
            Case "tableReplaceble", "tableProcessing"

                For i = tblSetting.FirstRow + 1 To tblSetting.LastRow
                    sKey = CStr(.Cells(i, tblSetting.FirstColumn).Value)
                    oData.Add sKey, .Cells(i, oColumns("C11")).Value
                Next i
                
        End Select
        
    End With
    
    Set GetTableValues = oData
End Function
