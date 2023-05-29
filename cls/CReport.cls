VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Собираем данные
Private Function SQLGetData(wbSource As Workbook, wsSource As Worksheet) As Object
    Dim tblSource As New CTableParametrs
    
    Dim setting As New CSetting
    Dim check As New CCheck
    Dim file As New CFile
    Dim func As New CFunctional
    
    Dim oCondition As Object, oSelects As Object, oAgregation As Object, oOrder As Object, oQuery As Object, oSQLColumns As Object, oBridge As Object
    Dim oConn As Object, oReplaceble As Object, oColumns As Object, oData As Object, oValue As Object
    Dim sTblRange As String, sSelect As String, sGroup As String, sOrder As String, sQuery As String, sTmp As String, sKey As String
    Dim vKey As Variant, vSubKey As Variant
    Dim i As Long
    
    Set oConn = file.SQLConnect(wbSource) 'Открываем соединение SQL
    Set oData = CreateObject("Scripting.Dictionary")
    
    With wsSource
        
        'Инициализация структур
        With setting
            Set oReplaceble = .TableReplaceble   'Заменяемые значения
            Set oBridge = .SQLBridge             'Мост столбцов
            Set oSQLColumns = .SQLLabel          'Заменяемые значения
            Set oSelects = .SQLSelect            'Выборка
            Set oCondition = .SQLCondition       'Условие
            Set oAgregation = .SQLAgregation     'Агрегация
            Set oOrder = .SQLOrder               'Сортировка
        End With
        
        tblSource.InitClass wsSource, tblSource.GetRange(oReplaceble("tmp01"), wsSource)
        Set oColumns = tblSource.TableColumn
        
        'Вся таблица
        sTblRange = "[" & .Name & "$" & func.GetAddress(tblSource.FirstColumn) & tblSource.FirstRow & ":" & _
                    func.GetAddress(tblSource.LastColumn) & tblSource.LastRow & "]"
        
        'Построитель запросы
        If Not check.IsEmptyData(oSelects) Then 'Выборка
            For Each vKey In oSelects
                If Not oCondition.exists(vKey) Then
                    sSelect = sSelect & "[" & vKey & "],"
                Else 'Условие
                    sTmp = oCondition(vKey)
                    For Each vSubKey In oSQLColumns
                        sTmp = Replace(sTmp, vSubKey, "[" & oSQLColumns(vSubKey) & "]")
                    Next vSubKey
                    
                    sSelect = sSelect & sTmp & " AS [" & vKey & "] ,"
                End If
            Next vKey
        End If

        If Not check.IsEmptyData(oAgregation) Then 'Агрегация
            For Each vKey In oAgregation
                sSelect = sSelect & oAgregation(vKey) & "([" & vKey & "]) AS [" & vKey & "],"
            Next vKey
        End If

        For Each vKey In oColumns 'Гуппировка
            If oAgregation.exists(vKey) Then
                If oAgregation(vKey) <> "MIN" And oAgregation(vKey) <> "MAX" Then  'Группируем все кроме min и max
                    sGroup = sGroup & "[" & vKey & "],"
                End If
            Else
                sGroup = sGroup & "[" & vKey & "],"
            End If
        Next vKey
        
        If Not check.IsEmptyData(oOrder) Then 'Сортировка
            For Each vKey In oOrder
                If oAgregation.exists(vKey) Then
                    sOrder = sOrder & oAgregation(vKey) & "([" & vKey & "]) " & oOrder(vKey) & ","
                Else
                    sOrder = sOrder & "[" & vKey & "] " & oOrder(vKey) & ","
                End If
                
            Next vKey
        End If

        sSelect = func.RTrip(sSelect, ",")
        sGroup = func.RTrip(sGroup, ",")
        sOrder = func.RTrip(sOrder, ",")

        sQuery = "SELECT " & sSelect & " FROM " & sTblRange & " WHERE [" & oSQLColumns("C01") & "] <> '' GROUP BY " & sGroup & " ORDER BY " & sOrder

        Set oQuery = oConn.Execute(sQuery)

        'Формируем структуру данных под основной шаблон
        Do While Not oQuery.EOF
            Set oValue = CreateObject("Scripting.Dictionary")
            sKey = ""
            
            For i = 0 To oQuery.Fields.Count - 1
                'Производим разбиение по регистрационным номера договоров
                If oQuery.Fields(i).Name = oSQLColumns("C01") Then
                    sKey = oQuery.Fields(i).Value
                End If

                If oBridge.exists(oQuery.Fields(i).Name) Then
                    oValue.Add oBridge(oQuery.Fields(i).Name), oQuery.Fields(i).Value
                End If
            Next i

            If Not oData.exists(sKey) Then
                oData.Add sKey, New Collection
            End If

            oData(sKey).Add oValue
            
            oQuery.MoveNext
        Loop
    End With
    
    oQuery.Close
    oConn.Close
    Set oConn = Nothing
    Set oQuery = Nothing
    
    Set SQLGetData = oData
    
End Function
