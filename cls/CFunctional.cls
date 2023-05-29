VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFunctional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------Формирование GUID номера-----------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Type GUID_TYPE
    lData1 As Long
    lData2 As Long
    lData3 As Long
    bData4(7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (oGuid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (oGuid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr

Public Property Get CreateUniqNumbers() As String
    Dim oGuid As GUID_TYPE, sGuid As String, retValue As LongPtr
    
    Const guidLength As Long = 39
    
    retValue = CoCreateGuid(oGuid)
    
    If retValue = 0 Then
        sGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(oGuid, StrPtr(sGuid), guidLength)
        If retValue = guidLength Then
            CreateUniqNumbers = Mid(sGuid, InStr(sGuid, "{") + 1, InStr(sGuid, "}") - InStr(sGuid, "{") - 1)
        End If
    End If
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------Убираем символы----------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Убираем символ с двух сторон
Public Property Get Trip(sText As String, Optional sDelimter As String = "") As String
    Dim sTextResult As String
    
    sTextResult = sText
    
    While Right(sTextResult, 1) = sDelimter
        sTextResult = Left(sTextResult, Len(sTextResult) - 1)
    Wend
    
    While Left(sTextResult, 1) = sDelimter
        sTextResult = Right(sTextResult, Len(sTextResult) - 1)
    Wend
    
    Trip = sTextResult
End Property

'Убираем символ справа
Public Property Get RTrip(sText As String, Optional sDelimter As String = "") As String
    Dim sTextResult As String
    
    sTextResult = sText
    
    While Right(sTextResult, 1) = sDelimter
        sTextResult = Left(sTextResult, Len(sTextResult) - 1)
    Wend
 
    RTrip = sTextResult
End Property

'Убираем символ слева
Public Property Get LTrip(sText As String, Optional sDelimter As String = "") As String
    Dim sTextResult As String
    
    sTextResult = sText

    While Left(sTextResult, 1) = sDelimter
        sTextResult = Right(sTextResult, Len(sTextResult) - 1)
    Wend

    LTrip = sTextResult
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------Объединение/Разбиение-------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Получаем список колонок из преобразования массива и ложим как словарь
Public Property Get SplitValue(sText As String, Optional sDelimter As String = "") As Object
    Dim oData As Object
    Dim vArray As Variant
    Dim i As Long
    
    vArray = Split(sText, sDelimter)
    
    Set oData = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(vArray) - LBound(vArray)
        If Not oData.exists(vArray(i)) And vArray(i) <> "" Then
            oData.Add vArray(i), ""
        End If
    Next i

    Set SplitValue = oData
End Property

'Получаем список из уникальных значений
Public Property Get JoinUniq(vArgs As Variant, Optional sDelimter As String = "") As String
    JoinUniq = Join(SplitValue(Join(vArgs, sDelimter), sDelimter).Keys, sDelimter)
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------Адресс/Дапозон/Имена--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Получаем адресс
Public Property Get GetAddress(lCol As Long) As String
    GetAddress = Split(Cells(1, lCol).Address(True, False), "$")(0)
End Property

'Получаем диапозон
Public Property Get GetAddressRange(wsTmp As Worksheet, lCol As Long, lFirstRow As Long, lLastRow As Long) As String
    GetAddressRange = wsTmp.Range(wsTmp.Cells(lFirstRow, lCol), wsTmp.Cells(lLastRow, lCol)).Address(False, True, external:=True)
End Property

'Получаем Имя именованой ячейки
Public Property Get GetNameRange(rngTmp As Range) As String
    GetNameRange = Split(rngTmp.Name.Name, "!")(1)
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------Обработка текста--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Первая буква заглавная
Public Property Get TitleText(sText As String) As String
    TitleText = Application.WorksheetFunction.Proper(sText)
End Property

'Получаем корекктные номера без ведущего нуля ( 0000123 -> 123 )
Public Property Get RemoveLeadingNull(sText As String, Optional sDelimter As String = "") As String
    Dim sResult As String
    Dim vArray As Variant
    Dim i As Long
    
    vArray = Split(sText, sDelimter)
    For i = 0 To UBound(vArray) - LBound(vArray)
        If vArray(i) <> "" Then
            sResult = sResult & CDbl(vArray(i)) & sDelimter
        End If
    Next i
    
    RemoveLeadingNull = RTrip(sResult, sDelimter)
End Property

'Преобразования из полного имени в короткий
Public Property Get ShortName(sFullName As String) As String
    On Error GoTo standartName
    
    ShortName = Split(sFullName)(0) & " " & Left(Split(sFullName)(1), 1) & "." & Left(Split(sFullName)(2), 1) & "."
    
    Exit Property
standartName:

    ShortName = sFullName
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------Обработка текста--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Получаем год
Public Property Get GetYear() As Long
    GetYear = DatePart("yyyy", Now())
End Property

'Получаем текущий месяц
Public Property Get GetMonth(Optional fText As Boolean = False, Optional fFirstNull As Boolean = False) As Variant
    Dim lMonth As Long
    Dim vMonth As Variant
    
    lMonth = DatePart("m", Now())
    vMonth = Array("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
    
    If fText Then
        GetMonth = vMonth(lMonth - 1)
    Else
        If fFirstNull Then
            GetMonth = Application.WorksheetFunction.Text(lMonth, "00")
        Else
            GetMonth = Trim(lMonth)
        End If
    End If
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------Структуры-----------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Считаем количество вхождение в массив (можно указать диапозон, тогда будет учитваться общая сумма)
Public Property Get CountInArray(vArgs As Variant, vArray As Variant) As Long
    Dim check As New CCheck
    Dim vArg As Variant
    Dim n As Long
    
    n = 0
    For Each vArg In vArgs
        If check.IsExistInArray(vArg, vArray) Then
            n = n + 1
        End If
    Next vArg
    
    CountInArray = n
End Property

'Переварачиваем словарь
Public Property Get ReversedObject(oData) As Object
    Dim oTmpData As Object
    Dim i As Long
    
    Set oTmpData = CreateObject("Scripting.Dictionary")
    For i = oData.Count - 1 To 0 Step -1
        oTmpData.Add oData.Keys()(i), oData.Items()(i)
    Next i

    Set ReversedObject = oTmpData
End Property

'Меняем местами текстовые ключ со значениями(не для объектов)
Public Property Get SwapTextKeyObject(oData As Object) As Object
    Dim vKey As Variant
    Dim oTmpData As Object
    
    Set oTmpData = CreateObject("Scripting.Dictionary")
    
    For Each vKey In oData
        If Not oTmpData.exists(oData(vKey)) Then
            oTmpData.Add oData(vKey), vKey
        End If
    Next vKey
    
    Set SwapTextKey = oTmpData
End Property

'Вытаскиваем значения из нижнего уровня словаря
Public Property Get UbondObject(oData As Object, Optional fCreateObject As Boolean = False) As Object
    Dim oTmpData As Object
    Dim vKey As Variant, vValue As Variant
    
    Set oTmpData = CreateObject("Scripting.Dictionary")
    
    For Each vKey In oData
        For Each vValue In oData(vKey)
            If fCreateObject Then
                oTmpData.Add vValue, CreateObject("Scripting.Dictionary")
            Else
                oTmpData.Add vValue, ""
            End If
        Next vValue
    Next vKey
    
    Set UbondObject = oTmpData
End Property

'Объединяем словари
Public Property Get MergeObject(oData As Object, oTmpData As Object) As Object
    Dim vKey As Variant
    
    For Each vKey In oTmpData
        If Not oData.exists(vKey) Then
            oData.Add vKey, oTmpData(vKey)
        End If
    Next vKey
    
    Set MergeObject = oData
End Property

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------Прочее------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Проставляем нумерацию
Public Sub RenumberPosition(wsTmp As Worksheet)
    Dim tblSheet As New CTableParametrs
    Dim oColumns As Object
    Dim i As Long, lCount As Long
    
    With wsTmp
        tblSheet.InitClass wsTmp
        Set oColumns = tblSheet.TableColumn
        
        lCount = 1
        For i = tblSheet.FirstRow + 1 To tblSheet.LastRow
            .Cells(i, tblSheet.FirstColumn).Value = lCount
            lCount = lCount + 1
        Next i
    End With
End Sub

