Attribute VB_Name = "modSummary1"
Option Explicit

'=====================================================================
'  Сводный отчет по услугам - сборка из папки "отчеты"
'  Самостоятельный вариант (не зависит от других модулей).
'  Запуск:
'    Главная           - пересобрать отчет и показать итог
'    ВыбратьПапку      - указать свою папку с отчетами
'    ДобавитьФайлы     - добавить отдельные файлы (Ctrl - несколько)
'    ОчиститьДопФайлы  - забыть добавленные файлы
'
'  Папка по умолчанию: подпапка "отчеты" рядом с этой книгой.
'  Результат: "Сводный отчет по услугам (макрос).xlsx" рядом с книгой.
'  Отчет всегда пересобирается заново по всем файлам папки.
'=====================================================================

Private Const FOLDER_NAME As String = "отчеты"
Private Const OUT_XLSX As String = "Сводный отчет по услугам (макрос).xlsx"
Private Const APP_TITLE As String = "Сводный отчет"

' свойства документа (сохраняются вместе с книгой)
Private Const PROP_FOLDER As String = "ПапкаОтчетов"
Private Const PROP_EXTRA  As String = "ДопФайлы"

Private Const SH_SUM   As String = "Сводка"
Private Const SH_PCT   As String = "% изменений"
Private Const SH_DIFF  As String = "Расхождения"
Private Const SH_CHART As String = "Графики"
Private Const SH_DET   As String = "Детализация"
Private Const SH_REG   As String = "Исходные отчёты"

Private Const FM_MONEY As String = "#,##0.00"
Private Const FM_PCT   As String = "+0.0%;-0.0%;0.0%"
Private Const FM_QTY   As String = "#,##0.##"
Private Const FM_THR   As String = "0.0%"

Private Const CLR_MAIN As Long = 9393152   ' 00548F
Private Const CLR_TEXT As Long = 8022875   ' 5B6B7A

Private mFolder  As String
Private mOutPath As String
Private mNSvc As Long, mNPers As Long, mNDisc As Long

'--- символы, отсутствующие в CP1251 ---------------------------------
Private Function RUB() As String: RUB = ChrW(&H20BD): End Function   ' знак рубля
Private Function DLT() As String: DLT = ChrW(&H394): End Function    ' дельта
Private Function ARW() As String: ARW = ChrW(&H2192): End Function   ' стрелка
Private Function NUM() As String: NUM = ChrW(&H2116): End Function   ' номер
Private Function LAQ() As String: LAQ = ChrW(&HAB): End Function     ' левая кавычка
Private Function RAQ() As String: RAQ = ChrW(&HBB): End Function     ' правая кавычка

'=====================================================================
'  ТОЧКА ВХОДА
'  silent = True - без диалогов, ошибка возвращается вызывающему
'=====================================================================
Public Sub BuildSummary(Optional ByVal silent As Boolean = False)
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim reports As Collection
    Dim nSeen As Long, nErr As Long, msg As String, s As String

    If Len(ThisWorkbook.Path) = 0 Then
        If silent Then Err.Raise 5, "BuildSummary", "Сохраните книгу с макросом на диске."
        MsgBox "Сохраните книгу с макросом на диске.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    mFolder = GetFolder()
    If Len(Dir(mFolder, vbDirectory)) = 0 Then
        If silent Then
            Err.Raise 5, "BuildSummary", "Не найдена папка: " & mFolder
        Else
            s = PickFolder()
            If Len(s) = 0 Then Exit Sub
            SetFolder s
            mFolder = s
        End If
    End If
    mOutPath = ThisWorkbook.Path & "\" & OUT_XLSX
    mNSvc = 0: mNPers = 0: mNDisc = 0

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    Set reports = New Collection
    ParseAll reports, nSeen, nErr
    If reports.Count = 0 Then
        If silent Then
            Err.Raise 5, "BuildSummary", "Ни один файл не разобран. Папка: " & mFolder
        End If
        msg = "Ни один файл не разобран." & vbCrLf & "Папка: " & mFolder
        If nSeen > 0 Then msg = msg & vbCrLf & "Файлов найдено: " & nSeen
        MsgBox msg, vbExclamation, APP_TITLE
        GoTo Done
    End If

    BuildBook reports

    msg = "Готово: " & mOutPath & vbCrLf & vbCrLf & _
          "Папка: " & mFolder & vbCrLf & _
          "Файлов найдено: " & nSeen & ", разобрано: " & reports.Count
    If nErr > 0 Then msg = msg & ", НЕ распознано: " & nErr
    msg = msg & vbCrLf & _
          "Услуг: " & mNSvc & ", периодов: " & mNPers & _
          ", расхождений: " & mNDisc
    If Len(GetProp(PROP_EXTRA)) > 0 Then
        msg = msg & vbCrLf & "Дополнительно подключено файлов: " & _
              (UBound(Split(GetProp(PROP_EXTRA), ";")) + 1)
    End If
    GoTo Done

Fail:
    msg = "Ошибка " & Err.Number & ": " & Err.Description
    If silent Then
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
        Err.Raise Err.Number, "BuildSummary", Err.Description
    End If
Done:
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    If Not silent Then
        If Err.Number = 0 Then
            MsgBox msg, vbInformation, APP_TITLE
        Else
            MsgBox msg, vbCritical, APP_TITLE
        End If
    End If
    Exit Sub
End Sub

Public Sub Главная()
    BuildSummary False
End Sub

'---------------------------------------------------------------- папка
Public Sub ВыбратьПапку()
    Dim s As String
    s = PickFolder()
    If Len(s) = 0 Then Exit Sub
    SetFolder s
    BuildSummary False
End Sub

Private Function PickFolder() As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(4)   ' 4 = msoFileDialogFolderPicker
    If fd Is Nothing Then Exit Function
    With fd
        .Title = "Папка с файлами отчетов"
        .InitialFileName = GetFolder()
        If .Show = -1 Then PickFolder = CStr(.SelectedItems(1))
    End With
    On Error GoTo 0
End Function

'---------------------------------------------------------------- доп. файлы
Public Sub ДобавитьФайлы()
    Dim fd As Object, it As Variant, s As String
    Dim add As String, cur As String
    On Error Resume Next
    Set fd = Application.FileDialog(3)   ' 3 = msoFileDialogFilePicker
    If fd Is Nothing Then Exit Sub
    With fd
        .Title = "Добавить файлы отчетов (несколько - через Ctrl)"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
        For Each it In .SelectedItems
            s = CStr(it)
            If InStr(1, ";" & add & ";", ";" & s & ";", vbTextCompare) = 0 Then
                If Len(add) > 0 Then add = add & ";"
                add = add & s
            End If
        Next it
    End With
    On Error GoTo 0
    If Len(add) = 0 Then Exit Sub

    cur = GetProp(PROP_EXTRA)
    If Len(cur) > 0 Then cur = cur & ";"
    SetProp PROP_EXTRA, cur & add
    BuildSummary False
End Sub

Public Sub ОчиститьДопФайлы()
    SetProp PROP_EXTRA, ""
    MsgBox "Список добавленных файлов очищен.", vbInformation, APP_TITLE
End Sub

'---------------------------------------------------------------- свойства
Private Function GetProp(ByVal nm As String) As String
    On Error Resume Next
    GetProp = CStr(ThisWorkbook.CustomDocumentProperties(nm).Value)
    On Error GoTo 0
End Function

Private Sub SetProp(ByVal nm As String, ByVal v As String)
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties(nm).Delete
    ThisWorkbook.CustomDocumentProperties.Add Name:=nm, LinkToContent:=False, _
                                              Type:=msoPropertyTypeString, Value:=v
    On Error GoTo 0
End Sub

Private Function GetFolder() As String
    Dim s As String
    s = GetProp(PROP_FOLDER)
    If Len(s) = 0 Then s = ThisWorkbook.Path & "\" & FOLDER_NAME
    GetFolder = s
End Function

Private Sub SetFolder(ByVal s As String)
    SetProp PROP_FOLDER, s
End Sub

'=====================================================================
'  РАЗБОР ИСХОДНЫХ ФАЙЛОВ
'=====================================================================
Private Sub ParseAll(ByRef reports As Collection, ByRef nSeen As Long, ByRef nErr As Long)
    Dim f As String, path As String, i As Long
    Dim extras As Variant, seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    nSeen = 0
    nErr = 0

    ' --- файлы из папки ---
    If Len(mFolder) > 0 Then
        f = Dir(mFolder & "\*.xlsx")
        Do While Len(f) > 0
            If Left$(f, 1) <> "~" Then
                path = mFolder & "\" & f
                If Not seen.Exists(path) Then
                    seen(path) = 1
                    nSeen = nSeen + 1
                    If Not ParseFile(path, reports) Then nErr = nErr + 1
                End If
            End If
            f = Dir
        Loop
    End If

    ' --- отдельные файлы, добавленные вручную ---
    extras = Split(GetProp(PROP_EXTRA), ";")
    For i = LBound(extras) To UBound(extras)
        path = Trim$(CStr(extras(i)))
        If Len(path) > 0 Then
            If Not seen.Exists(path) Then
                seen(path) = 1
                nSeen = nSeen + 1
                If Len(Dir(path)) = 0 Then
                    nErr = nErr + 1        ' файла уже нет на диске
                ElseIf Not ParseFile(path, reports) Then
                    nErr = nErr + 1
                End If
            End If
        End If
    Next i
End Sub

Private Function ParseFile(ByVal path As String, ByRef reports As Collection) As Boolean
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet
    Dim v As Variant, nR As Long, nC As Long, r As Long, c As Long
    Dim repNo As String, svc As String, perRaw As String
    Dim contract As String, manager As String
    Dim hdrRow As Long, cnt As Long
    Dim perKey As String, perEnd As Date
    Dim its As Collection
    Dim total As Double
    Dim fn As String, nm As String, tp As String, un As String
    Dim qt As Double, tf As Double, cs As Double

    Set its = New Collection
    On Error GoTo Fin
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    For Each sh In wb.Worksheets
        If LCase$(sh.Name) = LCase$("Общий") Then Set ws = sh: Exit For
    Next sh
    If ws Is Nothing Then GoTo Fin

    v = ws.UsedRange.Value2
    If Not IsArray(v) Then GoTo Fin
    nR = UBound(v, 1)
    nC = UBound(v, 2)

    ' --- метки и строка шапки ---
    hdrRow = 0
    For r = 1 To nR
        If r > 30 Then Exit For
        For c = 1 To nC
            Select Case CellText(v, r, c)
                Case "Отчет:":           If Len(repNo) = 0 Then repNo = NextValue(v, r, c, nC)
                Case "Услуга:":          If Len(svc) = 0 Then svc = NextValue(v, r, c, nC)
                Case "Отчетный период:": If Len(perRaw) = 0 Then perRaw = NextValue(v, r, c, nC)
                Case "Номер договора:":  If Len(contract) = 0 Then contract = NextValue(v, r, c, nC)
                Case "Сервис-менеджер:": If Len(manager) = 0 Then manager = NextValue(v, r, c, nC)
            End Select
        Next c
        If CellText(v, r, 1) = NUM() & NUM() Then hdrRow = r: Exit For
    Next r
    If hdrRow = 0 Then GoTo Fin

    perKey = PeriodKey(perRaw, perEnd)
    If Len(perKey) = 0 Then GoTo Fin
    fn = Mid$(path, InStrRev(path, "\") + 1)

    ' --- строки таблицы ---
    For r = hdrRow + 1 To nR
        If Len(CellText(v, r, 2)) = 0 Then Exit For
        If Len(ToNum(v, r, 1)) = 0 Then Exit For
        nm = CellText(v, r, 2)
        tp = CellText(v, r, 3)
        un = CellText(v, r, 4)
        qt = ToNum(v, r, 5)
        tf = ToNum(v, r, 6)
        cs = ToNum(v, r, 7)
        total = total + cs
        its.Add Array(nm, tp, un, qt, tf, cs)
        cnt = cnt + 1
    Next r

    reports.Add Array(fn, repNo, svc, perRaw, perKey, perEnd, contract, manager, cnt, total, its)
    ParseFile = True
Fin:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function CellText(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim s As String
    If r > UBound(v, 1) Or c > UBound(v, 2) Then Exit Function
    On Error Resume Next
    s = CStr(v(r, c))
    On Error GoTo 0
    CellText = Trim$(Replace(s, ChrW(160), " "))
End Function

Private Function NextValue(ByVal v As Variant, ByVal r As Long, ByVal c As Long, ByVal nC As Long) As String
    Dim k As Long
    For k = c + 1 To nC
        If Len(CellText(v, r, k)) > 0 Then
            NextValue = CellText(v, r, k)
            Exit Function
        End If
    Next k
End Function

' Число из ячейки: допускает пробелы, неразрывные пробелы и запятую
Private Function ToNum(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim s As String
    s = CellText(v, r, c)
    s = Replace(s, ChrW(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, ",", ".")
    ' отрицательные в скобках: (1 234,56)
    If Len(s) > 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If
    If Len(s) > 0 And IsNumeric(s) Then ToNum = Format$(CDbl(s), "0.############")
End Function

' Период вида "26.08.2026 - 25.09.2026" -> ключ "09.2026" и дата окончания
Private Function PeriodKey(ByVal s As String, ByRef endDate As Date) As String
    Dim parts() As String, d2 As Date
    If InStr(s, "-") = 0 Then Exit Function
    parts = Split(s, " - ")
    If UBound(parts) < 1 Then parts = Split(s, "-")
    If UBound(parts) < 1 Then Exit Function
    d2 = ParseDMY(Trim$(parts(UBound(parts))))
    If d2 = 0 Then Exit Function
    endDate = d2
    PeriodKey = Format$(d2, "mm.yyyy")
End Function

Private Function ParseDMY(ByVal s As String) As Date
    Dim d As Long, mo As Long, y As Long
    If Len(s) < 10 Then Exit Function
    If Not IsNumeric(Left$(s, 2)) Or Not IsNumeric(Mid$(s, 4, 2)) Then Exit Function
    If Not IsNumeric(Mid$(s, 7, 4)) Then Exit Function
    d = CLng(Left$(s, 2))
    mo = CLng(Mid$(s, 4, 2))
    y = CLng(Mid$(s, 7, 4))
    If d < 1 Or d > 31 Or mo < 1 Or mo > 12 Or y < 1990 Then Exit Function
    ParseDMY = DateSerial(y, mo, d)
End Function

'=====================================================================
'  СБОРКА КНИГИ
'=====================================================================
Private Sub BuildBook(ByVal reports As Collection)
    Dim wb As Workbook, ws As Worksheet, prev As Worksheet
    Dim services() As String, periods() As String
    Dim nS As Long, nP As Long, i As Long
    Dim dSvc As Object, dSvcF As Object, dComp As Object
    Dim rep As Variant, it As Variant, k As Variant, kk As String, cur As Variant
    Dim repItems As Collection, fl As String

    CollectAxes reports, services, nS, periods, nP
    mNSvc = nS
    mNPers = nP

    Set dSvc = CreateObject("Scripting.Dictionary")
    Set dSvcF = CreateObject("Scripting.Dictionary")
    Set dComp = CreateObject("Scripting.Dictionary")

    For i = 1 To reports.Count
        rep = reports(i)
        kk = rep(2) & "|" & rep(4)
        If dSvc.Exists(kk) Then
            dSvc(kk) = CDbl(dSvc(kk)) + CDbl(rep(9))
            dSvcF(kk) = dSvcF(kk) & "; " & rep(0)
        Else
            dSvc(kk) = CDbl(rep(9))
            dSvcF(kk) = CStr(rep(0))
        End If
        ' услуга x период x компонент -> сумма стоимости, кол-ва, файлы
        Set repItems = rep(10)
        For k = 1 To repItems.Count
            it = repItems(CLng(k))
            kk = rep(2) & "|" & rep(4) & "|" & CStr(it(0))
            If dComp.Exists(kk) Then
                cur = dComp(kk)
                fl = CStr(cur(2))
                If InStr(1, ";" & fl & ";", ";" & CStr(rep(0)) & ";", vbTextCompare) = 0 Then
                    If Len(fl) > 0 Then fl = fl & "; "
                    fl = fl & CStr(rep(0))
                End If
                dComp(kk) = Array(CDbl(cur(0)) + CDbl(it(5)), CDbl(cur(1)) + CDbl(it(3)), fl)
            Else
                dComp(kk) = Array(CDbl(it(5)), CDbl(it(3)), CStr(rep(0)))
            End If
        Next k
    Next i

    ' --- новая книга с шестью листами ---
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    Set ws = wb.Worksheets(1)
    ws.Name = SH_SUM
    Set prev = ws
    Dim names As Variant
    names = Array(SH_PCT, SH_DIFF, SH_CHART, SH_DET, SH_REG)
    For i = 0 To UBound(names)
        Set ws = wb.Worksheets.Add(After:=prev)
        ws.Name = CStr(names(i))
        Set prev = ws
    Next i

    Dim firstRow As Long, lastRow As Long, dCol As Long, pCol As Long
    firstRow = 8
    lastRow = firstRow + nS
    dCol = 2 + nP
    pCol = dCol + 1

    WriteSummary wb.Worksheets(SH_SUM), dSvc, services, nS, periods, nP, _
                firstRow, lastRow, dCol, pCol
    WritePctSheet wb.Worksheets(SH_PCT), services, nS, periods, nP, firstRow
    WriteDiff wb.Worksheets(SH_DIFF), dSvc, dSvcF, dComp, services, nS, periods, nP
    WriteCharts wb.Worksheets(SH_CHART), services, nS, periods, nP, firstRow, lastRow
    WriteDetail wb.Worksheets(SH_DET), reports, periods, nP
    WriteRegistry wb.Worksheets(SH_REG), reports, periods, nP

    If Len(Dir(mOutPath)) > 0 Then Kill mOutPath
    wb.SaveAs Filename:=mOutPath, FileFormat:=51
    wb.Close SaveChanges:=False
End Sub

Private Sub CollectAxes(ByVal reports As Collection, ByRef services() As String, ByRef nS As Long, _
                        ByRef periods() As String, ByRef nP As Long)
    Dim seenS As Object, seenP As Object
    Dim rep As Variant, i As Long, k As Variant
    Dim tmpS() As String, tmpP() As String

    Set seenS = CreateObject("Scripting.Dictionary")
    Set seenP = CreateObject("Scripting.Dictionary")
    For i = 1 To reports.Count
        rep = reports(i)
        If Not seenS.Exists(CStr(rep(2))) Then seenS(CStr(rep(2))) = 1
        If Not seenP.Exists(CStr(rep(4))) Then seenP(CStr(rep(4))) = rep(5)
    Next i

    ReDim tmpS(0 To seenS.Count - 1)
    i = 0
    For Each k In seenS.Keys
        tmpS(i) = CStr(k)
        i = i + 1
    Next k
    SortTextArray tmpS
    ReDim services(0 To seenS.Count - 1)
    For i = 0 To seenS.Count - 1
        services(i) = tmpS(i)
    Next i
    nS = seenS.Count

    ReDim tmpP(0 To seenP.Count - 1)
    i = 0
    For Each k In seenP.Keys
        tmpP(i) = CStr(k)
        i = i + 1
    Next k
    SortPeriodArray tmpP
    ReDim periods(0 To seenP.Count - 1)
    For i = 0 To seenP.Count - 1
        periods(i) = tmpP(i)
    Next i
    nP = seenP.Count
End Sub

Private Sub SortTextArray(ByRef a() As String)
    Dim i As Long, j As Long, t As String
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If StrComp(a(i), a(j), vbTextCompare) > 0 Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next j
    Next i
End Sub

' сортировка периодов mm.yyyy по дате
Private Sub SortPeriodArray(ByRef a() As String)
    Dim i As Long, j As Long, t As String, ki As Long, kj As Long
    For i = LBound(a) To UBound(a) - 1
        For j = i + 1 To UBound(a)
            ki = CLng(Mid$(a(i), 5, 4)) * 100 + CLng(Left$(a(i), 2))
            kj = CLng(Mid$(a(j), 5, 4)) * 100 + CLng(Left$(a(j), 2))
            If ki > kj Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next j
    Next i
End Sub

'=====================================================================
'  ЛИСТ 1. СВОДКА
'=====================================================================
Private Sub WriteSummary(ws As Worksheet, ByVal dSvc As Object, _
                         ByRef services() As String, ByVal nS As Long, _
                         ByRef periods() As String, ByVal nP As Long, ByVal firstRow As Long, _
                         ByVal lastRow As Long, ByVal dCol As Long, ByVal pCol As Long)
    Dim i As Long, pi As Long, r As Long
    Dim colP0 As Long, colPn As Long, lastCol As Long
    Dim L As String, P As String, L2 As String, LC As String, LP As String
    Dim Lcur As String, prevL As String, key As String
    colP0 = 2
    colPn = 1 + nP
    lastCol = IIf(nP >= 2, pCol, colPn)

    Title ws, "Сводный отчет по услугам", _
          "Источник: папка " & LAQ() & mFolder & RAQ() & " - услуг: " & nS & _
          " - периодов: " & nP & " - сформировано " & Format$(Now, "dd.mm.yyyy hh:nn") & _
          " - красная ячейка: рост стоимости к предыдущему месяцу, зеленая: снижение"

    ws.Range("A4").Value = "Порог подсветки роста, %:"
    ws.Range("A4").Font.Bold = True
    With ws.Range("B4")
        .Value = 0
        .NumberFormat = FM_THR
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(201, 212, 222)
    End With
    ws.Range("D4").Value = "Услуг в реестре:"
    ws.Range("D4").Font.Bold = True
    ws.Range("D4").Font.Color = CLR_TEXT
    ws.Range("E4").Value = nS
    ws.Range("E4").Font.Bold = True
    ws.Range("D5").Value = "Периодов:"
    ws.Range("D5").Font.Bold = True
    ws.Range("D5").Font.Color = CLR_TEXT
    ws.Range("E5").Value = nP
    ws.Range("E5").Font.Bold = True

    ' --- шапка ---
    ws.Cells(7, 1).Value = "Услуга"
    For pi = 0 To nP - 1
        With ws.Cells(7, colP0 + pi)
            ' текстовый формат обязателен: иначе Excel распознает
            ' "09.2025" как дату и покажет "9.2025"
            .NumberFormat = "@"
            .Value = periods(pi)
        End With
    Next pi
    If nP >= 2 Then
        ws.Cells(7, dCol).Value = DLT() & " к пред. периоду, " & RUB() & vbLf & _
                                  "(" & periods(nP - 2) & " " & ARW() & " " & periods(nP - 1) & ")"
        ws.Cells(7, pCol).Value = DLT() & ", %" & vbLf & _
                                  "(" & periods(nP - 2) & " " & ARW() & " " & periods(nP - 1) & ")"
    End If
    StyleHeader ws.Range(ws.Cells(7, 1), ws.Cells(7, lastCol))

    ' --- строки услуг ---
    For i = 0 To nS - 1
        r = firstRow + i
        ws.Cells(r, 1).Value = services(i)
        SetWrap ws.Cells(r, 1)
        For pi = 0 To nP - 1
            key = services(i) & "|" & periods(pi)
            If dSvc.Exists(key) Then
                ws.Cells(r, colP0 + pi).Value = Round(CDbl(dSvc(key)), 2)
                ws.Cells(r, colP0 + pi).NumberFormat = FM_MONEY
            Else
                ws.Cells(r, colP0 + pi).Interior.Color = RGB(242, 242, 242)
            End If
        Next pi
        If nP >= 2 Then
            L = ColLetter(colPn) & r
            P = ColLetter(colPn - 1) & r
            ws.Cells(r, dCol).Formula = _
                "=IF(OR(" & L & "=""""," & P & "=""""),""""," & L & "-" & P & ")"
            ws.Cells(r, dCol).NumberFormat = FM_MONEY
            ws.Cells(r, pCol).Formula = _
                "=IF(OR(" & L & "=""""," & P & "=""""," & P & "=0),""""," & _
                "(" & L & "-" & P & ")/" & P & ")"
            ws.Cells(r, pCol).NumberFormat = FM_PCT
        End If
        FitRow ws, r, services(i), 58
    Next i

    ' --- ИТОГО ---
    ws.Cells(lastRow, 1).Value = "ИТОГО"
    For pi = 0 To nP - 1
        L2 = ColLetter(colP0 + pi)
        ws.Cells(lastRow, colP0 + pi).Formula = _
            "=SUM(" & L2 & firstRow & ":" & L2 & (lastRow - 1) & ")"
        ws.Cells(lastRow, colP0 + pi).NumberFormat = FM_MONEY
    Next pi
    If nP >= 2 Then
        LC = ColLetter(colPn) & lastRow
        LP = ColLetter(colPn - 1) & lastRow
        ws.Cells(lastRow, dCol).Formula = _
            "=IF(OR(" & LC & "=""""," & LP & "=""""),""""," & LC & "-" & LP & ")"
        ws.Cells(lastRow, dCol).NumberFormat = FM_MONEY
        ws.Cells(lastRow, pCol).Formula = _
            "=IF(OR(" & LC & "=""""," & LP & "=""""," & LP & "=0),""""," & _
            "(" & LC & "-" & LP & ")/" & LP & ")"
        ws.Cells(lastRow, pCol).NumberFormat = FM_PCT
    End If
    With ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, lastCol))
        .Interior.Color = CLR_MAIN
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ThinBorders ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol))

    ' --- подсветка месяц к месяцу ---
    If nP >= 2 Then
        prevL = ColLetter(colP0)
        For pi = 1 To nP - 1
            Lcur = ColLetter(colP0 + pi)
            ' диапазон без строки ИТОГО: её синяя заливка должна остаться
            AddCF ws.Range(ws.Cells(firstRow, colP0 + pi), ws.Cells(lastRow - 1, colP0 + pi)), _
                  "=AND(ISNUMBER(" & Lcur & firstRow & "),ISNUMBER(" & prevL & firstRow & ")," & _
                  prevL & firstRow & ">0," & Lcur & firstRow & ">" & prevL & firstRow & "*(1+$B$4))", _
                  RGB(255, 199, 206), RGB(156, 0, 6)
            AddCF ws.Range(ws.Cells(firstRow, colP0 + pi), ws.Cells(lastRow - 1, colP0 + pi)), _
                  "=AND(ISNUMBER(" & Lcur & firstRow & "),ISNUMBER(" & prevL & firstRow & ")," & _
                  prevL & firstRow & ">0," & Lcur & firstRow & "<" & prevL & firstRow & "*(1-$B$4))", _
                  RGB(198, 239, 206), RGB(0, 97, 0)
            prevL = Lcur
        Next pi
        ' для Δ-колонок ISNUMBER обязателен: пустая ячейка - это текст "",
        ' а в Excel любой текст больше любого числа ("" > 0 истинно)
        AddCF ws.Range(ws.Cells(firstRow, dCol), ws.Cells(lastRow - 1, dCol)), _
              "=AND(ISNUMBER(" & ColLetter(dCol) & firstRow & ")," & ColLetter(dCol) & firstRow & ">0)", _
              RGB(255, 199, 206), RGB(156, 0, 6)
        AddCF ws.Range(ws.Cells(firstRow, dCol), ws.Cells(lastRow - 1, dCol)), _
              "=AND(ISNUMBER(" & ColLetter(dCol) & firstRow & ")," & ColLetter(dCol) & firstRow & "<0)", _
              RGB(198, 239, 206), RGB(0, 97, 0)
        AddCF ws.Range(ws.Cells(firstRow, pCol), ws.Cells(lastRow - 1, pCol)), _
              "=AND(ISNUMBER(" & ColLetter(pCol) & firstRow & ")," & ColLetter(pCol) & firstRow & ">0)", _
              RGB(255, 199, 206), RGB(156, 0, 6)
        AddCF ws.Range(ws.Cells(firstRow, pCol), ws.Cells(lastRow - 1, pCol)), _
              "=AND(ISNUMBER(" & ColLetter(pCol) & firstRow & ")," & ColLetter(pCol) & firstRow & "<0)", _
              RGB(198, 239, 206), RGB(0, 97, 0)
    End If

    ' --- оформление ---
    ws.Columns("A").ColumnWidth = 58
    For pi = 0 To nP - 1
        ws.Columns(ColLetter(colP0 + pi)).ColumnWidth = 16
    Next pi
    If nP >= 2 Then
        ws.Columns(ColLetter(dCol)).ColumnWidth = 20
        ws.Columns(ColLetter(pCol)).ColumnWidth = 14
    End If
    Freeze ws, "B" & firstRow
    On Error Resume Next
    ws.Range(ws.Cells(7, 1), ws.Cells(lastRow - 1, lastCol)).AutoFilter
    On Error GoTo 0
    PageSetupLandscape ws
End Sub

'=====================================================================
'  ЛИСТ 2. % ИЗМЕНЕНИЙ
'=====================================================================
Private Sub WritePctSheet(ws As Worksheet, ByRef services() As String, ByVal nS As Long, _
                          ByRef periods() As String, ByVal nP As Long, ByVal firstRow As Long)
    Dim i As Long, pi As Long, r As Long, colP0 As Long, lastCol As Long
    Dim L As String, P As String, sr As Long
    colP0 = 2
    lastCol = colP0 + nP - 1

    Title ws, "Изменение стоимости к предыдущему периоду, %", _
          "Красный - рост, зеленый - снижение. Пусто - нет данных за предыдущий период " & _
          "(первый месяц или новая услуга)."

    ws.Cells(4, 1).Value = "Услуга"
    For pi = 0 To nP - 1
        With ws.Cells(4, colP0 + pi)
            .NumberFormat = "@"
            .Value = periods(pi)
        End With
    Next pi
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, lastCol))

    For i = 0 To nS - 1
        r = 5 + i
        ws.Cells(r, 1).Value = services(i)
        SetWrap ws.Cells(r, 1)
        sr = firstRow + i
        For pi = 1 To nP - 1
            L = ColLetter(colP0 + pi)
            P = ColLetter(colP0 + pi - 1)
            ws.Cells(r, colP0 + pi).Formula = _
                "=IF(OR('" & SH_SUM & "'!" & L & sr & "=""""," & _
                "'" & SH_SUM & "'!" & P & sr & "=""""," & _
                "'" & SH_SUM & "'!" & P & sr & "=0),""""," & _
                "('" & SH_SUM & "'!" & L & sr & "-'" & SH_SUM & "'!" & P & sr & ")/" & _
                "'" & SH_SUM & "'!" & P & sr & ")"
            ws.Cells(r, colP0 + pi).NumberFormat = FM_PCT
        Next pi
        FitRow ws, r, services(i), 58
    Next i

    ThinBorders ws.Range(ws.Cells(5, 1), ws.Cells(4 + nS, lastCol))
    If nP >= 2 Then
        AddCF ws.Range(ws.Cells(5, colP0 + 1), ws.Cells(4 + nS, lastCol)), _
              "=AND(ISNUMBER(" & ColLetter(colP0 + 1) & "5)," & ColLetter(colP0 + 1) & "5>0)", _
              RGB(255, 199, 206), RGB(156, 0, 6)
        AddCF ws.Range(ws.Cells(5, colP0 + 1), ws.Cells(4 + nS, lastCol)), _
              "=AND(ISNUMBER(" & ColLetter(colP0 + 1) & "5)," & ColLetter(colP0 + 1) & "5<0)", _
              RGB(198, 239, 206), RGB(0, 97, 0)
    End If

    ws.Columns("A").ColumnWidth = 58
    For pi = 0 To nP - 1
        ws.Columns(ColLetter(colP0 + pi)).ColumnWidth = 14
    Next pi
    Freeze ws, "B5"
    PageSetupLandscape ws
End Sub

'=====================================================================
'  ЛИСТ 3. РАСХОЖДЕНИЯ
'=====================================================================
Private Sub WriteDiff(ws As Worksheet, ByVal dSvc As Object, ByVal dSvcF As Object, _
                      ByVal dComp As Object, ByRef services() As String, ByVal nS As Long, _
                      ByRef periods() As String, ByVal nP As Long)
    Dim cols As Variant
    cols = Array("Уровень", "Услуга", "Компонент", "Пред. период", "Текущий период", _
                 "Было, " & RUB(), "Стало, " & RUB(), DLT() & ", " & RUB(), DLT() & ", %", _
                 "Тариф было", "Тариф стало", DLT() & " тарифа, %", _
                 "Кол-во было", "Кол-во стало", "Тип", "Файл отчета")
    Dim nCols As Long
    nCols = UBound(cols) + 1

    Title ws, "Расхождения стоимости между периодами", _
          "Услуга - изменение итога по услуге; Компонент - изменение по строке отчета. " & _
          LAQ() & "Новая" & RAQ() & " - компонент или услуга появился впервые, " & _
          LAQ() & "Пропала" & RAQ() & " - был в прошлом периоде."

    Dim rows As Collection, r As Long, i As Long, pi As Long
    Dim pPrev As String, pCur As String, s As String, kPrev As String, kCur As String
    Dim hasPrev As Boolean, hasCur As Boolean
    Dim cp As Double, cc As Double, dd As Double, pc As Variant
    Set rows = New Collection

    For pi = 1 To nP - 1
        pPrev = periods(pi - 1)
        pCur = periods(pi)
        For i = 0 To nS - 1
            s = services(i)
            kPrev = s & "|" & pPrev
            kCur = s & "|" & pCur
            hasPrev = dSvc.Exists(kPrev)
            hasCur = dSvc.Exists(kCur)
            If hasPrev Or hasCur Then
                If Not hasPrev Then
                    rows.Add Array(0, s, "", pPrev, pCur, Empty, Round(CDbl(dSvc(kCur)), 2), _
                                   Empty, Empty, Empty, Empty, Empty, Empty, Empty, _
                                   "Новая услуга", FilesOf(dSvcF, kCur))
                ElseIf Not hasCur Then
                    rows.Add Array(0, s, "", pPrev, pCur, Round(CDbl(dSvc(kPrev)), 2), 0#, _
                                   Round(-CDbl(dSvc(kPrev)), 2), -1#, Empty, Empty, Empty, _
                                   Empty, Empty, "Пропала", FilesOf(dSvcF, kPrev))
                Else
                    cp = CDbl(dSvc(kPrev))
                    cc = CDbl(dSvc(kCur))
                    dd = cc - cp
                    If Abs(dd) >= 0.005 Then
                        pc = Empty
                        If cp <> 0 Then pc = dd / cp
                        rows.Add Array(0, s, "", pPrev, pCur, Round(cp, 2), Round(cc, 2), _
                                       Round(dd, 2), pc, Empty, Empty, Empty, Empty, Empty, _
                                       IIf(dd > 0, "Рост", "Снижение"), FilesOf(dSvcF, kCur))
                    End If
                    AddCompRows rows, dComp, s, pPrev, pCur, kPrev, kCur
                End If
            End If
        Next i
    Next pi

    mNDisc = rows.Count

    Dim sorted As Variant, row0 As Variant, c As Long, v As Variant
    sorted = SortDiff(rows)

    For c = 0 To UBound(cols)
        ws.Cells(4, c + 1).Value = cols(c)
    Next c
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, nCols))

    If Not IsEmpty(sorted) Then
        For i = 0 To UBound(sorted)
            r = 5 + i
            row0 = sorted(i)
            For c = 0 To nCols - 1
                v = row0(c)
                If c = 0 Then
                    ' колонка "Уровень" - текстом, как в эталоне
                    ws.Cells(r, 1).Value = IIf(CLng(v) = 0, "Услуга", "Компонент")
                ElseIf IsEmpty(v) Then
                    ' оставляем пустым
                ElseIf c <= 4 Or c = 14 Or c = 15 Then
                    ws.Cells(r, c + 1).Value = CStr(v)
                Else
                    ws.Cells(r, c + 1).Value = CDbl(v)
                End If
            Next c
            ws.Cells(r, 1).Font.Bold = True
            ws.Cells(r, 1).Font.Color = IIf(CLng(row0(0)) = 0, CLR_MAIN, CLR_TEXT)
            ws.Cells(r, 15).Font.Bold = True
            Select Case CStr(row0(14))
                Case "Рост", "Новая", "Новая услуга"
                    ws.Cells(r, 15).Font.Color = RGB(156, 0, 6)
                Case "Снижение", "Пропала"
                    ws.Cells(r, 15).Font.Color = RGB(0, 97, 0)
            End Select
            SetWrap ws.Cells(r, 16)

            ws.Range(ws.Cells(r, 6), ws.Cells(r, 8)).NumberFormat = FM_MONEY
            ws.Cells(r, 9).NumberFormat = FM_PCT
            ws.Range(ws.Cells(r, 10), ws.Cells(r, 11)).NumberFormat = FM_MONEY
            ws.Cells(r, 12).NumberFormat = FM_PCT
            ws.Range(ws.Cells(r, 13), ws.Cells(r, 14)).NumberFormat = FM_QTY

            FitRow ws, r, CStr(row0(1)), 46, CStr(row0(2)), 46, CStr(row0(15)), 40
        Next i

        Dim lastR As Long
        lastR = 4 + (UBound(sorted) + 1)
        ThinBorders ws.Range(ws.Cells(5, 1), ws.Cells(lastR, nCols))
        AddCF ws.Range(ws.Cells(5, 8), ws.Cells(lastR, 8)), "=AND(ISNUMBER(H5),H5>0)", RGB(255, 199, 206), RGB(156, 0, 6)
        AddCF ws.Range(ws.Cells(5, 8), ws.Cells(lastR, 8)), "=AND(ISNUMBER(H5),H5<0)", RGB(198, 239, 206), RGB(0, 97, 0)
        AddCF ws.Range(ws.Cells(5, 9), ws.Cells(lastR, 9)), "=AND(ISNUMBER(I5),I5>0)", RGB(255, 199, 206), RGB(156, 0, 6)
        AddCF ws.Range(ws.Cells(5, 9), ws.Cells(lastR, 9)), "=AND(ISNUMBER(I5),I5<0)", RGB(198, 239, 206), RGB(0, 97, 0)
        ws.Range(ws.Cells(4, 1), ws.Cells(lastR, nCols)).AutoFilter
    End If

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 46
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 13
    ws.Columns("E").ColumnWidth = 15
    ws.Range("F:H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 11
    ws.Range("J:K").ColumnWidth = 13
    ws.Range("L:N").ColumnWidth = 13
    ws.Columns("O").ColumnWidth = 14
    ws.Columns("P").ColumnWidth = 40
    Freeze ws, "C5"
    PageSetupLandscape ws
End Sub

Private Sub AddCompRows(ByRef rows As Collection, ByVal dComp As Object, ByVal s As String, _
                        ByVal pPrev As String, ByVal pCur As String, _
                        ByVal kPrev As String, ByVal kCur As String)
    Dim names As Object, k As Variant, nm As String, key As String
    Dim prePrev As String, preCur As String
    Dim cp As Variant, cc As Variant, dd As Variant, pc As Variant
    Dim qp As Variant, qc As Variant, tp As Variant, tc As Variant, dt As Variant
    Dim typ As String, files As String, i As Long
    Dim rawCP As Double, rawQP As Double, rawTP As Double, fP As String
    Dim rawCC As Double, rawQC As Double, rawTC As Double, fC As String
    Dim hasP As Boolean, hasC As Boolean
    Dim arr() As String

    If dComp.Count = 0 Then Exit Sub
    Set names = CreateObject("Scripting.Dictionary")
    prePrev = kPrev & "|"
    preCur = kCur & "|"
    For Each k In dComp.Keys
        If Left$(CStr(k), Len(prePrev)) = prePrev Then names(Mid$(CStr(k), Len(prePrev) + 1)) = 1
        If Left$(CStr(k), Len(preCur)) = preCur Then names(Mid$(CStr(k), Len(preCur) + 1)) = 1
    Next k
    If names.Count = 0 Then Exit Sub

    ReDim arr(0 To names.Count - 1)
    i = 0
    For Each k In names.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k
    SortTextArray arr

    For i = 0 To names.Count - 1
        nm = arr(i)
        hasP = CompData(dComp, kPrev & "|" & nm, rawCP, rawQP, rawTP, fP)
        hasC = CompData(dComp, kCur & "|" & nm, rawCC, rawQC, rawTC, fC)

        If Not hasP Then
            cp = Empty
            qp = Empty
            tp = 0
            cc = Round(rawCC, 2)
            qc = rawQC
            tc = rawTC
            typ = "Новая"
            dd = Empty
            pc = Empty
        ElseIf Not hasC Then
            cp = Round(rawCP, 2)
            qp = rawQP
            tp = rawTP
            cc = 0#
            qc = Empty
            tc = 0
            typ = "Пропала"
            dd = Round(-rawCP, 2)
            pc = -1#
        Else
            cp = Round(rawCP, 2)
            cc = Round(rawCC, 2)
            qp = rawQP
            qc = rawQC
            tp = rawTP
            tc = rawTC
            dd = Round(rawCC - rawCP, 2)
            If Abs(dd) < 0.005 Then GoTo NextComp
            typ = IIf(dd > 0, "Рост", "Снижение")
            pc = Empty
            If rawCP <> 0 Then pc = dd / rawCP
        End If

        dt = Empty
        If tp <> 0 And tc <> 0 Then dt = (tc - tp) / tp

        If hasC Then
            files = fC
        Else
            files = fP
        End If

        rows.Add Array(1, s, nm, pPrev, pCur, cp, cc, dd, pc, _
                      IIf(tp = 0, Empty, Round(tp, 2)), _
                      IIf(tc = 0, Empty, Round(tc, 2)), dt, qp, qc, typ, files)
NextComp:
    Next i
End Sub

' Достает агрегат компонента: стоимость, количество, файлы (тариф = стоимость/кол-во)
Private Function CompData(ByVal d As Object, ByVal key As String, ByRef outCost As Double, _
                          ByRef outQty As Double, ByRef outTariff As Double, _
                          ByRef outFiles As String) As Boolean
    Dim v As Variant
    If Not d.Exists(key) Then Exit Function
    v = d(key)
    outCost = CDbl(v(0))
    outQty = CDbl(v(1))
    outTariff = 0
    If outQty <> 0 Then outTariff = outCost / outQty
    outFiles = CStr(v(2))
    CompData = True
End Function

Private Function FilesOf(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then FilesOf = CStr(d(key)) Else FilesOf = ""
End Function

' сортировка: сначала уровень, затем тип, затем по модулю дельты
Private Function SortDiff(ByVal rows As Collection) As Variant
    Dim n As Long, i As Long, j As Long, t As Variant
    n = rows.Count
    If n = 0 Then
        SortDiff = Empty
        Exit Function
    End If
    Dim arr() As Variant
    ReDim arr(0 To n - 1)
    For i = 1 To n
        arr(i - 1) = rows(i)
    Next i

    Dim keyA As Long, keyB As Long
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            keyA = SortKey(arr(i))
            keyB = SortKey(arr(j))
            If keyB < keyA Or (keyA = keyB And AbsD(arr(j)) > AbsD(arr(i))) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    SortDiff = arr
End Function

Private Function SortKey(ByVal row0 As Variant) As Long
    Dim lv As Long, tp As Long
    lv = CLng(row0(0))
    Select Case CStr(row0(14))
        Case "Рост": tp = 0
        Case "Новая услуга": tp = 1
        Case "Новая": tp = 2
        Case "Снижение": tp = 3
        Case Else: tp = 4
    End Select
    SortKey = lv * 10 + tp
End Function

Private Function AbsD(ByVal row0 As Variant) As Double
    If IsEmpty(row0(7)) Then AbsD = 0 Else AbsD = Abs(CDbl(row0(7)))
End Function

'=====================================================================
'  ЛИСТ 4. ГРАФИКИ
'=====================================================================
Private Sub WriteCharts(ws As Worksheet, ByRef services() As String, ByVal nS As Long, _
                        ByRef periods() As String, ByVal nP As Long, _
                        ByVal firstRow As Long, ByVal lastRow As Long)
    Dim i As Long, pi As Long, r As Long, colP0 As Long, lastCol As Long
    Dim shp As Shape, ch As Chart, wsSum As Worksheet
    Dim L As String, idx As String, sLast As String
    If nS = 0 Or nP = 0 Then Exit Sub
    colP0 = 2
    lastCol = 1 + nP
    sLast = firstRow + nS - 1

    Title ws, "Графики по услугам", _
          "Выберите услуги в выпадающих списках - диаграмма построится автоматически. " & _
          "Ниже - общая динамика стоимости всех услуг по месяцам."

    ws.Range("A4").Value = "Услуга 1:"
    ws.Range("A4").Font.Bold = True
    ws.Range("A5").Value = "Услуга 2:"
    ws.Range("A5").Font.Bold = True
    ws.Range("A6").Value = "Услуга 3:"
    ws.Range("A6").Font.Bold = True
    For i = 1 To 3
        r = 3 + i
        ws.Cells(r, 2).Value = services(IIf(i - 1 < nS, i - 1, 0))
        With ws.Cells(r, 2)
            .Interior.Color = RGB(255, 242, 204)
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(201, 212, 222)
        End With
    Next i

    ' скрытый список услуг для выпадающих списков
    ws.Cells(1, 20).Value = "Список услуг (для выпадающего списка)"
    ws.Cells(1, 20).Font.Size = 9
    ws.Cells(1, 20).Font.Color = CLR_TEXT
    For i = 0 To nS - 1
        ws.Cells(2 + i, 20).Value = services(i)
    Next i
    ws.Columns("T").Hidden = True

    On Error Resume Next
    ws.Range("B4:B6").Validation.Delete
    With ws.Range("B4:B6").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=$T$2:$T$" & (1 + nS)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Услуга"
        .ErrorMessage = "Выберите услугу из списка"
    End With
    On Error GoTo 0

    ' вспомогательная таблица для диаграмм
    ws.Cells(8, 1).Value = "Период"
    ws.Cells(8, 1).Font.Bold = True
    ws.Cells(8, 1).Font.Color = CLR_TEXT
    For pi = 0 To nP - 1
        ws.Cells(8, colP0 + pi).Formula = "='" & SH_SUM & "'!" & ColLetter(colP0 + pi) & "7"
        ws.Cells(8, colP0 + pi).Font.Bold = True
        ws.Cells(8, colP0 + pi).Font.Color = CLR_TEXT
    Next pi

    For i = 1 To 3
        r = 8 + i
        ws.Cells(r, 1).Formula = "=IF($B$" & (3 + i) & "="""","""",$B$" & (3 + i) & ")"
        ws.Cells(r, 1).Font.Bold = True
        For pi = 0 To nP - 1
            L = ColLetter(colP0 + pi)
            idx = "INDEX('" & SH_SUM & "'!$" & ColLetter(2) & "$" & firstRow & ":$" & _
                  ColLetter(1 + nP) & "$" & sLast & ",MATCH($A" & r & ",'" & SH_SUM & "'!$A$" & _
                  firstRow & ":$A$" & sLast & ",0),MATCH(" & L & "$8,'" & SH_SUM & "'!$" & _
                  ColLetter(2) & "$7:$" & ColLetter(1 + nP) & "$7,0))"
            ws.Cells(r, colP0 + pi).Formula = "=IFERROR(IF(" & idx & "="""",NA()," & idx & "),NA())"
            ws.Cells(r, colP0 + pi).NumberFormat = FM_MONEY
        Next pi
    Next i

    ws.Columns("A").ColumnWidth = 56
    ws.Columns("B").ColumnWidth = 18
    For pi = 1 To nP - 1
        ws.Columns(ColLetter(colP0 + pi)).ColumnWidth = 18
    Next pi

    ' --- диаграмма 1: выбранные услуги ---
    Set shp = ws.Shapes.AddChart2(-1, xlColumnClustered, ws.Range("A14").Left, _
                                 ws.Range("A14").Top, 690, 340)
    Set ch = shp.Chart
    ClearSeries ch
    For i = 1 To 3
        With ch.SeriesCollection.NewSeries
            .Name = "='" & SH_CHART & "'!$A$" & (8 + i)
            .XValues = ws.Range(ws.Cells(8, colP0), ws.Cells(8, lastCol))
            .Values = ws.Range(ws.Cells(8 + i, colP0), ws.Cells(8 + i, lastCol))
        End With
    Next i
    FinishChart ch, "Стоимость выбранных услуг по периодам, " & RUB() & " (без НДС)"

    ' --- диаграмма 2: общая динамика ---
    Set wsSum = ws.Parent.Worksheets(SH_SUM)
    Set shp = ws.Shapes.AddChart2(-1, xlLine, ws.Range("A42").Left, ws.Range("A42").Top, 690, 340)
    Set ch = shp.Chart
    ClearSeries ch
    With ch.SeriesCollection.NewSeries
        .Name = "ИТОГО по всем услугам"
        .XValues = wsSum.Range(wsSum.Cells(7, colP0), wsSum.Cells(7, lastCol))
        .Values = wsSum.Range(wsSum.Cells(lastRow, colP0), wsSum.Cells(lastRow, lastCol))
    End With
    FinishChart ch, "Общая стоимость всех услуг по периодам, " & RUB() & " (без НДС)"

    PageSetupLandscape ws
End Sub

Private Sub ClearSeries(ch As Chart)
    Dim guard As Long
    Do While ch.SeriesCollection.Count > 0 And guard < 20
        ch.SeriesCollection(1).Delete
        guard = guard + 1
    Loop
End Sub

Private Sub FinishChart(ch As Chart, ByVal ttl As String)
    On Error Resume Next
    ch.HasTitle = True
    ch.ChartTitle.Text = ttl
    If ch.SeriesCollection.Count > 0 Then
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionRight
        ' True = легенда участвует в макете и не перекрывает область
        ' построения (в OOXML это overlay=False, как в эталоне)
        ch.Legend.IncludeInLayout = True
        ch.ChartGroups(1).VaryByCategories = False
    Else
        ch.HasLegend = False
    End If
    ch.Axes(xlCategory).HasTitle = True
    ch.Axes(xlCategory).AxisTitle.Text = "Период"
    ch.Axes(xlCategory).TickLabelSpacing = 1
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.DisplayBlanksAs = xlNotPlotted
    On Error GoTo 0
End Sub

'=====================================================================
'  ЛИСТ 5. ДЕТАЛИЗАЦИЯ
'=====================================================================
Private Sub WriteDetail(ws As Worksheet, ByVal reports As Collection, ByRef periods() As String, _
                        ByVal nP As Long)
    Dim cols As Variant
    cols = Array("Файл", NUM() & " отчета", "Услуга", "Период", "Компонент", _
                 "Тип", "Ед. изм.", "Кол-во", "Тариф, " & RUB(), "Стоимость, " & RUB())
    Dim nCols As Long
    nCols = UBound(cols) + 1
    Title ws, "Детализация по строкам отчетов", "Все компоненты всех разобранных файлов."

    Dim i As Long, c As Long, r As Long, k As Long, it As Variant, idx As Collection
    Dim rep As Variant, repItems As Collection
    For c = 0 To UBound(cols)
        ws.Cells(4, c + 1).Value = cols(c)
    Next c
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, nCols))

    Set idx = SortedIndices(reports, periods, nP)
    r = 5
    For i = 1 To idx.Count
        rep = reports(CLng(idx(i)))
        Set repItems = rep(10)
        For k = 1 To repItems.Count
            it = repItems(CLng(k))
            ws.Cells(r, 1).Value = rep(0)
            ws.Cells(r, 2).Value = rep(1)
            ws.Cells(r, 3).Value = rep(2)
            ws.Cells(r, 4).Value = rep(4)
            ws.Cells(r, 5).Value = it(0)
            ws.Cells(r, 6).Value = it(1)
            ws.Cells(r, 7).Value = it(2)
            ws.Cells(r, 8).Value = it(3)
            ws.Cells(r, 9).Value = it(4)
            ws.Cells(r, 10).Value = it(5)
            ws.Cells(r, 8).NumberFormat = FM_QTY
            ws.Range(ws.Cells(r, 9), ws.Cells(r, 10)).NumberFormat = FM_MONEY
            SetWrap ws.Cells(r, 1)
            FitRow ws, r, CStr(rep(0)), 52, CStr(rep(2)), 46, CStr(it(0)), 52
            r = r + 1
        Next k
    Next i

    Dim lastR As Long
    lastR = r - 1
    If lastR >= 5 Then
        ThinBorders ws.Range(ws.Cells(5, 1), ws.Cells(lastR, nCols))
        ws.Range(ws.Cells(4, 1), ws.Cells(lastR, nCols)).AutoFilter
    End If
    ws.Columns("A").ColumnWidth = 52
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 11
    ws.Columns("E").ColumnWidth = 52
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 13
    ws.Columns("I").ColumnWidth = 13
    ws.Columns("J").ColumnWidth = 16
    Freeze ws, "C5"
    PageSetupLandscape ws
End Sub

' порядок отчетов: период, затем услуга
Private Function SortedIndices(ByVal reports As Collection, ByRef periods() As String, _
                               ByVal nP As Long) As Collection
    Dim i As Long, j As Long, pi As Long, kj As Long
    Dim arr() As Variant, t As Variant, res As Collection
    Dim ri As Variant, rj As Variant
    Set res = New Collection
    If reports.Count = 0 Then
        Set SortedIndices = res
        Exit Function
    End If
    ReDim arr(1 To reports.Count)
    For i = 1 To reports.Count
        arr(i) = i
    Next i
    For i = 1 To reports.Count - 1
        For j = i + 1 To reports.Count
            ri = reports(arr(i))
            rj = reports(arr(j))
            pi = PeriodIdx(CStr(ri(4)), periods, nP)
            kj = PeriodIdx(CStr(rj(4)), periods, nP)
            If kj < pi Or (kj = pi And StrComp(CStr(rj(2)), CStr(ri(2)), vbTextCompare) < 0) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    For i = 1 To reports.Count
        res.Add arr(i)
    Next i
    Set SortedIndices = res
End Function

Private Function PeriodIdx(ByVal key As String, ByRef periods() As String, ByVal nP As Long) As Long
    Dim i As Long
    For i = 0 To nP - 1
        If periods(i) = key Then
            PeriodIdx = i
            Exit Function
        End If
    Next i
    PeriodIdx = 999
End Function

'=====================================================================
'  ЛИСТ 6. РЕЕСТР ИСХОДНЫХ ОТЧЕТОВ
'=====================================================================
Private Sub WriteRegistry(ws As Worksheet, ByVal reports As Collection, ByRef periods() As String, _
                         ByVal nP As Long)
    Dim cols As Variant
    cols = Array("Файл", NUM() & " отчета", "Услуга", "Отчетный период", _
                 "Период (мм.гггг)", "Компонентов", "Итого, " & RUB(), _
                 "Номер договора", "Сервис-менеджер")
    Dim nCols As Long
    nCols = UBound(cols) + 1
    Title ws, "Реестр разобранных файлов", _
          "Каждый файл папки " & LAQ() & mFolder & RAQ() & _
          " - его услуга, период и итоговая сумма."

    Dim i As Long, r As Long, rep As Variant, idx As Collection
    For i = 0 To UBound(cols)
        ws.Cells(4, i + 1).Value = cols(i)
    Next i
    StyleHeader ws.Range(ws.Cells(4, 1), ws.Cells(4, nCols))

    Set idx = SortedIndices(reports, periods, nP)
    r = 5
    For i = 1 To idx.Count
        rep = reports(CLng(idx(i)))
        ws.Cells(r, 1).Value = rep(0)
        ws.Cells(r, 2).Value = rep(1)
        ws.Cells(r, 3).Value = rep(2)
        ws.Cells(r, 4).Value = rep(3)
        ws.Cells(r, 5).Value = rep(4)
        ws.Cells(r, 6).Value = rep(8)
        ws.Cells(r, 7).Value = Round(CDbl(rep(9)), 2)
        ws.Cells(r, 8).Value = rep(6)
        ws.Cells(r, 9).Value = rep(7)
        ws.Cells(r, 7).NumberFormat = FM_MONEY
        SetWrap ws.Cells(r, 1)
        SetWrap ws.Range(ws.Cells(r, 8), ws.Cells(r, 9))
        FitRow ws, r, CStr(rep(0)), 56, CStr(rep(2)), 46, CStr(rep(6)), 40, CStr(rep(7)), 34
        r = r + 1
    Next i

    Dim lastR As Long
    lastR = r - 1
    If lastR >= 5 Then
        ThinBorders ws.Range(ws.Cells(5, 1), ws.Cells(lastR, nCols))
        ws.Range(ws.Cells(4, 1), ws.Cells(lastR, nCols)).AutoFilter
    End If
    ws.Columns("A").ColumnWidth = 56
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 46
    ws.Columns("D").ColumnWidth = 26
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 40
    ws.Columns("I").ColumnWidth = 34
    Freeze ws, "C5"
    PageSetupLandscape ws
End Sub

'=====================================================================
'  ХЕЛПЕРЫ ОФОРМЛЕНИЯ
'=====================================================================
Private Sub SetWrap(rng As Range)
    On Error Resume Next
    rng.WrapText = True
    rng.VerticalAlignment = xlCenter
    On Error GoTo 0
End Sub

Private Sub Title(ws As Worksheet, ByVal t As String, ByVal sub_ As String)
    With ws.Range("A1")
        .Value = t
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = CLR_MAIN
    End With
    With ws.Range("A2")
        .Value = sub_
        .Font.Size = 10
        .Font.Italic = True
        .Font.Color = CLR_TEXT
    End With
End Sub

Private Sub StyleHeader(rng As Range)
    Dim cel As Range, hasNl As Boolean
    With rng
        .Interior.Color = CLR_MAIN
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(201, 212, 222)
    End With
    For Each cel In rng
        If InStr(CStr(cel.Value), Chr(10)) > 0 Then hasNl = True
    Next cel
    rng.EntireRow.RowHeight = IIf(hasNl, 46, 32)
End Sub

Private Sub ThinBorders(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(201, 212, 222)
    End With
End Sub

' Имена функций в формулах условного форматирования.
' Русский Excel (UI 1049) ждёт РУССКИЕ имена: переданные по-английски
' он сохраняет их как _xludf.AND(...) - правило молча не работает.
Private Function CfF(ByVal s As String) As String
    On Error Resume Next
    If Application.LanguageSettings.LanguageID(2) = 1049 Then
        s = Replace(s, "ISNUMBER(", "ЕЧИСЛО(")
        s = Replace(s, "ISBLANK(", "ПУСТО(")
        s = Replace(s, "ISERROR(", "ЕШИБКА(")
        s = Replace(s, "AND(", "И(")
        s = Replace(s, "OR(", "ИЛИ(")
        s = Replace(s, "NOT(", "НЕ(")
    End If
    On Error GoTo 0
    CfF = s
End Function

Private Sub AddCF(rng As Range, ByVal formula As String, ByVal fill As Long, ByVal fontClr As Long)
    Dim fc As FormatCondition
    On Error Resume Next
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=CfF(formula))
    With fc
        .Interior.Color = fill
        .Font.Color = fontClr
        .Font.Bold = True
        .StopIfTrue = False
    End With
    On Error GoTo 0
End Sub

' высота строки с учетом переносов (до 4 пар "текст, ширина колонки")
Private Sub FitRow(ws As Worksheet, ByVal r As Long, _
                   Optional ByVal t1 As String = "", Optional ByVal w1 As Long = 0, _
                   Optional ByVal t2 As String = "", Optional ByVal w2 As Long = 0, _
                   Optional ByVal t3 As String = "", Optional ByVal w3 As Long = 0, _
                   Optional ByVal t4 As String = "", Optional ByVal w4 As Long = 0)
    Dim lines As Long, need As Long
    lines = 1
    need = LinesFor(t1, w1): If need > lines Then lines = need
    need = LinesFor(t2, w2): If need > lines Then lines = need
    need = LinesFor(t3, w3): If need > lines Then lines = need
    need = LinesFor(t4, w4): If need > lines Then lines = need
    If 14 * lines + 4 > 16 Then
        ws.Rows(r).RowHeight = 14 * lines + 4
    Else
        ws.Rows(r).RowHeight = 16
    End If
End Sub

Private Function LinesFor(ByVal t As String, ByVal w As Long) As Long
    If Len(t) = 0 Or w < 4 Then
        LinesFor = 1
    Else
        LinesFor = (Len(t) + w - 3) \ (w - 2)
    End If
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim s As String, n As Long
    n = c
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Sub Freeze(ws As Worksheet, ByVal addr As String)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(addr).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Sub PageSetupLandscape(ws As Worksheet)
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo 0
End Sub