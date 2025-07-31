Attribute VB_Name = "Module1"
Sub Из_Объёмы_от_СМУ()
        Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    If Workbooks.Count Then ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    Dim Prov As Worksheet
    Set Prov = Vb.Worksheets("Проверка")
Vb.Worksheets("Проверка").Activate    ' Переходим на лист "Проверка"
If ActiveWindow.FreezePanes Then ' Проверяем, есть ли закрепленные области
    ActiveWindow.FreezePanes = False    ' Если закрепленные области есть, снимаем их
End If


    Prov.Columns("A:N").ColumnWidth = 5
    Prov.Columns("G").ColumnWidth = 11
    Prov.Columns("O:T").ColumnWidth = 17
    Prov.Columns("P:P").ColumnWidth = 40
    Prov.Rows("1:8").ClearContents
    Prov.Rows("1:10").RowHeight = 10
    Prov.Rows("1:7").RowHeight = 1
    Prov.Rows("1:1").RowHeight = 3
    Prov.Rows("8").RowHeight = 30
    Prov.Rows("9").RowHeight = 15
    Prov.Rows("1:10").HorizontalAlignment = xlCenter
    Prov.Rows("1:10").VerticalAlignment = xlCenter
    Prov.Rows("8:8").VerticalAlignment = xlTop
    Prov.Rows("8:8").WrapText = True
    Prov.Range("D8") = "Ключ"
    Prov.Range("E8") = "Номер пакета"
    Prov.Range("F8") = "Фаза"
    Prov.Range("G8") = "Номер титула"
    Prov.Range("H8") = "Наименование Титула"
    Prov.Range("I8") = "Чертеж"
    Prov.Range("J8") = "Номер структуры"
    Prov.Range("K8") = "Элемент"
    Prov.Range("L8") = "Шифр Единичнной расценки"
    Prov.Range("M8") = "Описание Единичной Расценки"
    Prov.Range("N8") = "Ед изм"
    Prov.Range("Q8") = "Подразделение"
    Prov.Range("P8") = "ФИО Прораба / Мастера"
    Prov.Range("A8:Q8").Interior.Color = RGB(183, 222, 232)
    Prov.Range("A9:Q9").Interior.Color = RGB(218, 238, 243)
    Prov.Range("A10:S10").FormulaLocal = "=СТОЛБЕЦ()"
    Vb.Worksheets("Проверка").Rows("9:9").NumberFormat = "#,##0.00"
    ' Открываем диалоговое окно для выбора файлов
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog3(Vb.Path)
    
    ' Проверяем, был ли выбран файл
    If filePaths.Count = 0 Then
        MsgBox "Файл не выбран!", vbExclamation
        Exit Sub
    End If
       
    'Инициализация счетчиков
    Dim totalBooks As Long
    Dim processedBooks As Long
    processedBooks = 0
    totalBooks = filePaths.Count
    Vb.Worksheets("Главный").Cells(1, "D").Value = "Всего обработано 0 из " & totalBooks & " книг" ' Начальное значение
    
    ' Коллекция для сбора ошибок
    Dim errorMessages As Collection
    Set errorMessages = New Collection
    
    
    
    
    
    ' Обрабатываем каждый выбранный файл
    Dim filePath As Variant
    For Each filePath In filePaths
        ProcessFile CStr(filePath), Vb, Prov, errorMessages
        processedBooks = processedBooks + 1
        Application.ScreenUpdating = True
        Vb.Worksheets("Главный").Cells(1, "D").Value = "Всего обработано " & processedBooks & " из " & totalBooks & " книг"
        Application.ScreenUpdating = False
    Next filePath







    ' Проверка на пустые ячейки в столбцах D, P, Q
    ActiveSheet.UsedRange
    LastRowDPQ = Prov.Cells.SpecialCells(xlLastCell).Row
    Dim errorMessageDPQ As String
    errorMessageDPQ = ""
    ' Проверка столбца D (КЛЮЧ)
    Dim emptyCellFound As Boolean
    emptyCellFound = False
    Dim cellDPQempty As Range
    For Each cellDPQempty In Prov.Range("D11:D" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "В столбце КЛЮЧ есть пустые ячейки." & vbCrLf
    End If
     ' Проверка столбца G (Титул)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("G11:G" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "В столбце ТИТУЛ есть пустые ячейки." & vbCrLf
    End If
    ' Проверка столбца P (ПОДРАЗДЕЛЕНИЕ)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("P11:P" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "В столбце ПОДРАЗДЕЛЕНИЕ есть пустые ячейки." & vbCrLf
    End If
    ' Проверка столбца Q (УПРАВЛЕНИЕ)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("Q11:Q" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "В столбце УПРАВЛЕНИЕ есть пустые ячейки." & vbCrLf
    End If
    ' Вывод общего сообщения, если есть ошибки
    If errorMessageDPQ <> "" Then
        MsgBox "Обнаружены пустые ячейки:" & vbCrLf & vbCrLf & errorMessageDPQ, vbExclamation, "Ошибка"
    End If
    Prov.Range("D11").FormulaLocal = "=ЛЕВСИМВ(СЦЕПИТЬ(I11;L11;K11);190)"
    Prov.Range("D11").AutoFill Destination:=Prov.Range("D11:D" & LastRowDPQ), Type:=xlFillDefault
    
    Prov.Rows("11:" & LastRowDPQ).RowHeight = 15
    Dim sCell As Range
    For Each sCell In Prov.Range("S11:S" & LastRowDPQ)
        If Not IsEmpty(sCell.Value) And Trim(sCell.Value) <> "" Then
            Prov.Range("A" & sCell.Row & ":R" & sCell.Row).Interior.Color = RGB(218, 238, 243)
            Prov.Range("D" & sCell.Row & ":D" & sCell.Row).ClearContents
        End If
    Next sCell
    
    Prov.Range("O8") = "ФО"
    Prov.Range("R8").Value = "ФО" & Chr(10) & "(формула)"
    Prov.Range("S8").Value = "ФО" & Chr(10) & "(до фильтра)"
    Prov.Range("O9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;O11:O" & LastRowDPQ & ")"
    Prov.Range("R9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;R11:R" & LastRowDPQ & ")"
    Prov.Range("S9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;S11:S" & LastRowDPQ & ")"
    Prov.Range("P9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(2;O11:O" & LastRowDPQ & ")"
   Prov.Range("A9:V9").Font.Bold = True
    Prov.Range("A11:S" & LastRowDPQ).WrapText = False
    Dim cellDPQ As Range
    For Each cellDPQ In Prov.Range("P11:P" & LastRowDPQ)
        If cellDPQ.ColumnWidth < Len(cellDPQ.Value) Then
            cellDPQ.WrapText = True
            cellDPQ.EntireRow.AutoFit
        End If
    Next cellDPQ
    Prov.Range("A11:S" & LastRowDPQ).VerticalAlignment = xlCenter
    Prov.Range("A10:U" & LastRowDPQ).AutoFilter
    If errorMessages.Count > 0 Then
        Dim errorMessage As String
        errorMessage = "Обнаружены ошибки при обработке файлов:" & vbCrLf
        Dim i As Integer
        For i = 1 To errorMessages.Count
            errorMessage = errorMessage & errorMessages(i) & vbCrLf
        Next i
        MsgBox errorMessage, vbExclamation, "Ошибки"
    End If
    
 Vb.Worksheets("Проверка").Rows("11:11").Select ' Выбираем строку 11
ActiveWindow.FreezePanes = True ' Закрепляем строки с 1 по 10
    
    
    Prov.Range("O9").Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Workbooks.Count Then
        ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True
    End If
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ProcessFile(filePath As String, Vb As Workbook, Prov As Worksheet, errorMessages As Collection)
    Dim Ob As Workbook
    On Error Resume Next
    Set Ob = Workbooks.Open(filePath)
    On Error GoTo 0
    
    ' Проверяем, успешно ли открылся файл
    If Ob Is Nothing Then
        errorMessages.Add "Не удалось открыть файл: " & filePath
        Exit Sub
    End If
    
    ' Получаем имя файла без пути
    Dim fileName As String
    fileName = Right(Ob.Name, Len(Ob.Name) - InStrRev(Ob.Name, "\") - 1)
    
    ' Создаём объект регулярного выражения для поиска даты в формате ГОД.МЕСЯЦ.ДЕНЬ
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "(\d{4})\.(\d{2})\.(\d{2})" ' Шаблон: GGGG.MM.DD
        .Global = False
        .IgnoreCase = True
    End With
    
    ' Ищем совпадение с шаблоном в имени файла
    Dim match As Object
    If regex.Test(fileName) Then
        Set match = regex.Execute(fileName)
        ' Извлекаем год, месяц и день из найденной даты
        Dim YearPart As String
        Dim Mes As String
        Dim Day As String
        YearPart = match(0).SubMatches(0) ' Год (четыре цифры)
        Mes = match(0).SubMatches(1)     ' Месяц (две цифры)
        Day = match(0).SubMatches(2)     ' День (две цифры)
    Else
        ' Добавляем сообщение об ошибке в коллекцию
        errorMessages.Add "Формат даты в файле " & fileName & " не корректен (ожидается GGGG.MM.DD)."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Проверяем, существует ли лист "Объёмы ООО ""Р-СТРОЙ"""
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""")
    On Error GoTo 0
    ' Если лист не найден, добавляем сообщение об ошибке
    If targetSheet Is Nothing Then
        errorMessages.Add "Лист ""Объёмы ООО ""Р-СТРОЙ"""" не найден в файле " & fileName & "."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Раскрываем автофильтр и скрытые строки/столбцы на всех листах книги Ob
    Dim ws As Worksheet
    For Each ws In Ob.Worksheets
        ' Раскрываем автофильтр, если он активен
        If ws.AutoFilterMode And ws.FilterMode Then ws.ShowAllData
        ' Раскрываем скрытые строки
        If ws.Cells.EntireRow.Hidden Then ws.Cells.EntireRow.Hidden = False
        ' Раскрываем скрытые столбцы
        If ws.Cells.EntireColumn.Hidden Then ws.Cells.EntireColumn.Hidden = False
    Next ws
    
    ' Активируем лист "Объёмы ООО ""Р-СТРОЙ"""
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Activate
    
    ' Поиск номера ячейки, содержащей "ФО за МесяцТочкаДень"
    Dim Rngt As Range
    Set Rngt = Cells.Find("ФО за " & Day & "." & Mes, , xlFormulas, xlWhole)
    If Rngt Is Nothing Then
        errorMessages.Add "В файле " & fileName & " отсутствует запись: ""ФО за " & Day & "." & Mes & """."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lRow As Long
    lRow = Rngt.Row ' Строка, где было найдено слово
    Dim lCol As Long
    lCol = Rngt.Column ' Столбец, где было найдено слово "ФО за"
    
    ' Обновляем UsedRange для листа "Проверка"
    Vb.Worksheets("Проверка").UsedRange
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, lCol).End(xlUp).Row
    Dim sAddress As String
    sAddress = Rngt.Address ' Адрес ячейки, где было найдено слово
    
    ' Проверяем, есть ли значение в ячейке
    Dim volumeValue As Variant
    volumeValue = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(9, lCol).Value
    If IsEmpty(volumeValue) Or volumeValue = 0 Then
        ' Если значение отсутствует или равно 0, добавляем сообщение об ошибке
        errorMessages.Add "В файле " & fileName & " отсутствуют объёмы за отчетный день."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Автофильтр без пустых значений и без нулей в столбце с объёмом за день
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Cells(lRow + 2, lCol), "FF" & LastRow).AutoFilter Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
    
    ' Поиск номера ячейки, содержащей "Подразделение"
    Dim Rngt2 As Range
    Set Rngt2 = Cells.Find("Подразделение", , xlFormulas, xlWhole)
    If Rngt2 Is Nothing Then
        errorMessages.Add "В файле " & fileName & " отсутствует запись: ""Подразделение""."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lCol2 As Long
    lCol2 = Rngt2.Column ' Столбец, где было найдено слово "Подразделение"
    
    ' Скрываем столбцы
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Columns(lCol + 2), Columns(lCol2 - 1)).Hidden = True
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Columns(15), Columns(lCol - 1)).Hidden = True
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Columns("A:C").EntireColumn.Hidden = True
    
    ' Копируем массив в новую книгу
    Application.CutCopyMode = False
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Cells(11, 4), Cells(LastRow, lCol2)).Copy
    Vb.Worksheets("Проверка").Activate
    Vb.Worksheets("Проверка").UsedRange
    Dim lastRow2 As Long
    lastRow2 = Vb.Worksheets("Проверка").Cells.SpecialCells(xlLastCell).Row
    Vb.Worksheets("Проверка").Cells(lastRow2 + 1, 4).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Закрываем книгу Ob
    Ob.Close SaveChanges:=True
    
    ' Обновляем UsedRange для листа "Проверка"
    Vb.Worksheets("Проверка").Activate
    Vb.Worksheets("Проверка").UsedRange
    Dim LastRow3 As Long
    LastRow3 = Vb.Worksheets("Проверка").Cells.SpecialCells(xlLastCell).Row
    
    ' Добавляем формулы и проверки
    Prov.Range("R" & LastRow3 + 1).FormulaLocal = "=СУММ(O" & lastRow2 + 1 & ":O" & LastRow3 & ")"
    Prov.Range("S" & LastRow3 + 1) = volumeValue
    Prov.Range("A" & LastRow3 + 1 & ":R" & LastRow3 + 1).Interior.Color = RGB(218, 238, 243)
    Prov.Rows(LastRow3 + 1).NumberFormat = "#,##0.00"
    
    If Prov.Range("R" & LastRow3 + 1).Value <> Range("S" & LastRow3 + 1).Value Then
        Prov.Range("R" & LastRow3 + 1).Interior.Color = RGB(219, 179, 182)
    End If
End Sub

Function OpenFileDialog3(Optional InitialFolder As String = "") As Collection
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim filePaths As Collection
    Set filePaths = New Collection
    
    With fd
        .Title = "Выберите файл или файлы ""Объёмы ООО Р-СТРОЙ Р3 АГПЗ ОЗХ _..."" по СМУ"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        .AllowMultiSelect = True ' Разрешаем выбор нескольких файлов
        
        ' Указываем начальную папку, если она задана
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                filePaths.Add .SelectedItems(i) ' Добавляем каждый выбранный файл в коллекцию
            Next i
        End If
    End With
    
    Set OpenFileDialog3 = filePaths
End Function

