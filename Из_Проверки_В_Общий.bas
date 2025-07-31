Attribute VB_Name = "Module2"
Sub Из_Проверки_В_Общий()
Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    If Workbooks.Count Then ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ' Объявляем и инициализируем коллекцию для хранения сообщений об ошибках
    Dim errorMessages As Collection
    Set errorMessages = New Collection
Dim Vb As Workbook
    Set Vb = ThisWorkbook
    ' Открываем диалоговое окно для выбора файла
    Dim filePath As Variant
    filePath = OpenFileDialog5(Vb.Path) ' Указываем папку, где находится текущая книга
    ' Проверяем, был ли выбран файл
    If filePath = False Then
        MsgBox "Файл не выбран!", vbExclamation
        Exit Sub
    End If
' Открываем выбранный файл
Dim Ob As Workbook
Set Ob = Workbooks.Open(filePath)
'    ' Проверяем, успешно ли открылся файл
'    If Ob Is Nothing Then
'        errorMessages.Add "Не удалось открыть файл: " & filePath
'        Exit Sub
'    End If
' Проверяем, существует ли лист "Объёмы ООО ""Р-СТРОЙ"""
Dim targetSheet As Worksheet
On Error Resume Next
Set targetSheet = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""")
On Error GoTo 0
If targetSheet Is Nothing Then
    MsgBox "Лист ""Объёмы ООО ""Р-СТРОЙ"""" не найден!", vbExclamation
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

    
                                                                                  ' Проверяем, есть ли значение в ячейке
                                                                                    Dim volumeValue As Variant
                                                                                    volumeValue = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(9, lCol).Value
                                                                                    If Not IsEmpty(volumeValue) And volumeValue <> 0 Then
                                                                                    ' Получаем буквенное обозначение столбца
                                                                                    Dim columnLetter As String
                                                                                    columnLetter = Split(Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(1, lCol).Address, "$")(1)
                                                                                    ' Если значение есть, спрашиваем пользователя, продолжать ли выполнение
                                                                                    Dim userResponse As VbMsgBoxResult
                                                                                    userResponse = MsgBox("В ""Объёмы ООО ""Р-СТРОЙ"" в столбце " & columnLetter & " (ФО) уже есть значения с объёмами." & vbCrLf & _
                                                                                                    "Продолжить выполнение макроса?", vbYesNo + vbExclamation, "Внимание")
                                                                                    ' Если пользователь выбрал "Нет", завершаем выполнение
                                                                                    If userResponse = vbNo Then
                                                                                    Ob.Close SaveChanges:=False
                                                                                    Exit Sub
                                                                                    End If
                                                                                    End If

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

  ' Поиск номера ячейки, содержащей "ФИО Прораба / Мастера"
   Dim Rngt3 As Range
    Set Rngt3 = Cells.Find("ФИО Прораба / Мастера", , xlFormulas, xlWhole)
    If Rngt3 Is Nothing Then
        errorMessages.Add "В файле " & fileName & " отсутствует запись: ""ФИО Прораба / Мастера""."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lCol3 As Long
    lCol3 = Rngt3.Column ' Столбец, где было найдено слово "ФИО Прораба / Мастера"

''В Prov.Range("D") столбце D начиная от строки 11 и до последней строки (LastRowDPQ) в столбце D находятся значения, являющиеся критерием для поиска.
'Нужно найти в отрывшемся файле В Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range("D") столбце D начиная от строки 11 и до последней строки (LastRowDPQ2), эти же значения критерии поиска из Prov.Range("D"), и если такое одно  значение нашлось, то
'1. Из соответствующей строки искомого критерия (находящейся в массиве Prov.Range("O11:O" & LastRowDPQ)) нужно из столбца O скопировать значение и вставить  в  Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").  в  столбец lCol (в массиве начиная от строки 11 и до последней строки (LastRowDPQ2)) в соответствующую строку, в которой был найден искомый критерий из Prov.Range("D").
'2.  Из соответствующей строки искомого критерия (находящейся в массиве Prov.Range("P11:P" & LastRowDPQ)) нужно из столбца P скопировать значение и вставить  в  Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").  в  столбец lCol2 (в массиве начиная от строки 11 и до последней строки (LastRowDPQ2)) в соответствующую строку, в которой был найден искомый критерий из Prov.Range("D").
'3.  Из соответствующей строки искомого критерия (находящейся в массиве Prov.Range("Q11:Q" & LastRowDPQ)) нужно из столбца Q скопировать значение и вставить  в  Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").  в  столбец lCol3 (в массиве начиная от строки 11 и до последней строки (LastRowDPQ2)) в соответствующую строку, в которой был найден искомый критерий из Prov.Range("D").
'
'А если такое одно  значение не нашлось, то выделить массив Prov.Range("A:D") в соответствующей строке, на которой не был найден искомый критерий из Prov.Range("D") и окрасить его  RGB(219, 179, 182). И в Конце макроса сообщить обо всех не найденных критериях в  одном окне "Не были обнаружены некоторые КЛЮЧИ в "Объёмы ООО ""Р-СТРОЙ""
'
'А если такое значение нашлось не одно, то выделить массив Prov.Range("A:D") в соответствующей строке, на которой было найдено более 1 искомого критерия из Prov.Range("D") и окрасить его  RRGB(129, 131, 143). И в Конце макроса сообщить обо всех найденных критериях более одного в  одном отдельном окне "Были обнаружены два или более КЛЮЧЕЙ в "Объёмы ООО ""Р-СТРОЙ""
' Определяем последнюю строку на листе "Проверка" в столбце D
    Dim Prov As Worksheet
    Set Prov = Vb.Worksheets("Проверка")
    Vb.Worksheets("Проверка").Activate
    Dim LastRowDPQ As Long
    LastRowDPQ = Prov.Cells(Prov.Rows.Count, "D").End(xlUp).Row
    
     

' Определяем последнюю строку на листе "Объёмы ООО ""Р-СТРОЙ""" в столбце D
Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Activate
Dim LastRowDPQ2 As Long
LastRowDPQ2 = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Rows.Count, "D").End(xlUp).Row


' Создаем коллекции для хранения сообщений об ошибках
Dim notFoundKeys As Collection
Dim duplicateKeys As Collection
Set notFoundKeys = New Collection
Set duplicateKeys = New Collection

' Переменная для хранения найденных совпадений
Dim foundCount As Long





' Цикл по строкам на листе "Проверка"
Dim i As Long
For i = 11 To LastRowDPQ

Application.ScreenUpdating = True
' Подсчитываем количество заполненных строк до текущей строки i
Dim filledRowsCount As Long
filledRowsCount = WorksheetFunction.CountA(Prov.Range("D11:D" & i))  ' Вычитаем 1, чтобы исключить текущую строку i
' Обновляем прогресс в ячейке D1
Vb.Worksheets("Главный").Cells(1, "D").Value = "Обработано " & filledRowsCount & " строк из " _
                                                  & WorksheetFunction.CountA(Prov.Range("D11:D" & LastRowDPQ))
Application.ScreenUpdating = False
                                                      
    ' Получаем значение из столбца D на листе "Проверка"
    Dim searchValue As Variant
    searchValue = Prov.Cells(i, "D").Value
    
    ' Сбрасываем счетчик найденных совпадений
    foundCount = 0
    
    ' Цикл по строкам на листе "Объёмы ООО ""Р-СТРОЙ"""
    Dim j As Long
    For j = 11 To LastRowDPQ2
        ' Получаем значение из столбца D на листе "Объёмы ООО ""Р-СТРОЙ"""
        Dim targetValue As Variant
        targetValue = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(j, "D").Value
        
        ' Если значения совпадают
        If searchValue = targetValue Then
            foundCount = foundCount + 1
            
            ' Если найдено первое совпадение, копируем данные из столбцов O, P и Q
            If foundCount = 1 Then
                ' Копируем значение из столбца O
                Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(j, lCol).Value = Prov.Cells(i, "O").Value
                ' Копируем значение из столбца Q
                Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(j, lCol2).Value = Prov.Cells(i, "Q").Value
                ' Копируем значение из столбца P
                Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(j, lCol3).Value = Prov.Cells(i, "P").Value
            End If
        End If
    Next j
    
     ' Если совпадений не найдено
    If foundCount = 0 Then
        ' Выделяем строку на листе "Проверка" и окрашиваем её в цвет RGB(219, 179, 182)
        Prov.Range("A" & i & ":D" & i).Interior.Color = RGB(219, 179, 182)
        Prov.Cells(i, "T").Value = Prov.Cells(i, "O").Value    ' Копируем значение из ячейки O и вставляем его в ячейку T
        notFoundKeys.Add "Ключ не найден: " & searchValue   ' Добавляем сообщение об ошибке в коллекцию
    ' Если найдено более одного совпадения
    ElseIf foundCount > 1 Then
        ' Выделяем строку на листе "Проверка" и окрашиваем её в цвет RGB(129, 131, 143)
        Prov.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 242, 204)
        ' Добавляем сообщение об ошибке в коллекцию
        duplicateKeys.Add "Найдено более одного ключа: " & searchValue
    End If
Next i
 Prov.Range("T8") = "Не подтянувшийся объём"
 Prov.Range("U8") = "Дельта"
 Prov.Range("U9") = "=O9-T9"
  Prov.Range("T9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;T11:T" & LastRowDPQ & ")"
 ' убрать красноту из предварительных итогов
    Dim sCell As Range
    For Each sCell In Prov.Range("S11:S" & LastRowDPQ)    ' Проверка значений в столбце S и окрашивание строк
        If Not IsEmpty(sCell.Value) And Trim(sCell.Value) <> "" Then
            Prov.Range("A" & sCell.Row & ":R" & sCell.Row).Interior.Color = RGB(218, 238, 243)   'синий
        End If
    Next sCell


 
 












' Выводим сообщения об ошибках
If notFoundKeys.Count > 0 Then
    Dim notFoundMessage As String
    notFoundMessage = "Не были обнаружены некоторые КЛЮЧИ в ""Объёмы ООО ""Р-СТРОЙ"": " & vbCrLf
    For i = 1 To notFoundKeys.Count
        notFoundMessage = notFoundMessage & notFoundKeys(i) & vbCrLf
    Next i
    MsgBox notFoundMessage, vbExclamation, "Ошибка"
End If

If duplicateKeys.Count > 0 Then
    Dim duplicateMessage As String
    duplicateMessage = "Были обнаружены два или более КЛЮЧЕЙ в ""Объёмы ООО ""Р-СТРОЙ"": " & vbCrLf
    For i = 1 To duplicateKeys.Count
        duplicateMessage = duplicateMessage & duplicateKeys(i) & vbCrLf
    Next i
    MsgBox duplicateMessage, vbExclamation, "Ошибка"
End If

' Восстанавливаем настройки Excel
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'' Закрываем файл Ob без сохранения изменений
'Ob.Close SaveChanges:=False

End Sub



' Функция для открытия диалогового окна выбора файла с указанием начальной папки
Function OpenFileDialog5(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Выберите файл ""Объёмы ООО Р-СТРОЙ Р3 АГПЗ ОЗХ _ ..."" за отчётный период"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' Указываем начальную папку, если она задана
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            OpenFileDialog5 = .SelectedItems(1) ' Возвращаем выбранный файл
        Else
            OpenFileDialog5 = False ' Если файл не выбран
        End If
    End With
End Function


