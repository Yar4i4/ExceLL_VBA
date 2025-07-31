Attribute VB_Name = "Module5"
  Sub Из_P3_в_Факт_ФО()
      Application.ScreenUpdating = False 'Больше не обновляем страницы после каждого действия
    Application.Calculation = xlCalculationManual 'Расчёты переводим в ручной режим
    Application.EnableEvents = False 'Отключаем события
    If Workbooks.Count Then
    ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False 'Не отображаем границы ячеек
    End If
    Application.DisplayStatusBar = False 'Отключаем статусную строку
        Application.DisplayAlerts = False 'Отключаем сообщения Excel
    
    
     ' Присваиваем текущую книгу переменной Vb
    Set Vb = ThisWorkbook
    
    ' Открываем диалоговое окно для выбора файла
    Dim filePath As Variant
    filePath = OpenFileDialog2(Vb.Path) ' Указываем папку, где находится текущая книга
    
    ' Проверяем, был ли выбран файл
    If filePath = False Then
        MsgBox "Файл не выбран!", vbExclamation
        Exit Sub
    End If
    
    ' Открываем выбранный файл и присваиваем его переменной P3
    Set P3 = Workbooks.Open(filePath)
    
    ' Список имен листов для проверки
    sheetNames = Array("Сводная по СМУ", "Сводная по Прорабам", "Факт ФО на текущий день")
    missingSheets = "" ' Переменная для хранения отсутствующих листов
    
    ' Проверяем существование каждого листа
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = P3.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            ' Если лист не найден, добавляем его имя в список отсутствующих
            missingSheets = missingSheets & sheetName & vbCrLf
        Else
            ' Если лист существует, очищаем ссылку на него
            Set ws = Nothing
        End If
    Next sheetName
    
    ' Выводим результат проверки
    If missingSheets <> "" Then
        MsgBox "Отсутствуют следующие листы:" & vbCrLf & missingSheets, vbExclamation
        P3.Close SaveChanges:=False ' Закрываем файл без сохранения изменений
        Exit Sub
    End If
    
    ' Создаем новую книгу
    Dim newBook As Workbook
    Set newBook = Workbooks.Add
       
    
    ' Копируем листы в новую книгу
    For Each sheetName In sheetNames
        P3.Sheets(sheetName).Copy After:=newBook.Sheets(newBook.Sheets.Count)
    Next sheetName
    
     ' Удаляем лист "Лист1" из новой книги, если он существует
    On Error Resume Next ' Игнорируем ошибку, если лист не существует
    Application.DisplayAlerts = False ' Отключаем предупреждения
    newBook.Sheets("Лист1").Delete ' Удаляем лист
    Application.DisplayAlerts = True ' Включаем предупреждения обратно
    On Error GoTo 0 ' Возвращаем стандартную обработку ошибок
    
    
    ' Работаем с листом "Факт ФО на текущий день" в новой книге
    Dim фактФО As Worksheet
    Set фактФО = newBook.Sheets("Факт ФО на текущий день")
    
    ' Определяем диапазон для обработки (столбцы P, Q, R, S, T, Z, AA, AB)
    Dim LastRow As Long
    LastRow = фактФО.Cells(фактФО.Rows.Count, "P").End(xlUp).Row ' Находим последнюю заполненную строку в столбце P
    
    ' Обрабатываем каждый столбец
    Dim col As Variant
    For Each col In Array("P", "Q", "R", "S", "T", "Z", "AA", "AB")
        ' Заменяем формулы на значения
        фактФО.Range(col & "4:" & col & LastRow).Value = фактФО.Range(col & "4:" & col & LastRow).Value
    Next col
    
  




   
   
   
   
    ' Работаем с листом "Сводная по СМУ" в новой книге
    Dim своднаяСМУ As Worksheet
    Set своднаяСМУ = newBook.Sheets("Сводная по СМУ")
    
    ' Определяем столбцы для обработки
    Dim columnsToProcess As Variant
    columnsToProcess = Array("C", "D", "F", "G", "J", "K", "M", "N") ' Указываем столбцы для обработки
    
    ' Обрабатываем каждый столбец
    Dim colSMU As Variant
    For Each colSMU In columnsToProcess
        ' Определяем диапазон для обработки (текущий столбец)
        Dim lastRow2 As Long
        lastRow2 = своднаяСМУ.Cells(своднаяСМУ.Rows.Count, colSMU).End(xlUp).Row ' Находим последнюю заполненную строку в столбце
        
        ' Обрабатываем ячейки в столбце
        Dim cell5 As Range
        For Each cell5 In своднаяСМУ.Range(colSMU & "4:" & colSMU & lastRow2)
            If InStr(1, cell5.Formula, "[") > 0 Then
                ' Если в формуле есть ссылка на внешний файл (квадратные скобки), удаляем её
                Dim formulaText As String
                formulaText = cell5.Formula
                
                ' Удаляем все ссылки на внешние файлы в формуле
                Do While InStr(1, formulaText, "[") > 0
                    Dim startPos As Long
                    Dim endPos As Long
                    startPos = InStr(1, formulaText, "[") ' Находим начало ссылки
                    endPos = InStr(startPos, formulaText, "]") ' Находим конец ссылки
                    
                    If startPos > 0 And endPos > 0 Then
                        ' Удаляем путь к файлу и оставляем только имя листа и ячейки
                        formulaText = Left(formulaText, startPos - 1) & Mid(formulaText, endPos + 1)
                    Else
                        Exit Do
                    End If
                Loop
                
                ' Обновляем формулу в ячейке
                cell5.Formula = formulaText
            End If
        Next cell5
    Next colSMU
   









' Работаем с листом "Сводная по Прорабам" в новой книге
    Dim своднаяПрорабы As Worksheet
    Set своднаяПрорабы = newBook.Sheets("Сводная по Прорабам")
    
    ' Определяем столбцы для обработки (от D до Z)
    Dim columnsToProcessPro As Variant
    columnsToProcessPro = Array("D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "U", "V", "W", "X", "Y", "Z")
    
    ' Обрабатываем каждый столбец
    Dim colPro As Variant
    For Each colPro In columnsToProcessPro
        ' Определяем диапазон для обработки (текущий столбец)
        Dim lastRowPro As Long
        lastRowPro = своднаяПрорабы.Cells(своднаяПрорабы.Rows.Count, colPro).End(xlUp).Row ' Находим последнюю заполненную строку в столбце
        
        ' Обрабатываем ячейки в столбце
        Dim cellPro As Range
        For Each cellPro In своднаяПрорабы.Range(colPro & "4:" & colPro & lastRowPro)
            If InStr(1, cellPro.Formula, "[") > 0 Then
                ' Если в формуле есть ссылка на внешний файл (квадратные скобки), удаляем её
                Dim formulaTextPro As String
                formulaTextPro = cellPro.Formula
                
                ' Удаляем все ссылки на внешние файлы в формуле
                Do While InStr(1, formulaTextPro, "[") > 0
                    Dim startPosPro As Long
                    Dim endPosPro As Long
                    startPosPro = InStr(1, formulaTextPro, "[") ' Находим начало ссылки
                    endPosPro = InStr(startPosPro, formulaTextPro, "]") ' Находим конец ссылки
                    
                    If startPosPro > 0 And endPosPro > 0 Then
                        ' Удаляем путь к файлу и оставляем только имя листа и ячейки
                        formulaTextPro = Left(formulaTextPro, startPosPro - 1) & Mid(formulaTextPro, endPosPro + 1)
                    Else
                        Exit Do
                    End If
                Loop
                
                ' Обновляем формулу в ячейке
                cellPro.Formula = formulaTextPro
            End If
        Next cellPro
    Next colPro



' Формируем полный путь для сохранения файла
Dim folderPath As String
folderPath = Vb.Path ' Используем путь текущей книги (ThisWorkbook)

' Проверяем, существует ли папка
If Dir(folderPath, vbDirectory) = "" Then
    MsgBox "Папка для сохранения не существует: " & folderPath, vbExclamation
    Exit Sub
End If

' Формируем имя файла
Dim fileName As String
With P3.Sheets("Сводная по Прорабам")
    ' Извлекаем части даты из ячейки A1
    Dim YearPart As String
    Dim monthPart As String
    Dim dayPart As String

    ' Год: 4 символа справа
    YearPart = Right(.Range("A1").Value, 4) ' Последние 4 символа

    ' Месяц: 7й и 6й символы справа
    monthPart = Mid(.Range("A1").Value, Len(.Range("A1").Value) - 6, 2)

    ' День: 10й и 9й символы справа
    dayPart = Mid(.Range("A1").Value, Len(.Range("A1").Value) - 9, 2)

    ' Формируем имя файла
    fileName = "Факт ФО по дням _ " & YearPart & "." & monthPart & ".01-" & dayPart & ".xlsb"
End With

' Формируем полный путь
Dim fullPath As String
fullPath = folderPath & "\" & fileName


' Разрываем все внешние связи в новой книге
Dim link As Variant
For Each link In newBook.LinkSources(xlExcelLinks)
    newBook.BreakLink Name:=link, Type:=xlLinkTypeExcelLinks
Next link


' Сохраняем новую книгу
Application.DisplayAlerts = False ' Отключаем предупреждения
newBook.SaveAs fileName:=fullPath, FileFormat:=xlExcel12
Application.DisplayAlerts = True ' Включаем предупреждения обратно



    Application.ScreenUpdating = True 'Включаем обновление экрана после каждого события
    Application.Calculation = xlCalculationAutomatic 'Расчёты формул - снова в автоматическом режиме
    Application.EnableEvents = True  'Включаем события
    If Workbooks.Count Then
    ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True 'Показываем границы ячеек
    End If
    Application.DisplayStatusBar = True 'Возвращаем статусную строку
    Application.DisplayAlerts = True 'Разрешаем
  End Sub
' Функция для открытия диалогового окна выбора файла с указанием начальной папки
Function OpenFileDialog2(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) ' Используем FileDialog
    
    With fd
        .Title = "Выберите файл ""Р3 АГПЗ ОЗХ _ Факт..."" за отчётный период"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' Указываем начальную папку, если она задана
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            OpenFileDialog2 = .SelectedItems(1)
        Else
            OpenFileDialog2 = False
        End If
    End With
End Function

