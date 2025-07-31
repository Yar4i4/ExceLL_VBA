Attribute VB_Name = "Module3"
' ФИО из ячейки в созданную ниже строку раскидать, скопировать
Type Item
    fio As String
    number As String
End Type
Function Is_Fio(sl As String) As Item()
    Dim itm() As Item ' Объявляем массив структур Item для хранения результатов
    ReDim itm(0) ' Инициализируем массив с одним элементом
    sl = Replace(sl, "-", "") ' Удаляем все дефисы из входной строки
    Static RegExp As Object ' Объявляем статическую переменную для объекта RegExp (регулярное выражение)
    If RegExp Is Nothing Then ' Если объект RegExp еще не создан
        Set RegExp = CreateObject("VBScript.RegExp") ' Создаем объект RegExp
        RegExp.IgnoreCase = True ' Устанавливаем флаг игнорирования регистра
        RegExp.Global = True ' Устанавливаем флаг глобального поиска
    End If
    RegExp.Pattern = "(([а-яё\s]+){3,})(\(?([0-9,]+)?)" ' Задаем шаблон регулярного выражения для поиска ФИО и чисел
    paralast = -1 ' Инициализируем переменную для отслеживания последнего индекса в массиве
    Set oMatches = RegExp.Execute(sl) ' Выполняем поиск по регулярному выражению
    If oMatches.Count > 0 Then ' Если найдены совпадения
        Is_Numeric = oMatches(0).SubMatches(0) ' Записываем первое совпадение в переменную (не используется далее)
    End If
    For n = 0 To oMatches.Count - 1 ' Цикл по всем найденным совпадениям
        If Len(Trim(oMatches(n).SubMatches(0))) > 4 Then ' Если длина найденного ФИО больше 4 символов
            ' Проверяем, содержит ли ФИО определенное ключевое слово (игнорируем, если содержит)
            If InStr(1, " " & Trim(oMatches(n).SubMatches(0)) & " ", "корректировкараннебыла,но не стал менять код", vbTextCompare) > 0 Then
                ReDim itm(0) ' Если найдено ключевое слово, сбрасываем массив
                Is_Fio = itm ' Возвращаем пустой массив
                Exit Function ' Завершаем функцию
            End If
            paralast = UBound(itm) + 1 ' Увеличиваем индекс массива
            ReDim Preserve itm(paralast) ' Расширяем массив с сохранением данных
            itm(paralast).fio = Trim(oMatches(n).SubMatches(0)) ' Записываем ФИО в массив
            itm(paralast).number = oMatches(n).SubMatches(3) ' Записываем число (если есть) в массив
        End If
    Next
    Is_Fio = itm ' Возвращаем заполненный массив
End Function
Sub Извлечь_данные_из_Объёмы_ООО()
  Application.ScreenUpdating = False ' Больше не обновляем страницы после каждого действия
    Application.Calculation = xlCalculationManual ' Расчёты переводим в ручной режим
    Application.EnableEvents = False ' Отключаем события
    If Workbooks.Count Then
        ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False ' Не отображаем границы ячеек
    End If
    Application.DisplayStatusBar = False ' Отключаем статусную строку
    Application.DisplayAlerts = False ' Отключаем сообщения Excel
    ' Присваиваем текущую книгу переменной Vb
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    Vb.Worksheets("Проверка").Cells.Clear
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
    End If
    ' Открываем диалоговое окно для выбора файла
    Dim filePath As Variant
    filePath = OpenFileDialog1(Vb.Path) ' Указываем папку, где находится текущая книга
    ' Проверяем, был ли выбран файл
    If filePath = False Then
        MsgBox "Файл не выбран!", vbExclamation
        Exit Sub
    End If
' Открываем выбранный файл
Dim Ob As Workbook
Set Ob = Workbooks.Open(filePath)

' Проверяем, существует ли лист "Объёмы ООО ""Р-СТРОЙ"""
Dim targetSheet As Worksheet
On Error Resume Next
Set targetSheet = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""")
On Error GoTo 0

If targetSheet Is Nothing Then
    MsgBox "Лист ""Объёмы ООО ""Р-СТРОЙ"""" не найден!", vbExclamation
    Exit Sub
End If
' Раскрываем автофильтр и скрытые строки/столбцы на всех листах книги Ob
Dim ws As Worksheet
For Each ws In Ob.Worksheets
    ' Раскрываем автофильтр, если он активен
    If ws.AutoFilterMode And ws.FilterMode Then
        ws.ShowAllData ' Если есть фильтр и не все данные видны, показываем всё
    End If
    ' Проверяем и раскрываем скрытые строки
    If ws.Cells.EntireRow.Hidden Then
        ws.Cells.EntireRow.Hidden = False ' Раскрываем все строки
    End If
    ' Проверяем и раскрываем скрытые столбцы
    If ws.Cells.EntireColumn.Hidden Then
        ws.Cells.EntireColumn.Hidden = False ' Раскрываем все столбцы
    End If
Next ws
' Раскрываем скрытые строки и столбцы во всех листах текущей книги (Ob)
Dim currentSheet As Worksheet
For Each currentSheet In Ob.Sheets
    If currentSheet.Cells.EntireRow.Hidden Then
        currentSheet.Cells.EntireRow.Hidden = False ' Раскрываем все строки
    End If
    If currentSheet.Cells.EntireColumn.Hidden Then
        currentSheet.Cells.EntireColumn.Hidden = False ' Раскрываем все столбцы
    End If
Next currentSheet
    ' Активируем лист "Объёмы ООО ""Р-СТРОЙ"""
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Activate
    ' Получаем имя файла (без пути и расширения)
    fileName = Left(Ob.Name, InStrRev(Ob.Name, ".") - 1)
    ' Извлекаем последние два символа (числа) из имени файла для переменной Den
    Den = Right(fileName, 2)
    ' Извлекаем 5-е и 4-е числа справа из имени файла для переменной Mesyac
    If Len(fileName) >= 5 Then
        Mesyac = Mid(fileName, Len(fileName) - 4, 2)
    Else
        Mesyac = "" ' Если длина имени файла меньше 5 символов
    End If
    ' Поиск номера ячейки, содержащей "ФО за МесяцТочкаДень"
    Set Rngt = Cells.Find("ФО за " & Den & "." & Mesyac, , xlFormulas, xlWhole)
    If Rngt Is Nothing Then
        MsgBox "Название файла должно заканчиваться на МесяцТочкаДень, типа:" & Chr(10) & " 12.30" & Chr(10) & _
               " Либо на листе Объёмы ООО ""Р-СТРОЙ"" в строке 8" & vbCrLf & " отсутствует запись: ФО за " & Den & "." & Mesyac
        Exit Sub
    End If
    lRow = Rngt.Row ' Строка, где было найдено слово
    lCol = Rngt.Column ' Столбец, где было найдено слово
    LR = Cells(Rows.Count, lCol).End(xlUp).Row
    sAddress = Rngt.Address ' Адрес ячейки, где было найдено слово
     ''''проверка объема за день до фильтра''''
     Application.Calculation = xlCalculationAutomatic 'Расчёты переводим в авто режим
      ' Проверяем, есть ли значение в ячейке
   volumeValue = Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Cells(9, lCol).Value
    If IsEmpty(volumeValue) Or volumeValue = 0 Then
        ' Если значение отсутствует или равно 0, выводим сообщение
        MsgBox "За отчетный день отсутствуют объёмы. Проверьте наименование файла ""Объёмы..."" (последние 5 символов - это ""месяц точка день"").", vbExclamation, "Ошибка"
        Exit Sub
    Else
        ' Если значение есть, записываем его в ячейку на листе "Проверка"
        Vb.Worksheets("Проверка").Cells(9, 15).Value = volumeValue
    End If
      Vb.Worksheets("Проверка").Cells(8, 15) = "Проверка объёма до фильтра"
      Application.Calculation = xlCalculationManual 'Расчёты переводим в ручной режим
    ' Автофильтр без пустых значений и без нулей в столбце с объёмом за день
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Cells(lRow + 2, lCol), "FF" & LR).AutoFilter Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
       '  '  '  Range(Columns(15), Columns(lCol - 1)).Hidden = True ' Скрыть столбцы
    ' Поиск номера ячейки, содержащей "Подразделение"
    Set Rngt2 = Cells.Find("Подразделение", , xlFormulas, xlWhole)
    If Rngt2 Is Nothing Then
        MsgBox "На листе Объёмы ООО ""Р-СТРОЙ"" в строке 8" & vbCrLf & "отсутствует запись: Подразделение (без всяких пробелов и прочих добавлений)"
        Exit Sub
    End If
    
    lCol2 = Rngt2.Column ' Столбец, где было найдено слово
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Columns(lCol + 1), Columns(lCol2 - 1)).Hidden = True ' Скрыть столбцы
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Columns(15), Columns(lCol - 1)).Hidden = True ' Скрыть столбцы
   Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Columns("A:D").EntireColumn.Hidden = True ' Скрыть столбцы
    ' Копируем массив в новую книгу
    Ob.Worksheets("Объёмы ООО ""Р-СТРОЙ""").Range(Cells(8, 5), Cells(LR, lCol2 + 1)).Copy
    ' Вставляем скопированные данные в книгу Vb на лист "Проверка" как значения
        
    Vb.Worksheets("Проверка").Cells(8, 1).PasteSpecial Paste:=xlPasteValues
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
        Dim Prov As Worksheet
        Set Prov = Vb.Worksheets("Проверка")
    Vb.Worksheets("Проверка").Activate
    Vb.Worksheets("Проверка").Columns("L:M").Cut
    Vb.Worksheets("Проверка").Columns("A:A").Insert Shift:=xlToRight
    Vb.Worksheets("Проверка").Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    LRN = Cells(Rows.Count, 14).End(xlUp).Row
    Vb.Worksheets("Проверка").Range("A11:A" & LRN) = Den
    Vb.Worksheets("Проверка").Columns("C:C").ColumnWidth = 40
    Vb.Worksheets("Проверка").Columns("C:C").ReadingOrder = xlContext
    Vb.Worksheets("Проверка").Columns("C:C").WrapText = True
     Vb.Worksheets("Проверка").Rows("1:10").RowHeight = 11
    Vb.Worksheets("Проверка").Rows("8:8").RowHeight = 25
    Vb.Worksheets("Проверка").Rows("8:10").HorizontalAlignment = xlCenter
    Vb.Worksheets("Проверка").Rows("8:10").VerticalAlignment = xlCenter
    Vb.Worksheets("Проверка").Rows("8:8").WrapText = True
    Vb.Worksheets("Проверка").Rows("9:40").EntireRow.AutoFit
    Vb.Worksheets("Проверка").Columns("D:M").ColumnWidth = 1
    Vb.Worksheets("Проверка").Columns("N:R").ColumnWidth = 17
    Vb.Worksheets("Проверка").Columns("A:A").ColumnWidth = 3
    Vb.Worksheets("Проверка").Columns("B:B").ColumnWidth = 11
    



Vb.Worksheets("Проверка").Range("Q1:Q9") = Ob.Worksheets("Свод по ИД (Р)").Range("F18:F25").Value  '  Измени, если изменится количество СМУ
Vb.Worksheets("Проверка").Range("R1:R9") = Ob.Worksheets("Свод по ИД (Р)").Range("K18:K25").Value '  Измени, если изменится количество СМУ

Vb.Worksheets("Проверка").Range("R10") = Ob.Worksheets("Свод по ИД (Р)").Range("K3").Value ' вставляем  массив
Vb.Worksheets("Проверка").Range("Q10") = "Выполнено всего" '
        ' выделить латиницу красным от Alex_ST
        Range("C11:C" & LRN).Select
        If TypeName(Selection) <> "Range" Then Exit Sub
        If Intersect(Selection, ActiveSheet.UsedRange) Is Nothing Then Exit Sub
        Dim rCell As Range, i4%, ASCII%, iColor%
        Application.ScreenUpdating = False
        For Each rCell In Intersect(Selection, ActiveSheet.UsedRange)
        For i4 = 1 To Len(rCell)
        ASCII = Asc(Mid(rCell, i4, 1))
        If (ASCII >= 192 And ASCII <= 255) Or ASCII = 168 Or ASCII = 184 Then iColor = 1   ''цвет символов РУС черный
        If (ASCII >= 65 And ASCII <= 90) Or (ASCII >= 97 And ASCII <= 122) Then iColor = 3   ''цвет символов LAT красный
        rCell.Characters(Start:=i4, Length:=1).Font.ColorIndex = iColor
        Next i4
        Next rCell
        Application.ScreenUpdating = True
        Intersect(Selection, ActiveSheet.UsedRange).Select
    ' Закрываем книгу Ob
    Ob.Close SaveChanges:=True
    Range("A8:N8", "A3:N3").Interior.Color = RGB(183, 222, 232)
    Range("A9:N9").Interior.Color = RGB(218, 238, 243) '.Interior.Color = RGB(200, 138, 143)
    Range("A8") = "День"
    Range("O9").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;N4:N" & LRN & ")"
'
    Vb.Worksheets("Проверка").Range("R1:R10").NumberFormat = "#,##0.00"
    Vb.Worksheets("Проверка").Range("N9:O9").NumberFormat = "#,##0.00"
    ' проверка на превышение при сравнении с итогами в шапке и по дням
    If Range("N9") <> Range("O9") Then
    Range("N9").Interior.Color = RGB(200, 138, 143)
    End If
    Range("A11:N" & LRN).Select

        ' Разбить ячейки столбца C, если в ячейке больше одной ФИО.
        Dim itm() As Item, A()
        Vb.Worksheets("Проверка").Activate
        Set Prov = Vb.Worksheets("Проверка")
         LastRow = Prov.Cells(Prov.Rows.Count, "C").End(xlUp).Row
        dx = Prov.Range("C1:O" & LastRow)  'dx = Prov.Range("C1:N" & LastRow)
        Delta = 0
        For n = UBound(dx) To 11 Step -1
            s5$ = dx(n, 1)
            If s5 = "" Then Exit For
            itm = Is_Fio(s5)
            If UBound(itm) > 1 Then
                LastRow1 = n + 1
                lastRow2 = UBound(itm) + n
                ReDim A(1 To UBound(itm), 1 To 12)  ' столбец 12 для чесел
                For i = 1 To UBound(itm)
                    A(i, 1) = itm(i).fio
                    If Val(Replace(itm(i).number, ",", ".")) <> 0 Then
                     A(i, 12) = Val(Replace(itm(i).number, ",", "."))  ' столбец 12 для чесел
                    End If
                Next
                 Prov.Rows(LastRow1 & ":" & lastRow2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove  'вставить строки ниже объед ячейки с ФИО
                 Prov.Range("C" & LastRow1).Resize(UBound(itm), 12) = A   ' столбец 12 в , 12) = A  столбец 12 для чесел
                 Prov.Range("N" & LastRow1).Resize(UBound(itm), 1).Interior.Color = RGB(255, 250, 235) 'заливка ячейки-донора бледно светло-желтый
                 Prov.Cells(n, 3).Interior.Color = RGB(255, 246, 221) 'заливка ячейки-донора бледно желтый
                 Prov.Cells(n, 14).Interior.Color = RGB(255, 246, 221) 'заливка ячейки-донора бледно желтый
                 Prov.Cells(n, 16) = "Данная строка удалится"
                  Prov.Cells(n, 14).Copy
                  Prov.Cells(n, 15).PasteSpecial Paste:=xlPasteValues
                  Prov.Cells(n, 15).PasteSpecial Paste:=xlPasteFormats
                  
                   'Prov.Cells(n, 15).Interior.Color = RGB(255, 249, 231) 'заливка ячейки-потребителя  бледно желтый -1
            End If
        Next
    Prov.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    LastRow3 = Prov.Cells(Prov.Rows.Count, "C").End(xlUp).Row
    Prov.Range("A11:A" & LastRow3).Formula = "=ROW()-10"

    ' сравни разнесенные объёмы по ячейкам столбца О ниже (до удаления строки-донора) и подкрасить фон .Interior.Color = 13431551
                Application.ScreenUpdating = False
                Ae = Range("O11").End(xlDown).Row
                b = Application.CountA(Range("N11:N" & Ae))
                d = 11
                For C = 1 To b
                e = Application.match("*", Range("N" & d & ":N" & Ae), 0)
                d = d + e
                f = Application.match("*", Range("N" & d & ":N" & Ae), 0)
                g = d - 1
                If C <> b Then
                h = d + f - 2
                Else
                h = Ae
                End If
                If g <> h Then
                i = Application.Round(Range("O" & g).Value, 4) ' округли до 4 знаков после запятой
                j = Application.Round(Application.Sum(Range("O" & g + 1 & ":O" & h)), 4) ' округли до 4 знаков после запятой
                If i <> j Then Range("O" & g + 1 & ":O" & h).Interior.Color = RGB(219, 179, 182) ' светло красный
                End If
                Next
                
' Проверяем, есть ли строки ниже nx и удаляем пустые ниже последней заполненной
Dim nx As Long
' Worksheets("Проверка").UsedRange  'сбросить результат с последней ячейкой, строкой
If LastRow3 < Prov.Rows.Count Then
    Prov.Rows(LastRow3 + 1 & ":" & Prov.Rows.Count).Delete ' Удаляем все строки ниже последней заполненной
End If
                
                
    Dim LastRow5 As Long  ' Для хранения номера строки последней непустой ячейки
    ActiveSheet.UsedRange  'сбросить результат с последней ячейкой, строкой
    LastRow5 = Cells.SpecialCells(xlLastCell).Row  'определение последней заполненной строки вне зависимости от столбца
' проверить если в столбце 2 между ячейкой B4  lastRow = Prov.Cells(Prov.Rows.Count, "C").End(xlUp).Row нижней заполненной есть пустые ячейки,
'' то в каждую пустую ячейку столбца 2 вставить верхнюю не пустую ячейку столбца 2
'' Проходим по строкам от 11 до последней заполненной строки
   If LastRow5 >= 11 Then
    ' Проходим по строкам от LastRow до 11
    Dim iw As Long
    For iw = LastRow5 To 11 Step -1
        ' Проверяем, пустая ли ячейка в столбце B
        If IsEmpty(Prov.Cells(iw, "B").Value) Or Prov.Cells(iw, "B").Value = "" Then
            ' Если ячейка пустая, ищем первую непустую ячейку выше
            Dim ji As Long
            For ji = iw - 1 To 11 Step -1
                If Not IsEmpty(Prov.Cells(ji, "B").Value) And Prov.Cells(ji, "B").Value <> "" Then
                    ' Копируем значения из первой непустой ячейки выше
                    Prov.Range("B" & iw & ":C" & iw).Value = Prov.Range("B" & ji & ":C" & ji).Value
                    Prov.Range("E" & iw & ":N" & iw).Value = Prov.Range("E" & ji & ":N" & ji).Value
                    Exit For
                End If
            Next ji
        Else
            ' Если ячейка не пустая, запоминаем ее номер строки
            LastRow5 = iw
        End If
    Next iw
End If
    
   ' 2 проверка...
     LastRow7 = Cells.SpecialCells(xlLastCell).Row  'определение последней заполненной строки вне зависимости от столбца
 ' сравни суммы итого,
                     ''    '  [P2].FormulaLocal = "=СУММ(O4:O" & LastRow7 & ")-СУММ(P4:P" & LastRow7 & ")"
    Prov.[P2].FormulaLocal = "=ОКРУГЛ(СУММ(O11:O" & LastRow7 & ")-СУММ(P11:P" & LastRow7 & ");4)"
    Range("O9").Value = Round(Range("O9").Value, 4)
    Range("P9").Value = Round(Range("P9").Value, 4)
RoundedValue = Round(Worksheets("Проверка").Range("O9").Value, 5)
Worksheets("Проверка").Range("O9").Value = RoundedValue
RoundedValue = Round(Worksheets("Проверка").Range("P9").Value, 5)
Worksheets("Проверка").Range("P9").Value = RoundedValue
    If Prov.Range("O9").Value <> Range("P9").Value Then
    Prov.Range("O9").Interior.Color = RGB(219, 179, 182)
    End If
 ' формулу снова, т.к. при сверке стало как значение
  Prov.[P9].FormulaLocal = "=ОКРУГЛ(СУММ(O11:O" & LastRow7 & ")-СУММ(P11:P" & LastRow7 & ");4)"

     ' удали первые и последние пробелы
    For Row = 11 To LastRow7
    Do While Right(Cells(Row, "D").Value, 1) = " " ' что удаляем в конце ячейки столбца 2
    Cells(Row, "D").Value = Left(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Left(Cells(Row, "D").Value, 1) = " " ' что удаляем в конце ячейки столбца 2
    Cells(Row, "D").Value = Right(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Right(Cells(Row, "D").Value, 1) = Chr(10) ' символ возврата каретки+смещения на одну строку
    Cells(Row, "D").Value = Left(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Left(Cells(Row, "D").Value, 1) = Chr(10) ' символ возврата каретки+смещения на одну строку
    Cells(Row, "D").Value = Right(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Next
    ' ФИО в порядок ЧанВанЮ Тимур Владимирович  корректировка
    Range("D11:D" & LastRow7).Replace "ЧанВанЮ Тимур Владимирович", "Чан-Ван-Ю Тимур Владимирович", xlPart
    Range("D11:D" & LastRow7).Replace "корректировка", "Корректировка", xlPart
    Prov.Range("A10:P" & LastRow7).AutoFilter 'Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
    
    
    Prov.Activate
  
'    ' Установить фиксацию строк с 1 по 10
'    Prov.Rows("11:16").Select
'    ActiveWindow.FreezePanes = True
 
Dim ik As Long
Dim cell7 As Range
Dim Value As String
Dim foundInvalidChar As Boolean
Dim invalidChars As String
Dim char As Variant
' Список нежелательных символов
invalidChars = "!@#$%^&*(){}[]<>?|/~+=`"
' Проходим по всем ячейкам в столбце O, начиная с O4
For ik = 11 To LastRow7
Set cell7 = Prov.Cells(ik, "O")
Value = cell7.Value
' Проверяем, является ли значение пустым
If Not IsEmpty(Value) And Value <> "" Then
foundInvalidChar = False
' Проверка на наличие точки вместо запятой
If InStr(Value, ".") > 0 Then
cell7.Interior.Color = RGB(219, 179, 182) ' Подсветка ячейки светло-красным
MsgBox "Обнаружена точка вместо запятой в ячейке " & cell7.Address
End If
' Проверка на лишние пробелы
If Trim(Value) <> Value Then
cell7.Interior.Color = RGB(219, 179, 182) ' Подсветка ячейки желтым
MsgBox "Обнаружены лишние пробелы в ячейке " & cell7.Address
End If
' Проверка на наличие нежелательных символов
For irt = 1 To Len(invalidChars)
char = Mid(invalidChars, irt, 1) ' Получаем один символ из списка
If InStr(Value, char) > 0 Then
cell7.Interior.Color = RGB(219, 179, 182) ' Подсветка ячейки светло-красным
MsgBox "Обнаружен недопустимый символ '" & char & "' в ячейке " & cell7.Address
foundInvalidChar = True
Exit For
End If
Next irt
' Если найдены нежелательные символы, пропускаем дальнейшие проверки для этой ячейки
If foundInvalidChar Then GoTo NextCell
End If
NextCell:
Next ik
Application.GoTo Reference:=Prov.Range("A1"), Scroll:=True
MsgBox "Данные из файла " & fileName & " скопированы на лист ""Проверка""." & Chr(10) _
& "Проверьте правильно ли разнесены ФИО и объёмы." & Chr(10) & "В случае расхождения исправьте данные"
' выкл ускорение макроса
Application.ScreenUpdating = True 'Включаем обновление экрана после каждого события
Application.Calculation = xlCalculationAutomatic 'Расчёты формул - снова в автоматическом режиме
Application.EnableEvents = True 'Включаем события
If Workbooks.Count Then
ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True 'Показываем границы ячеек
End If
Application.ScreenUpdating = True 'Включаем обновление экрана после каждого события
Application.Calculation = xlCalculationAutomatic 'Расчёты формул - снова в автоматическом режиме
Application.EnableEvents = True 'Включаем события
If Workbooks.Count Then
ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True 'Показываем границы ячеек
End If
Application.DisplayStatusBar = True 'Возвращаем статусную строку
Application.DisplayAlerts = True 'Разрешаем
End Sub

' Функция для открытия диалогового окна выбора файла с указанием начальной папки
Function OpenFileDialog1(Optional InitialFolder As String = "") As Variant
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
            OpenFileDialog1 = .SelectedItems(1) ' Возвращаем выбранный файл
        Else
            OpenFileDialog1 = False ' Если файл не выбран
        End If
    End With
End Function


