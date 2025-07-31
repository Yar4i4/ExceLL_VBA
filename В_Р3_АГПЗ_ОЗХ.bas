Attribute VB_Name = "Module4"
Sub В_Р3_АГПЗ_ОЗХ()
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
    Vb.Worksheets("Проверка").Activate
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
    End If

    ' Копируем данные из столбца O
    Vb.Sheets("Проверка").Range("O:O").Copy
    ' Вставляем данные как значения
    Vb.Sheets("Проверка").Range("O:O").PasteSpecial Paste:=xlPasteValues

    ' Проверяем в столбце Q наличие текста "Данная строка удалится" и удаляем строки
    Set Prov = Vb.Sheets("Проверка")
    ActiveSheet.UsedRange  'сбросить результат с последней ячейкой, строкой
    LastRow = Prov.Cells.SpecialCells(xlLastCell).Row 'определение последней заполненной строки вне зависимости от столбца
    For i = LastRow To 1 Step -1
        If Prov.Cells(i, "Q").Value = "Данная строка удалится" Then
            Prov.Rows(i).Delete
        End If
    Next i

    ' Подсчитываем количество строк от 11 до LastRow
    Dim KolProv As Long
    ActiveSheet.UsedRange  'сбросить результат с последней ячейкой, строкой
    LastRowPD = Prov.Cells.SpecialCells(xlLastCell).Row 'определение последней заполненной строки вне зависимости от столбца
    If KolProv = 0 Then
        KolProv = LastRowPD - 10 ' Вычитаем 10, так как начинаем с 11 строки
    End If

    ' Открываем диалоговое окно для выбора файла
    Dim filePath As Variant
    filePath = OpenFileDialog(Vb.Path) ' Указываем папку, где находится текущая книга

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
    If missingSheets = "" Then
        ' MsgBox "Все необходимые листы существуют.", vbInformation
    Else
        MsgBox "Отсутствуют следующие листы:" & vbCrLf & missingSheets, vbExclamation
        P3.Close SaveChanges:=False ' Закрываем файл без сохранения изменений, если листы отсутствуют
        Exit Sub
    End If
    
    ' Проверяем, есть ли скрытые строки и столбцы на листах "Сводная по СМУ", "Сводная по Прорабам", "Факт ФО на текущий день"
    sheetNames = Array("Сводная по СМУ", "Сводная по Прорабам", "Факт ФО на текущий день")
    
    ' Проходим по каждому листу
    For i2 = LBound(sheetNames) To UBound(sheetNames)
        Set ws = P3.Sheets(sheetNames(i2))
        
        ' Проверяем наличие скрытых строк
        hasHiddenRows = False
        On Error Resume Next
        hasHiddenRows = ws.Cells.SpecialCells(xlCellTypeVisible).Rows.Count < ws.Rows.Count
        On Error GoTo 0
        
        ' Проверяем наличие скрытых столбцов
        hasHiddenColumns = False
        On Error Resume Next
        hasHiddenColumns = ws.Cells.SpecialCells(xlCellTypeVisible).Columns.Count < ws.Columns.Count
        On Error GoTo 0
        
        ' Если есть скрытые строки, раскрываем их
        If hasHiddenRows Then
            ws.Rows.Hidden = False
            ' MsgBox "На листе '" & ws.Name & "' были скрытые строки. Они раскрыты."
        End If
        
        ' Если есть скрытые столбцы, раскрываем их
        If hasHiddenColumns Then
            ws.Columns.Hidden = False
            ' MsgBox "На листе '" & ws.Name & "' были скрытые столбцы. Они раскрыты."
        End If
    Next i2
     
     
     
    ' На листе "Факт ФО на текущий день" в столбце B найти значение из книги Vb лист "Проверка" ячейки B4, где s22 это значение из ячейки B4
    Set wsFact = P3.Sheets("Факт ФО на текущий день")
    s22 = Prov.Range("B11").Value
    ' Ищем значение только в столбце B
'    Set FoundCell = wsFact.Columns("B").Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
    ' Ищем значение только в столбце B, начиная с ячейки B4 и ниже
Set foundCell = wsFact.Range("B4:B" & wsFact.Cells(wsFact.Rows.Count, "B").End(xlUp).Row).Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
' Если значение найдено, выводим сообщение и спрашиваем пользователя
If Not foundCell Is Nothing Then
    response = MsgBox("Отчетный день " & s22 & " в Р3 АГПЗ ОЗХ присутствует." & Chr(10) & "Удалите данные по отчётному дню из ""Р3 АГПЗ ОЗХ""" & Chr(10) & "Продолжить выполнение макроса?", vbInformation + vbYesNo, "Вопрос")
        ' Если пользователь нажал "Завершить" (кнопка "Нет"), завершаем выполнение макроса
    If response = vbNo Then
        Exit Sub
    End If
End If
' Если значение не найдено, копируем массив из книги Vb и вставляем в книгу P3
  ActiveSheet.UsedRange  'сбросить результат с последней ячейкой, строкой
    lastRow2 = Prov.Cells.SpecialCells(xlLastCell).Row 'определение последней заполненной строки вне зависимости от столбца
Set copyRange = Prov.Range("A11:O" & lastRow2)
' Находим строку для вставки в книге P3
insertRow = wsFact.Cells(wsFact.Rows.Count, "D").End(xlUp).Row + 1

                             ' формат новой строки
                                                                    wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).NumberFormat = "General"
                                                                With wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).Font
                                                                    .Name = "Calibri"
                                                                    .Size = 11
                                                                    .Bold = False
                                                                    .Italic = False
                                                                    .Color = 0
                                                                End With
                                                                wsFact.Range("A" & insertRow & ":BD" & insertRow).Interior.Color = RGB(255, 255, 255) ' Белый цвет
                                                                wsFact.Range("A" & insertRow & ":BD" & insertRow).Interior.Pattern = xlNone ' Отключаем заливку
                                                                With wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).Borders
                                                                    .LineStyle = 1
                                                                    .Weight = 1
                                                                    .Color = 6773025
                                                                End With
                                                                With wsFact.Range("A" & insertRow & ":BD" & insertRow)
                                                                .HorizontalAlignment = xlCenter ' Горизонтальное выравнивание по центру
                                                                .VerticalAlignment = xlCenter   ' Вертикальное выравнивание по центру
                                                                .WrapText = False                ' Перенос текста выкл
                                                                .Orientation = 0                ' Ориентация текста (0 градусов)
                                                                .IndentLevel = 0                ' Уровень отступа
                                                                End With
                                        wsFact.Range("D" & insertRow).HorizontalAlignment = xlLeft
                                        wsFact.Range("H" & insertRow & ":M" & insertRow).HorizontalAlignment = xlLeft
                          wsFact.Range("P" & insertRow & ":BB" & insertRow).NumberFormat = "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
                        wsFact.Range("AO" & insertRow & ":AR" & insertRow).NumberFormat = "_-* #,##0.0000000000000 _?_-;-* #,##0.0000000000000 _?_-;_-* ""-""?? _?_-;_-@_-"
                        wsFact.Range("T" & insertRow).NumberFormat = "_-* #,##0.000000000000 _?_-;-* #,##0.000000000000 _?_-;_-* ""-""?? _?_-;_-@_-"
    ' формулы
 wsFact.Range("P" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],Прайс!C[-14]:C[-6],6,0),0)" 'Един. расценка на оплату труда за ед. руб.
wsFact.Range("Q" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],Прайс!C[-15]:C[-7],7,0),0)" 'Един. расценка на строительное оборудование (эксплуатация машин и механизмов - ЭММ), руб.
wsFact.Range("R" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],Прайс!C[-16]:C[-8],8,0),0)" 'Един. расценка на мат-лы за ед., руб.
wsFact.Range("S" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],Прайс!C[-17]:C[-9],9,0),0)" 'Един. расценка на прямые затраты, руб.
wsFact.Range("T" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],'%КЗ_k эскалации'!C[-11]:C[-10],2,0),0)" '%КЗ
wsFact.Range("U" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" 'ФОТ ПЗ
wsFact.Range("V" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" 'ЭММ ПЗ
wsFact.Range("W" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" 'МТР ПЗ
wsFact.Range("X" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-3],RC[-2],RC[-1]),2)" 'ПЗ
wsFact.Range("Y" & insertRow).FormulaR1C1 = "=ROUND(RC[-1]*RC[-5],2)" 'КЗ
wsFact.Range("Z" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%КЗ_k эскалации'!R8C3-RC[-5],2)" 'ФОТ k
wsFact.Range("AA" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%КЗ_k эскалации'!R8C3-RC[-5],2)" 'ЭММ k
wsFact.Range("AB" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%КЗ_k эскалации'!R8C3-RC[-5],2)" 'МТР k
wsFact.Range("AC" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-3],RC[-2],RC[-1]),2)" 'Всего k
wsFact.Range("AD" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" 'ФОТ всего
wsFact.Range("AE" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" 'ЭММ всего
wsFact.Range("AF" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" 'МТР всего
wsFact.Range("AG" & insertRow).FormulaR1C1 = "=RC[-8]" 'КЗ
wsFact.Range("AH" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-4],RC[-3],RC[-2],RC[-1]),2)" 'Всего
wsFact.Range("AI" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-5],RC[-4],RC[-2]),2)" 'СМР
wsFact.Range("AJ" & insertRow).FormulaR1C1 = "=RC[-4]" 'МТР
wsFact.Range("AK" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'Всего с индексом-дефлятора на 2025 г.
wsFact.Range("AL" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'СМР с индексом-дефлятора на 2025 г.
wsFact.Range("AM" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'МТР с индексом-дефлятора на 2025 г.
wsFact.Range("AN" & insertRow).FormulaR1C1 = "" 'Буфер
wsFact.Range("AO" & insertRow).FormulaR1C1 = "3.5869670184836" 'КФОТ до ТКП
wsFact.Range("AP" & insertRow).FormulaR1C1 = "9.14237560208042" 'КЭММ до ТКП
wsFact.Range("AQ" & insertRow).FormulaR1C1 = "1.96204565821346" 'КМТР до ТКП
wsFact.Range("AR" & insertRow).FormulaR1C1 = "3.59699312139276" 'К КЗ РХИ
wsFact.Range("AS" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" 'ФОТ всего
wsFact.Range("AT" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" 'ЭММ всего
wsFact.Range("AU" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" 'МТР всего
wsFact.Range("AV" & insertRow).FormulaR1C1 = "=ROUND(RC[-23]*RC[-4],2)" 'КЗ
wsFact.Range("AW" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-4],RC[-3],RC[-2],RC[-1]),2)" 'Всего
wsFact.Range("AX" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-5],RC[-4],RC[-2]),2)" 'СМР
wsFact.Range("AY" & insertRow).FormulaR1C1 = "=RC[-4]" 'МТР
wsFact.Range("AZ" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'Всего с индексом-дефлятора на 2025 г.
wsFact.Range("BA" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'СМР с индексом-дефлятора на 2025 г.
wsFact.Range("BB" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" 'МТР с индексом-дефлятора на 2025 г.

        ' Копируем диапазон "P & (insertRow - 1) : BD & (insertRow - 1)" и вставляем его KolProv раз
If KolProv > 0 Then
    ' Определяем диапазон для копирования
    Dim sourceRange As Range
    Set sourceRange = wsFact.Range("A" & (insertRow) & ":BD" & (insertRow))
        ' Копируем диапазон
    sourceRange.Copy
           ' Вставляем скопированный диапазон KolProv раз
    wsFact.Range("A" & (insertRow) & ":BD" & (insertRow + KolProv - 1)).PasteSpecial Paste:=xlPasteAll
End If
' Удали под Range("A" & insertRow & ":O" & insertRow + copyRange.Rows.Count - 1)
  wsFact.Range("A" & (insertRow) & ":O" & (insertRow + KolProv - 1)).ClearContents
    
' Вставляем данные как значения без форматирования
wsFact.Range("A" & insertRow & ":O" & insertRow + copyRange.Rows.Count - 1).Value = copyRange.Value
                                                       
' Очищаем буфер обмена
Application.CutCopyMode = False
wsFact.Range("O2").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;O4:O" & CStr(insertRow + KolProv - 1) & ")"
       
' Если на wsFact. в столбце E есть значение "2", то в соответствующую строку столбца Z и AA вставить "0", т.к. на 2 пакете коэф. 0 (не 1).
Dim checkRangeFact As Range
Set checkRangeFact = wsFact.Range("E" & insertRow & ":E" & (insertRow + KolProv - 1))
Dim irr As Long
For irr = 1 To checkRangeFact.Rows.Count
    If checkRangeFact.Cells(irr, 1).Value = "2" Then
        wsFact.Cells(insertRow + irr - 1, "Z").Value = "0"
        wsFact.Cells(insertRow + irr - 1, "AA").Value = "0"
    End If
Next irr






    ' Применяем формулу к первому столбцу вставленного массива
    Set formulaRange = wsFact.Range("A" & insertRow & ":A" & insertRow + copyRange.Rows.Count - 1)
    formulaRange.Formula = "=ROW()-3"
    
    ' Находим номер строки в столбце Q на листе "Сводная по СМУ", где находится значение s22
    Set wsSvodSMU = P3.Sheets("Сводная по СМУ")
    Set FoundCell2 = wsSvodSMU.Columns("Q").Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
        
    ' Если значение найдено, копируем массив из столбца S на листе "Проверка" и вставляем на лист "Сводная по СМУ"
    If Not FoundCell2 Is Nothing Then
        ' Определяем диапазон для копирования из столбца S
        Set copyRangeS = Prov.Range("S1:S8") '  Измени, если изменится количество СМУ   S8
        ' Определяем диапазон для вставки на лист "Сводная по СМУ" показателей из Оперфакта из первого листа
        Set insertRangeS = wsSvodSMU.Range("W" & FoundCell2.Row + 3 & ":W" & FoundCell2.Row + 10)  '  Измени, если изменится количество СМУ   +10
        ' Копируем и вставляем данные
        copyRangeS.Copy
        insertRangeS.PasteSpecial xlPasteValues
        ' Очищаем буфер обмена
        Application.CutCopyMode = False
         ' копируем из Проверка столбец R ищем "Выполнено всего" + 1 столбец вправо
'         и из него вставляем в лист wsSvodSMU в столбце AC ищем "Свод (Р) ОФ" + 1 столбец вправо, т.е. напротив и сюда вставляем
    Else
        MsgBox "Отчетный день " & s22 & " не найден на листе 'Сводная по СМУ'.", vbExclamation
    End If
          
        
        ' Найти первую ячейку в 1 столбце с заливкой для даты 1го дня. Переменная для хранения номера строки с заливкой
Dim filledRow As Long
filledRow = 0
' Переменная для проверки наличия заливки
Dim hasFilledCell As Boolean
hasFilledCell = False
' Переменная для хранения цвета заливки
Dim fillColor As Long
fillColor = RGB(218, 238, 243) 'синий
' Находим последнюю заполненную ячейку в столбце A
P3.Sheets("Сводная по Прорабам").Activate
Set wsPro = P3.Sheets("Сводная по Прорабам")
lastRowPro = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row
' Проходим по ячейкам столбца A с 6 строки до последней заполненной
For i = 6 To lastRowPro
    ' Проверяем заливку ячейки
    If wsPro.Cells(i, "A").Interior.Color = fillColor Then
        filledRow = i
        hasFilledCell = True
        Exit For
    End If
Next i
' Если заливка не найдена, выводим сообщение и завершаем выполнение
If Not hasFilledCell Then
    MsgBox "На листе 'Сводная по Прорабам' отсутствует заливка ячеек RGB(218, 238, 243) в столбце А. Верните прежнее оформление.", vbExclamation
    Exit Sub
End If
'' В зависимости от номера строки с заливкой выполняем соответствующие действия
Select Case filledRow
    Case 6
        ' Вставляем две пустые строки после 5 строки
        wsPro.Rows("6:7").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' Убираем заливку вставленных строк
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case 7
        ' Вставляем одну пустую строку после 5 строки
        wsPro.Rows("6:6").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' Убираем заливку вставленной строки
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case 8
        ' Убираем заливку строк
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case Is >= 9
        ' Удаляем строки с 6 по (filledRow - 3)
        wsPro.Rows("6:" & (filledRow - 3)).Delete Shift:=xlUp
        ' Убираем заливку строк
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
End Select

'' Очищаем содержимое ячеек A6:B6
wsPro.Range("A6:B7").ClearContents
'
' Применяем формулу к ячейке D6
wsPro.Range("D6").FormulaLocal = "=ЕСЛИОШИБКА(СУММЕСЛИМН(D$9:D$" & lastRowPro + 999 & ";$A$9:$A$" & lastRowPro + 999 & ";$A6;$B$9:$B$" & lastRowPro + 999 & ";$B6);0)"
 wsPro.[D6].Copy
 wsPro.Range("D6:I6").PasteSpecial Paste:=xlPasteFormulas
' Применяем формулу к ячейке D6
wsPro.Range("K6").FormulaLocal = "=ЕСЛИОШИБКА(СУММЕСЛИМН(K$9:K$" & lastRowPro + 999 & ";$A$9:$A$" & lastRowPro + 999 & ";$A6;$B$9:$B$" & lastRowPro + 999 & ";$B6);0)"
 wsPro.[K6].Copy
 wsPro.Range("K6:P6").PasteSpecial Paste:=xlPasteFormulas
 
' Окрашиваем диапазон D4:I4 в C синий цвет
    With wsPro.Range("D4:I4").Interior
        .Color = RGB(218, 238, 243) ' С cиний цвет
        .Pattern = xlSolid ' Сплошная заливка
    End With
 
 
 
' На листе Факт ФО на текущий день начиная с ячейки C4 до нижней заполненной ячейки в столбце D запоминаем данные из данного массива
' и присваиваем эти данные в переменную. Далее на листе Факт ФО на текущий день делать ничего не нужно, а работаем лишь с переменной, на которую назначили массив.
' Нужно с данными, присвоенными переменной удалить дубликаты, с учетом двух столбцов в массиве. Подсчитать количество строк, оставшихся после удаления дубликатов и
'  скопировать содержимое строки 6  на лист Сводная по Прорабам и  вставить её такое же количество раз на лист Сводная по Прорабам после строки 6, сколько раз мы по подсчитали после удаления дубликатов.
'Далее нужно вставить данные из массива, где верхней левой ячейков вставки будет A6 на листе Сводная по Прорабам
' Определяем диапазон данных на листе "Факт ФО на текущий день"
Set DataRange = wsFact.Range("B4:D" & wsFact.Cells(wsFact.Rows.Count, "C").End(xlUp).Row)
dataArray = DataRange.Value ' Загружаем данные в массив

' Создаем коллекцию для хранения уникальных данных
Set uniqueData = New Collection

' Удаляем дубликаты, учитывая два столбца (C и D)
On Error Resume Next
For i = 1 To UBound(dataArray, 1)
    ' Создаем уникальный ключ из столбцов C и D (индексы 2 и 3 в массиве, так как B4:D начинается с B)
    Key = dataArray(i, 2) & "|" & dataArray(i, 3) ' Уникальный ключ из столбцов C и D
    uniqueData.Add Key, Key ' Добавляем ключ в коллекцию (дубликаты будут игнорироваться)
Next i
On Error GoTo 0

' Подсчитываем количество уникальных строк
insertRowsCount = uniqueData.Count

' Копируем строку 6 на листе "Сводная по Прорабам"
wsPro.Rows(6).Copy
 wsPro.Rows("6:" & insertRowsCount + 5 - 1).Insert Shift:=xlDown ' Вставляем столько раз, сколько уникальных строк

' Очищаем буфер обмена
Application.CutCopyMode = False

' Вставляем уникальные данные на лист "Сводная по Прорабам", начиная с A6
For i = 1 To uniqueData.Count
    ' Разделяем ключ обратно на два столбца
    Dim splitData() As String
    splitData = Split(uniqueData(i), "|")

    ' Вставляем данные в ячейки A и B
    wsPro.Cells(6 + i - 1, "A").Value = splitData(0) ' Столбец C (первая часть ключа)
    wsPro.Cells(6 + i - 1, "B").Value = splitData(1) ' Столбец D (вторая часть ключа)
Next i
  'сортируем данные от A6 до (B insertRowsCount + 5), где insertRowsCount +5 это нижняя правая ячейка массива для сортировки по второму столбцу от А до Я.
   ' затем сортируем данные по первому столбцу от А до Я.+
    Range("A6").Select
    ActiveWindow.SmallScroll Down:=78
    wsPro.Range("A6:B" & insertRowsCount + 5).Select
    wsPro.Sort.SortFields.Clear
    wsPro.Sort.SortFields.Add2 Key:= _
        Range("B6:B" & insertRowsCount + 5), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wsPro.Sort
        .SetRange Range("A5:B" & insertRowsCount + 5)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    wsPro.Sort.SortFields.Clear
   wsPro.Sort.SortFields.Add2 Key:= _
        Range("A6:A" & insertRowsCount + 5), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With wsPro.Sort
        .SetRange Range("A5:B" & insertRowsCount + 5)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

                                                                ' выделить массив
                                                                wsPro.Range("A6:B" & insertRowsCount + 5).Select
    ' doober выделить цветом почти похожие частично совпадающие ФИО
    Dim rng As Range, res As Double, key1 As String, key2 As String
    Set rng = Selection
    dx = rng
    For n = 1 To UBound(dx)
        key1 = dx(n, 2)
        For i = 1 To UBound(dx)
            If i <> n Then
                key2 = dx(i, 2)
                If key2 <> "" And key2 <> "" Then
                    If Simil(key1, key2) >= 0.7 Then
                        rng(i, 2).Interior.Color = 13431551
                        If i > n Then Exit For
                    End If
                End If
            End If
        Next
    Next
    ' Указываем лист и диапазон
insertRowsCount = wsPro.Cells(wsPro.Rows.Count, "B").End(xlUp).Row - 5 ' Определяем количество строк (исключая заголовки)
' Копируем данные из диапазона в массив
dataArr = wsPro.Range("B6:B" & insertRowsCount + 5).Value ' Диапазон для проверки (начиная с B6)
' Цикл по ячейкам в массиве
For yzy = 1 To UBound(dataArr, 1) ' Перебираем строки массива
    ' Проверяем, залита ли текущая ячейка цветом 13431551
    If wsPro.Cells(yzy + 5, 2).Interior.Color = 13431551 Then ' Учитываем смещение от B6
        ' Проверяем, есть ли залитая ячейка сверху или снизу
        Dim hasNeighbor As Boolean
        hasNeighbor = False
        ' Проверка сверху (если это не первая ячейка)
        If yzy > 1 Then
            If wsPro.Cells(yzy + 4, 2).Interior.Color = 13431551 Then ' Ячейка сверху
                hasNeighbor = True
            End If
        End If
        ' Проверка снизу (если это не последняя ячейка)
        If yzy < UBound(dataArr, 1) Then
            If wsPro.Cells(yzy + 6, 2).Interior.Color = 13431551 Then ' Ячейка снизу
                hasNeighbor = True
            End If
        End If
        ' Если нет соседних залитых ячеек, перекрашиваем текущую в прозрачный фон
        If Not hasNeighbor Then
            wsPro.Cells(yzy + 5, 2).Interior.ColorIndex = xlNone ' Прозрачный фон
        End If
    End If
Next yzy
   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Найти нижнюю ячейку в столбце A, содержащую заливку fillColor
Dim filledRow2 As Long
filledRow2 = 0
' Переменная для проверки наличия заливки
Dim hasFilledCell2 As Boolean
hasFilledCell2 = False
' Переменная для хранения цвета заливки
Dim fillColor2 As Long
fillColor2 = RGB(218, 238, 243) 'синий

' Находим последнюю заполненную ячейку в столбце A
P3.Sheets("Сводная по Прорабам").Activate
'Set wsPro = P3.Sheets("Сводная по Прорабам")
LastRowPro2 = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row

' Проходим по ячейкам столбца A с 6 строки до последней заполненной
For i = 6 To LastRowPro2
    ' Проверяем заливку ячейки
    If wsPro.Cells(i, "A").Interior.Color = fillColor Then
        filledRow2 = i
        hasFilledCell2 = True
    End If
Next i

' Если заливка не найдена, выводим сообщение и завершаем выполнение
If Not hasFilledCell2 Then
    MsgBox "На листе 'Сводная по Прорабам' отсутствует заливка ячеек RGB(218, 238, 243) в столбце А. Верните прежнее оформление.", vbExclamation
    Exit Sub
End If
    
   ' Копируем строки с filledRow2 до filledRow2 + 3 на листе "Сводная по Прорабам"
wsPro.Rows(filledRow2 & ":" & filledRow2 + 1).Copy
wsPro.Rows(LastRowPro2 + 2).Insert Shift:=xlDown
Application.CutCopyMode = False

 ' очистим А:С  с  LastRowPro2 + 3 до LastRowPro2 + 3 на листе "Сводная по Прорабам"
 wsPro.Range("A" & LastRowPro2 + 3 & ":C" & LastRowPro2 + 3 & "").ClearContents
 
'  На листе Факт ФО на текущий день начиная с ячейки B4 до нижней заполненной ячейки в столбце D запоминаем данные из данного массива
' и присваиваем эти данные в переменную. Далее на листе Факт ФО на текущий день делать ничего не нужно, а работаем лишь с переменной, на которую назначили массив.
'Нужно с данными, присвоенными переменной оставить для дальнейшей работы те строки,
'которые содержат в первом столбце массива (столбец B) значение, которой ранее в коде мы присвоили переменной s22.
'Далее нужно удалить дубликаты, с учетом двух столбцов С и D в массиве. Подсчитать количество строк, оставшихся после удаления дубликатов и
'  скопировать содержимое массива от A до P столбцов  на листе Сводная по Прорабам в строке  LastRowPro + 2 и  вставить его такое же количество раз на лист Сводная по Прорабам
'после строки  LastRowPro + 2, сколько раз мы по подсчитали после удаления дубликатов.
'Далее нужно вставить данные из массива, где верхней левой ячейкой вставки будет LastRowPro + 3 на листе Сводная по Прорабам
' Из массива dataArray выбираем только те строки, которые расположены напротив значения s22
' На листе "Факт ФО на текущий день" начиная с ячейки B4 до нижней заполненной ячейки в столбце D запоминаем данные из данного массива
Set DataRange = wsFact.Range("B4:D" & wsFact.Cells(wsFact.Rows.Count, "D").End(xlUp).Row)
dataArray = DataRange.Value ' Загружаем данные в массив

' Создаем новый массив для хранения строк, где в первом столбце (столбец B) содержится значение s22
Dim filteredData() As Variant
Dim filteredRowCount As Long
filteredRowCount = 0
'
' Проходим по массиву и отбираем строки, где в первом столбце (столбец B) содержится значение s22
For i = 1 To UBound(dataArray, 1)
    If dataArray(i, 1) = s22 Then
        filteredRowCount = filteredRowCount + 1
        ReDim Preserve filteredData(1 To 3, 1 To filteredRowCount)
        filteredData(1, filteredRowCount) = dataArray(i, 1) ' Столбец B
        filteredData(2, filteredRowCount) = dataArray(i, 2) ' Столбец C
        filteredData(3, filteredRowCount) = dataArray(i, 3) ' Столбец D
    End If
Next i

' Создаем коллекцию для хранения уникальных данных
Set uniqueData = New Collection

' Удаляем дубликаты, учитывая два столбца (C и D)
On Error Resume Next
For i = 1 To filteredRowCount
    ' Создаем уникальный ключ из столбцов C и D (индексы 2 и 3 в массиве)
    Key = filteredData(2, i) & "|" & filteredData(3, i) ' Уникальный ключ из столбцов C и D
    uniqueData.Add Key, Key ' Добавляем ключ в коллекцию (дубликаты будут игнорироваться)
Next i
On Error GoTo 0

' Подсчитываем количество уникальных строк
insertRowsCount = uniqueData.Count

'' Копируем строку и вставляем её insertRowsCount раз
    wsPro.Rows(LastRowPro2 + 3).Copy
'    Вставляем скопированную строку insertRowsCount раз
wsPro.Rows(LastRowPro2 + 3 & ":" & LastRowPro2 + 3 + insertRowsCount - 2).Insert Shift:=xlDown


' Вставляем уникальные данные на лист "Сводная по Прорабам", начиная с LastRowPro2 + 3
For i = 1 To uniqueData.Count
    ' Разделяем ключ обратно на два столбца
    Dim splitData5() As String
    splitData5 = Split(uniqueData(i), "|")

    ' Вставляем данные в ячейки A и B
    wsPro.Cells(LastRowPro2 + 3 + i - 1, "A").Value = splitData5(0) ' Столбец C (первая часть ключа)
    wsPro.Cells(LastRowPro2 + 3 + i - 1, "B").Value = splitData5(1) ' Столбец D (вторая часть ключа)
Next i
  ' Вставляем данные в ячейки С
 wsPro.Range("C" & LastRowPro2 + 3 & ":C" & LastRowPro2 + 3 + insertRowsCount - 1) = s22
 ' Сортируем  массив wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1) от А до Я сначала по 2 столбцу, затем по 1 столбцу
' Сортируем массив от A до Я сначала по второму столбцу (B), затем по первому столбцу (A)
With wsPro.Sort
    .SortFields.Clear
    ' Сортировка по второму столбцу (B)
    .SortFields.Add2 Key:=wsPro.Range("B" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     ' Устанавливаем диапазон для сортировки
    .SetRange wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1)
    .Header = xlNo ' Указываем, что заголовков нет
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
        .SortFields.Clear
    ' Сортировка по первому столбцу (A)
    .SortFields.Add2 Key:=wsPro.Range("A" & LastRowPro2 + 3 & ":A" & LastRowPro2 + 3 + insertRowsCount - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Устанавливаем диапазон для сортировки
    .SetRange wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1)
    .Header = xlNo ' Указываем, что заголовков нет
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

wsPro.UsedRange  'сбросить результат с последней ячейкой, строкой
Dim LastRow5 As Long
LastRow5 = Cells(Rows.Count, "A").End(xlUp).Row
' Устанавливаем область печати
'    область печати: вертикаль - последняя строка, горизонталь - восьмой столбец h
    LastRow5 = wsPro.UsedRange.Row + wsPro.UsedRange.Rows.Count - 1
    wsPro.PageSetup.PrintArea = wsPro.Range(Cells(1, 1), Cells(LastRow5, 16)).Address

    ' Удаляем все строки ниже LastRow5
If LastRow5 < Rows.Count Then
     wsPro.Rows(LastRow5 + 1 & ":" & Rows.Count).Delete
End If


'Делаем дату. В ячейке wsPro.Range("A" & LastRowPro2 + 2) находится число. Нужно это число запомнить, прибавить к нему 1 и результат вернуть в эту же ячейку
' Либо В ячейке wsPro.Range("A" & LastRowPro + 2) находится дата, нужно запомнить ее, присвоив переменной, затем
'меняем первые два числа слева на значение из переменной s22, где s22 это любые два числа
' Определяем целевую ячейку
Dim targetCell As Range
Set targetCell = wsPro.Range("A" & LastRowPro2 + 2)

' Запоминаем текущее значение (дату) в переменную
Dim originalDate As String
originalDate = CStr(targetCell.Value)
' Проверяем, что длина строки позволяет заменить первые два символа
If Len(originalDate) >= 2 Then
    ' Преобразуем s22 в число для проверки
    Dim s22Value As Long
    s22Value = CLng(s22)
    ' Проверяем, находится ли значение в диапазоне от 1 до 9
    If s22Value >= 1 And s22Value <= 9 Then
        ' Форматируем s22 как текст с ведущим нулём (например, "01", "02", ...)
        Dim formattedS22 As String
        formattedS22 = Format(s22Value, "00")
        ' Заменяем первые два числа на отформатированное значение
        Dim modifiedDate As String
        modifiedDate = formattedS22 & Mid(originalDate, 3)
        ' Обновляем значение ячейки
        targetCell.Value = modifiedDate
    Else
        ' Если значение не в диапазоне от 1 до 9, просто заменяем первые два символа
        Dim modifiedDate4 As String
        modifiedDate4 = s22 & Mid(originalDate, 3)
        ' Обновляем значение ячейки
        targetCell.Value = modifiedDate4
    End If
Else
    MsgBox "Дата слишком короткая для замены первых двух символов!", vbExclamation
End If
  
    ' формулу в шапку текущего дня промежуточные итоги
     wsPro.Range("D" & LastRowPro2 + 2).FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;D" & LastRowPro2 + 3 & ":D" & LastRowPro2 + 3 + insertRowsCount - 1 & ")"
     wsPro.Range("D" & LastRowPro2 + 2).Copy
     wsPro.Range("D" & LastRowPro2 + 2 & ":I" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas
     wsPro.Range("K" & LastRowPro2 + 2 & ":P" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas

    
 ' Найти предпоследнюю ячейку в 1 столбце с заливкой для даты вчерашнего дня
filledRow3 = 0
' Переменная для проверки наличия заливки
hasFilledCell3 = False
Dim fillColor3 As Long
fillColor3 = RGB(218, 238, 243) 'синий
' Проходим по ячейкам столбца A с 6 строки до последней заполненной
For i = LastRowPro2 To 6 Step -1
    ' Проверяем заливку ячейки
    If wsPro.Cells(i, "A").Interior.Color = fillColor3 Then
        filledRow3 = i
        hasFilledCell3 = True
        Exit For
    End If
Next i


' Найти  ' Найти первую ячейку в 1 столбце с заливкой для даты 1го дня. Переменная для хранения номера строки с заливкой
filledRow5 = 0
' Переменная для проверки наличия заливки
hasFilledCell5 = False
Dim fillColor5 As Long
fillColor5 = RGB(218, 238, 243) 'синий
' Проходим по ячейкам столбца A с 6 строки до последней заполненной
For i = 6 To LastRow5
    ' Проверяем заливку ячейки
    If wsPro.Cells(i, "A").Interior.Color = fillColor5 Then
        filledRow5 = i
        hasFilledCell5 = True
        Exit For
    End If
Next i







        
     wsPro.Range("D4").FormulaLocal = "=ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;D6:D" & filledRow5 - 2 & ")"
     wsPro.Range("D4").Copy
     wsPro.Range("D4:I4").PasteSpecial Paste:=xlPasteFormulas
     wsPro.Range("K4:P4").PasteSpecial Paste:=xlPasteFormulas
'    ' формула проверки за полями столбец U сумма всех Итого
wsPro.Range("U" & LastRowPro2 + 2 & ":Z" & LastRowPro2 + 3 + insertRowsCount - 1).ClearContents

wsPro.Range("U" & LastRowPro2 + 2).Formula = "=D" & (LastRowPro2 + 2)
wsPro.Range("U" & LastRowPro2 + 2).Copy
wsPro.Range("U" & LastRowPro2 + 2 & ":Z" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas

wsPro.Range("U" & LastRowPro2 + 3).Formula = "='Сводная по СМУ'!C" & (FoundCell2.Row + 11)  '  Измени, если изменится количество СМУ +11
wsPro.Range("U" & LastRowPro2 + 3).Copy
wsPro.Range("U" & LastRowPro2 + 3 & ":Z" & LastRowPro2 + 3).PasteSpecial Paste:=xlPasteFormulas

wsPro.Range("U" & LastRowPro2 + 4).Formula = "=U" & (LastRowPro2 + 3) & "=U" & (LastRowPro2 + 2) ' истина ложь
wsPro.Range("U" & LastRowPro2 + 4).Copy ' истина ложь
wsPro.Range("U" & LastRowPro2 + 4 & ":Z" & LastRowPro2 + 4).PasteSpecial Paste:=xlPasteFormulas ' истина ложь

' Нужно найти для проверки накопительный объём на предыдущий день, который находится в ячейкеах U:Z
'для этого мы ищем снизу вверх второе найденное значение   fillColor4 = RGB(218, 238, 243)
Dim filledRow4 As Long ' вчерашний предыдущий день
Dim fillColor4 As Long
fillColor4 = RGB(218, 238, 243) ' Цвет заливки 'синий
Dim foundCount As Long
foundCount = 0 ' Счетчик найденных строк с заливкой
' Находим последнюю заполненную ячейку в столбце A
LastRowPro7 = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row
' Проходим по ячейкам столбца A снизу вверх
For i = LastRowPro7 To 1 Step -1
    ' Проверяем заливку ячейки
    If wsPro.Cells(i, "A").Interior.Color = fillColor4 Then
        foundCount = foundCount + 1 ' Увеличиваем счетчик найденных строк
        If foundCount = 2 Then ' Если это второе найденное значение
            filledRow4 = i ' Присваиваем номер строки переменной
            Exit For ' Выходим из цикла
        End If
    End If
Next i
' Проверяем, найдено ли второе значение
If foundCount < 2 Then
    MsgBox "Второе значение с заливкой RGB(218, 238, 243) в столбце A не найдено!", vbExclamation
'    Exit Sub
End If

                                                      ' здесь важно, чтобы формула определила текущие объемы с учетом добавленных за сегодняшний день поэтому  Application.Calculation = xlCalculationAutomatic
                                                       Application.Calculation = xlCalculationAutomatic
                                                    wsPro.Range("U" & LastRowPro2 + 5).Formula = "=U" & (LastRowPro2 + 2) & "+U" & (filledRow4 + 3) '
                                                    wsPro.Range("U" & LastRowPro2 + 5).Copy '
                                                    wsPro.Range("U" & LastRowPro2 + 5 & ":Z" & LastRowPro2 + 5).PasteSpecial Paste:=xlPasteFormulas '
                                                    
                                                    wsPro.Range("U" & LastRowPro2 + 6 & ":Z" & LastRowPro2 + 6).Value = wsSvodSMU.Range("C13:H13").Value   '  Измени, если изменится количество СМУ C13:H13
                                                    wsPro.Range("U" & LastRowPro2 + 7).Formula = "=U" & (LastRowPro2 + 6) & "=U" & (LastRowPro2 + 5) ' истина ложь
                                                    wsPro.Range("U" & LastRowPro2 + 7).Copy ' истина ложь
                                                    wsPro.Range("U" & LastRowPro2 + 7 & ":Z" & LastRowPro2 + 7).PasteSpecial Paste:=xlPasteFormulas ' истина ложь




'
' Ищем "Выполнено всего" на листе "Проверка" в столбце R
Set FoundCell5 = Vb.Sheets("Проверка").Columns("R").Find(What:="Выполнено всего", LookIn:=xlValues, LookAt:=xlPart)
' Если найдено "Выполнено всего", копируем значение из соседней ячейки (на 1 столбец вправо)
If Not FoundCell5 Is Nothing Then
    ' Копируем значение из ячейки справа от "Выполнено всего"
    copyValue = FoundCell5.Offset(0, 1).Value
    ' Ищем "Свод (Р) ОФ" на листе "Сводная по СМУ" в столбце AC
'    Set wsSvodSMU = P3.Sheets("Сводная по СМУ")
    Set FoundCell55 = wsSvodSMU.Columns("AC").Find(What:="Свод (Р) ОФ", LookIn:=xlValues, LookAt:=xlPart)
    ' Если найдено "Свод (Р) ОФ", вставляем скопированное значение в соседнюю ячейку (на 1 столбец вправо)
    If Not FoundCell55 Is Nothing Then
        FoundCell55.Offset(0, 1).Value = copyValue    ' Вставляем значение в ячейку справа от "Свод (Р) ОФ"
'        FoundCell55.Offset(1, 1).Value = Delta        ' Вставляем значение Delta в ячейку ниже и справа от "Свод (Р) ОФ"
        ' Проверяем, если значение в FoundCell55.Offset(1, 1) меньше -10 или больше 10
        
        
        
        
        
        
        
        If FoundCell55.Offset(1, 1).Value < -10 Or FoundCell55.Offset(1, 1).Value > 10 Then
                ' Применяем красный фон (RGB(219, 179, 182))
                 wsPro.Cells(filledRow5 - 1, "B").Interior.Color = RGB(219, 179, 182) ' красный
            ' Вставляем сообщение в ячейку
          wsPro.Cells(filledRow5 - 1, "B") = "Проверьте на листе 'Сводная по СМУ' значения в столбце W 'Свод (Р) ОФ' по управлениям"
'             "Проверьте на листе 'Сводная по СМУ' значение по накопительной в столбце AD 'Свод (Р) ОФ'"
        End If
    Else
        MsgBox "Значение 'Свод (Р) ОФ' не найдено на листе 'Сводная по СМУ'.", vbExclamation
    End If
Else
    MsgBox "Значение 'Выполнено всего' не найдено на листе 'Проверка'.", vbExclamation
End If
' Определяем диапазон для проверки
Set checkRange = wsSvodSMU.Range("W" & FoundCell2.Row + 3 & ":W" & FoundCell2.Row + 9).Offset(0, 1)
Dim hasExceeded As Boolean
hasExceeded = False ' Флаг для проверки наличия значений > 1 или < -1

' Перебираем каждую ячейку в диапазоне
For Each cell In checkRange
    If Not IsEmpty(cell) And IsNumeric(cell.Value) Then
        Dim cellValue As Double
        cellValue = CDbl(cell.Value) ' Преобразуем значение в число
        If cellValue > 1 Or cellValue < -1 Then
            hasExceeded = True ' Если найдено значение > 1 или < -1, устанавливаем флаг
            Exit For ' Прекращаем цикл, так как условие выполнено
        End If
    End If
Next cell
















' Если хотя бы одно значение больше 1 или меньше -1, выполняем действия
If hasExceeded Then
    ' Выделяем ячейку на листе "Сводная по Прорабам" в столбце D, строка filledRow - 1
  ' Применяем красный фон (RGB(219, 179, 182))
                 wsPro.Cells(filledRow5 - 1, "D").Interior.Color = RGB(219, 179, 182) ' красный
            ' Вставляем сообщение в ячейку
          wsPro.Cells(filledRow5 - 1, "D") = "Проверьте на листе 'Сводная по СМУ' значения в столбце W 'Свод (Р) ОФ' по управлениям"
End If
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
   ' Перебираем каждую ячейку в диапазоне C12:H12 на листе wsSvodSMU                ПРОВЕРКА
For Each CellsMU In wsSvodSMU.Range("C13:H13")
    ' Находим соответствующую ячейку в диапазоне D4:I4 на листе wsPro
    Dim cellProrabs As Range
    Set cellProrabs = wsPro.Range("D4").Offset(0, CellsMU.Column - wsSvodSMU.Range("C13").Column)
        ' Получаем значения из ячеек и округляем их до 2 знаков после запятой
    If IsNumeric(CellsMU.Value) Then
        valueSMU = Round(CellsMU.Value, 2)
    Else
        valueSMU = 0 ' Если значение не число, присваиваем 0
    End If
        If IsNumeric(cellProrabs.Value) Then
        valueProrabs = Round(cellProrabs.Value, 2)
    Else
        valueProrabs = 0 ' Если значение не число, присваиваем 0
    End If
        ' Сравниваем округленные значения
    If valueSMU <> valueProrabs Then
        ' Если значения не равны, окрашиваем соответствующую ячейку на листе "Сводная по Прорабам"
        cellProrabs.Interior.Color = RGB(219, 179, 182)
    End If
Next CellsMU
       
       
       
       
       
       
       
       
       
' Устанавливаем цвет заливки
wsPro.Range("J6:J" & filledRow5 - 2).Interior.Color = RGB(146, 205, 220)
' Устанавливаем шаблон заливки
wsPro.Range("J6:J" & filledRow5 - 2).Interior.Pattern = xlSolid



        ' Проверка пустых ячеек в указанных столбцах
                                                                                Dim checkRow As Long
                                                                                Dim emptyCellsMsg As String
                                                                                emptyCellsMsg = ""
                                                                                For checkRow = insertRow To insertRow + KolProv - 1
                                                                                    ' Проверка столбца O (Объем)
                                                                                    If IsEmpty(wsFact.Range("O" & checkRow)) Then
                                                                                        wsFact.Range("O" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует объём в ячейке O" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца C (СМУ)
                                                                                    If IsEmpty(wsFact.Range("C" & checkRow)) Then
                                                                                        wsFact.Range("C" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует СМУ в ячейке C" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца D (ФИО)
                                                                                    If IsEmpty(wsFact.Range("D" & checkRow)) Then
                                                                                        wsFact.Range("D" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует ФИО в ячейке D" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца E (Номер пакета)
                                                                                    If IsEmpty(wsFact.Range("E" & checkRow)) Then
                                                                                        wsFact.Range("E" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует номер пакета в ячейке E" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца F (Фаза)
                                                                                    If IsEmpty(wsFact.Range("F" & checkRow)) Then
                                                                                        wsFact.Range("F" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует Фаза в ячейке F" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца G (Титул)
                                                                                    If IsEmpty(wsFact.Range("G" & checkRow)) Then
                                                                                        wsFact.Range("G" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует Титул в ячейке G" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца L (ЕР)
                                                                                    If IsEmpty(wsFact.Range("L" & checkRow)) Then
                                                                                        wsFact.Range("L" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует ЕР в ячейке L" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' Проверка столбца N (Ед. Изм.)
                                                                                    If IsEmpty(wsFact.Range("N" & checkRow)) Then
                                                                                        wsFact.Range("N" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "На листе ""Факт ФО..."" отсутствует Ед. Изм. в ячейке N" & checkRow & vbCrLf
                                                                                    End If
                                                                                Next checkRow
                                                                                ' Вывод сообщения пользователю, если найдены пустые ячейки
                                                                                If emptyCellsMsg <> "" Then
                                                                                    MsgBox emptyCellsMsg, vbExclamation, "Пустые ячейки"
                                                                                End If

' Если на wsFact. в столбцах P Q R S  формулы, которые находятся в данных ячейках будут визуально выдавать значение 0, то
' Выдать сообщение "На листе ""Факт ФО..."" ЕР не представлена в Прайсе в ячейке " и указать в сообщении номер ячейки (ячеек) и выделить ячейку RGB(219, 179, 182) в столбце L
'' Если на wsFact. в столбцах P Q R S  формулы, которые находятся в данных ячейках будут визуально выдавать значение 0, то
' Выдать сообщение "На листе ""Факт ФО..."" ЕР не представлена в Прайсе в ячейке " и указать в сообщении номер ячейки (ячеек) и выделить ячейку RGB(219, 179, 182) в столбце L
' Объявление переменных
Dim checkRowFormula As Long
Dim zeroFormulaCellsMsg As String ' Переменная для хранения сообщений об ошибках
zeroFormulaCellsMsg = "" ' Инициализация переменной

For checkRowFormula = insertRow To insertRow + KolProv - 1
    ' Проверяем ячейки в столбцах P, Q, R, S
    Dim cellFactP As Range, cellFactQ As Range, cellFactR As Range, cellFactS As Range
    Set cellFactP = wsFact.Cells(checkRowFormula, "P")
    Set cellFactQ = wsFact.Cells(checkRowFormula, "Q")
    Set cellFactR = wsFact.Cells(checkRowFormula, "R")
    Set cellFactS = wsFact.Cells(checkRowFormula, "S")
    
    ' Проверяем, если все ячейки содержат формулы и их значения равны 0
    If cellFactP.HasFormula And cellFactQ.HasFormula And cellFactR.HasFormula And cellFactS.HasFormula Then
        If cellFactP.Value = 0 And cellFactQ.Value = 0 And cellFactR.Value = 0 And cellFactS.Value = 0 Then
            ' Добавляем информацию об ошибке
            zeroFormulaCellsMsg = zeroFormulaCellsMsg & "На листе ""Факт ФО..."" ЕР не представлена в Прайсе в ячейках P" & checkRowFormula & ", Q" & checkRowFormula & ", R" & checkRowFormula & ", S" & checkRowFormula & vbCrLf
            
            ' Выделяем соответствующую ячейку в столбце L светло-красным цветом
            wsFact.Cells(checkRowFormula, "L").Interior.Color = RGB(219, 179, 182)
        End If
    End If
Next checkRowFormula

' Вывод сообщения пользователю, если найдены ошибки
If zeroFormulaCellsMsg <> "" Then
    MsgBox zeroFormulaCellsMsg, vbExclamation, "Отсутствие ЕР в Прайсе"
End If




    ' в ячейку A6
    Application.GoTo Range("A6"), True
'
    Application.ScreenUpdating = True 'Включаем обновление экрана после каждого события
    Application.Calculation = xlCalculationAutomatic 'Расчёты формул - снова в автоматическом режиме
    Application.EnableEvents = True  'Включаем события
    If Workbooks.Count Then
    ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True 'Показываем границы ячеек
    End If
    Application.DisplayStatusBar = True 'Возвращаем статусную строку
    Application.DisplayAlerts = True 'Разрешаем

    MsgBox "Данные из листа ""Проверка"" добавлены в " & filePath & " на лист ""Факт ФО на текущий день"". Проверь данные на листе ""Сводная по Прорабам"", которые выделены желтизной в один массив (Это проверка на уникальность ФИО с процентом подобия 40. Также здесь проверь покраснения в строке 4 и перед 01 числом. Это проверка на совпадение расчетов с листом ""Сводная по СМУ"""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
' Функция для открытия диалогового окна выбора файла с указанием начальной папки
Function OpenFileDialog(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Выберите файл ""Р3 АГПЗ ОЗХ _ Факт..."" за отчётный период"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' Указываем начальную папку, если она задана
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            OpenFileDialog = .SelectedItems(1)
        Else
            OpenFileDialog = False
        End If
    End With
End Function


