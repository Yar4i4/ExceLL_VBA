Attribute VB_Name = "Module1"
Sub ��_������_��_���()
        Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    If Workbooks.Count Then ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    Dim Prov As Worksheet
    Set Prov = Vb.Worksheets("��������")
Vb.Worksheets("��������").Activate    ' ��������� �� ���� "��������"
If ActiveWindow.FreezePanes Then ' ���������, ���� �� ������������ �������
    ActiveWindow.FreezePanes = False    ' ���� ������������ ������� ����, ������� ��
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
    Prov.Range("D8") = "����"
    Prov.Range("E8") = "����� ������"
    Prov.Range("F8") = "����"
    Prov.Range("G8") = "����� ������"
    Prov.Range("H8") = "������������ ������"
    Prov.Range("I8") = "������"
    Prov.Range("J8") = "����� ���������"
    Prov.Range("K8") = "�������"
    Prov.Range("L8") = "���� ���������� ��������"
    Prov.Range("M8") = "�������� ��������� ��������"
    Prov.Range("N8") = "�� ���"
    Prov.Range("Q8") = "�������������"
    Prov.Range("P8") = "��� ������� / �������"
    Prov.Range("A8:Q8").Interior.Color = RGB(183, 222, 232)
    Prov.Range("A9:Q9").Interior.Color = RGB(218, 238, 243)
    Prov.Range("A10:S10").FormulaLocal = "=�������()"
    Vb.Worksheets("��������").Rows("9:9").NumberFormat = "#,##0.00"
    ' ��������� ���������� ���� ��� ������ ������
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog3(Vb.Path)
    
    ' ���������, ��� �� ������ ����
    If filePaths.Count = 0 Then
        MsgBox "���� �� ������!", vbExclamation
        Exit Sub
    End If
       
    '������������� ���������
    Dim totalBooks As Long
    Dim processedBooks As Long
    processedBooks = 0
    totalBooks = filePaths.Count
    Vb.Worksheets("�������").Cells(1, "D").Value = "����� ���������� 0 �� " & totalBooks & " ����" ' ��������� ��������
    
    ' ��������� ��� ����� ������
    Dim errorMessages As Collection
    Set errorMessages = New Collection
    
    
    
    
    
    ' ������������ ������ ��������� ����
    Dim filePath As Variant
    For Each filePath In filePaths
        ProcessFile CStr(filePath), Vb, Prov, errorMessages
        processedBooks = processedBooks + 1
        Application.ScreenUpdating = True
        Vb.Worksheets("�������").Cells(1, "D").Value = "����� ���������� " & processedBooks & " �� " & totalBooks & " ����"
        Application.ScreenUpdating = False
    Next filePath







    ' �������� �� ������ ������ � �������� D, P, Q
    ActiveSheet.UsedRange
    LastRowDPQ = Prov.Cells.SpecialCells(xlLastCell).Row
    Dim errorMessageDPQ As String
    errorMessageDPQ = ""
    ' �������� ������� D (����)
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
        errorMessageDPQ = errorMessageDPQ & "� ������� ���� ���� ������ ������." & vbCrLf
    End If
     ' �������� ������� G (�����)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("G11:G" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "� ������� ����� ���� ������ ������." & vbCrLf
    End If
    ' �������� ������� P (�������������)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("P11:P" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "� ������� ������������� ���� ������ ������." & vbCrLf
    End If
    ' �������� ������� Q (����������)
    emptyCellFound = False
    For Each cellDPQempty In Prov.Range("Q11:Q" & LastRowDPQ)
        If IsEmpty(cellDPQempty.Value) Or Trim(cellDPQempty.Value) = "" Then
            cellDPQempty.Interior.Color = RGB(219, 179, 182)
            emptyCellFound = True
        End If
    Next cellDPQempty
    If emptyCellFound Then
        errorMessageDPQ = errorMessageDPQ & "� ������� ���������� ���� ������ ������." & vbCrLf
    End If
    ' ����� ������ ���������, ���� ���� ������
    If errorMessageDPQ <> "" Then
        MsgBox "���������� ������ ������:" & vbCrLf & vbCrLf & errorMessageDPQ, vbExclamation, "������"
    End If
    Prov.Range("D11").FormulaLocal = "=�������(�������(I11;L11;K11);190)"
    Prov.Range("D11").AutoFill Destination:=Prov.Range("D11:D" & LastRowDPQ), Type:=xlFillDefault
    
    Prov.Rows("11:" & LastRowDPQ).RowHeight = 15
    Dim sCell As Range
    For Each sCell In Prov.Range("S11:S" & LastRowDPQ)
        If Not IsEmpty(sCell.Value) And Trim(sCell.Value) <> "" Then
            Prov.Range("A" & sCell.Row & ":R" & sCell.Row).Interior.Color = RGB(218, 238, 243)
            Prov.Range("D" & sCell.Row & ":D" & sCell.Row).ClearContents
        End If
    Next sCell
    
    Prov.Range("O8") = "��"
    Prov.Range("R8").Value = "��" & Chr(10) & "(�������)"
    Prov.Range("S8").Value = "��" & Chr(10) & "(�� �������)"
    Prov.Range("O9").FormulaLocal = "=�������������.�����(9;O11:O" & LastRowDPQ & ")"
    Prov.Range("R9").FormulaLocal = "=�������������.�����(9;R11:R" & LastRowDPQ & ")"
    Prov.Range("S9").FormulaLocal = "=�������������.�����(9;S11:S" & LastRowDPQ & ")"
    Prov.Range("P9").FormulaLocal = "=�������������.�����(2;O11:O" & LastRowDPQ & ")"
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
        errorMessage = "���������� ������ ��� ��������� ������:" & vbCrLf
        Dim i As Integer
        For i = 1 To errorMessages.Count
            errorMessage = errorMessage & errorMessages(i) & vbCrLf
        Next i
        MsgBox errorMessage, vbExclamation, "������"
    End If
    
 Vb.Worksheets("��������").Rows("11:11").Select ' �������� ������ 11
ActiveWindow.FreezePanes = True ' ���������� ������ � 1 �� 10
    
    
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
    
    ' ���������, ������� �� �������� ����
    If Ob Is Nothing Then
        errorMessages.Add "�� ������� ������� ����: " & filePath
        Exit Sub
    End If
    
    ' �������� ��� ����� ��� ����
    Dim fileName As String
    fileName = Right(Ob.Name, Len(Ob.Name) - InStrRev(Ob.Name, "\") - 1)
    
    ' ������ ������ ����������� ��������� ��� ������ ���� � ������� ���.�����.����
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "(\d{4})\.(\d{2})\.(\d{2})" ' ������: GGGG.MM.DD
        .Global = False
        .IgnoreCase = True
    End With
    
    ' ���� ���������� � �������� � ����� �����
    Dim match As Object
    If regex.Test(fileName) Then
        Set match = regex.Execute(fileName)
        ' ��������� ���, ����� � ���� �� ��������� ����
        Dim YearPart As String
        Dim Mes As String
        Dim Day As String
        YearPart = match(0).SubMatches(0) ' ��� (������ �����)
        Mes = match(0).SubMatches(1)     ' ����� (��� �����)
        Day = match(0).SubMatches(2)     ' ���� (��� �����)
    Else
        ' ��������� ��������� �� ������ � ���������
        errorMessages.Add "������ ���� � ����� " & fileName & " �� ��������� (��������� GGGG.MM.DD)."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' ���������, ���������� �� ���� "������ ��� ""�-�����"""
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = Ob.Worksheets("������ ��� ""�-�����""")
    On Error GoTo 0
    ' ���� ���� �� ������, ��������� ��������� �� ������
    If targetSheet Is Nothing Then
        errorMessages.Add "���� ""������ ��� ""�-�����"""" �� ������ � ����� " & fileName & "."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' ���������� ���������� � ������� ������/������� �� ���� ������ ����� Ob
    Dim ws As Worksheet
    For Each ws In Ob.Worksheets
        ' ���������� ����������, ���� �� �������
        If ws.AutoFilterMode And ws.FilterMode Then ws.ShowAllData
        ' ���������� ������� ������
        If ws.Cells.EntireRow.Hidden Then ws.Cells.EntireRow.Hidden = False
        ' ���������� ������� �������
        If ws.Cells.EntireColumn.Hidden Then ws.Cells.EntireColumn.Hidden = False
    Next ws
    
    ' ���������� ���� "������ ��� ""�-�����"""
    Ob.Worksheets("������ ��� ""�-�����""").Activate
    
    ' ����� ������ ������, ���������� "�� �� ��������������"
    Dim Rngt As Range
    Set Rngt = Cells.Find("�� �� " & Day & "." & Mes, , xlFormulas, xlWhole)
    If Rngt Is Nothing Then
        errorMessages.Add "� ����� " & fileName & " ����������� ������: ""�� �� " & Day & "." & Mes & """."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lRow As Long
    lRow = Rngt.Row ' ������, ��� ���� ������� �����
    Dim lCol As Long
    lCol = Rngt.Column ' �������, ��� ���� ������� ����� "�� ��"
    
    ' ��������� UsedRange ��� ����� "��������"
    Vb.Worksheets("��������").UsedRange
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, lCol).End(xlUp).Row
    Dim sAddress As String
    sAddress = Rngt.Address ' ����� ������, ��� ���� ������� �����
    
    ' ���������, ���� �� �������� � ������
    Dim volumeValue As Variant
    volumeValue = Ob.Worksheets("������ ��� ""�-�����""").Cells(9, lCol).Value
    If IsEmpty(volumeValue) Or volumeValue = 0 Then
        ' ���� �������� ����������� ��� ����� 0, ��������� ��������� �� ������
        errorMessages.Add "� ����� " & fileName & " ����������� ������ �� �������� ����."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' ���������� ��� ������ �������� � ��� ����� � ������� � ������� �� ����
    Ob.Worksheets("������ ��� ""�-�����""").Range(Cells(lRow + 2, lCol), "FF" & LastRow).AutoFilter Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
    
    ' ����� ������ ������, ���������� "�������������"
    Dim Rngt2 As Range
    Set Rngt2 = Cells.Find("�������������", , xlFormulas, xlWhole)
    If Rngt2 Is Nothing Then
        errorMessages.Add "� ����� " & fileName & " ����������� ������: ""�������������""."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lCol2 As Long
    lCol2 = Rngt2.Column ' �������, ��� ���� ������� ����� "�������������"
    
    ' �������� �������
    Ob.Worksheets("������ ��� ""�-�����""").Range(Columns(lCol + 2), Columns(lCol2 - 1)).Hidden = True
    Ob.Worksheets("������ ��� ""�-�����""").Range(Columns(15), Columns(lCol - 1)).Hidden = True
    Ob.Worksheets("������ ��� ""�-�����""").Columns("A:C").EntireColumn.Hidden = True
    
    ' �������� ������ � ����� �����
    Application.CutCopyMode = False
    Ob.Worksheets("������ ��� ""�-�����""").Range(Cells(11, 4), Cells(LastRow, lCol2)).Copy
    Vb.Worksheets("��������").Activate
    Vb.Worksheets("��������").UsedRange
    Dim lastRow2 As Long
    lastRow2 = Vb.Worksheets("��������").Cells.SpecialCells(xlLastCell).Row
    Vb.Worksheets("��������").Cells(lastRow2 + 1, 4).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' ��������� ����� Ob
    Ob.Close SaveChanges:=True
    
    ' ��������� UsedRange ��� ����� "��������"
    Vb.Worksheets("��������").Activate
    Vb.Worksheets("��������").UsedRange
    Dim LastRow3 As Long
    LastRow3 = Vb.Worksheets("��������").Cells.SpecialCells(xlLastCell).Row
    
    ' ��������� ������� � ��������
    Prov.Range("R" & LastRow3 + 1).FormulaLocal = "=����(O" & lastRow2 + 1 & ":O" & LastRow3 & ")"
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
        .Title = "�������� ���� ��� ����� ""������ ��� �-����� �3 ���� ��� _..."" �� ���"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        .AllowMultiSelect = True ' ��������� ����� ���������� ������
        
        ' ��������� ��������� �����, ���� ��� ������
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                filePaths.Add .SelectedItems(i) ' ��������� ������ ��������� ���� � ���������
            Next i
        End If
    End With
    
    Set OpenFileDialog3 = filePaths
End Function

