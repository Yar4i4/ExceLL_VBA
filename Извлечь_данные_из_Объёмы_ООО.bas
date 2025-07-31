Attribute VB_Name = "Module3"
' ��� �� ������ � ��������� ���� ������ ���������, �����������
Type Item
    fio As String
    number As String
End Type
Function Is_Fio(sl As String) As Item()
    Dim itm() As Item ' ��������� ������ �������� Item ��� �������� �����������
    ReDim itm(0) ' �������������� ������ � ����� ���������
    sl = Replace(sl, "-", "") ' ������� ��� ������ �� ������� ������
    Static RegExp As Object ' ��������� ����������� ���������� ��� ������� RegExp (���������� ���������)
    If RegExp Is Nothing Then ' ���� ������ RegExp ��� �� ������
        Set RegExp = CreateObject("VBScript.RegExp") ' ������� ������ RegExp
        RegExp.IgnoreCase = True ' ������������� ���� ������������� ��������
        RegExp.Global = True ' ������������� ���� ����������� ������
    End If
    RegExp.Pattern = "(([�-��\s]+){3,})(\(?([0-9,]+)?)" ' ������ ������ ����������� ��������� ��� ������ ��� � �����
    paralast = -1 ' �������������� ���������� ��� ������������ ���������� ������� � �������
    Set oMatches = RegExp.Execute(sl) ' ��������� ����� �� ����������� ���������
    If oMatches.Count > 0 Then ' ���� ������� ����������
        Is_Numeric = oMatches(0).SubMatches(0) ' ���������� ������ ���������� � ���������� (�� ������������ �����)
    End If
    For n = 0 To oMatches.Count - 1 ' ���� �� ���� ��������� �����������
        If Len(Trim(oMatches(n).SubMatches(0))) > 4 Then ' ���� ����� ���������� ��� ������ 4 ��������
            ' ���������, �������� �� ��� ������������ �������� ����� (����������, ���� ��������)
            If InStr(1, " " & Trim(oMatches(n).SubMatches(0)) & " ", "����������������������,�� �� ���� ������ ���", vbTextCompare) > 0 Then
                ReDim itm(0) ' ���� ������� �������� �����, ���������� ������
                Is_Fio = itm ' ���������� ������ ������
                Exit Function ' ��������� �������
            End If
            paralast = UBound(itm) + 1 ' ����������� ������ �������
            ReDim Preserve itm(paralast) ' ��������� ������ � ����������� ������
            itm(paralast).fio = Trim(oMatches(n).SubMatches(0)) ' ���������� ��� � ������
            itm(paralast).number = oMatches(n).SubMatches(3) ' ���������� ����� (���� ����) � ������
        End If
    Next
    Is_Fio = itm ' ���������� ����������� ������
End Function
Sub �������_������_��_������_���()
  Application.ScreenUpdating = False ' ������ �� ��������� �������� ����� ������� ��������
    Application.Calculation = xlCalculationManual ' ������� ��������� � ������ �����
    Application.EnableEvents = False ' ��������� �������
    If Workbooks.Count Then
        ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False ' �� ���������� ������� �����
    End If
    Application.DisplayStatusBar = False ' ��������� ��������� ������
    Application.DisplayAlerts = False ' ��������� ��������� Excel
    ' ����������� ������� ����� ���������� Vb
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    Vb.Worksheets("��������").Cells.Clear
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
    End If
    ' ��������� ���������� ���� ��� ������ �����
    Dim filePath As Variant
    filePath = OpenFileDialog1(Vb.Path) ' ��������� �����, ��� ��������� ������� �����
    ' ���������, ��� �� ������ ����
    If filePath = False Then
        MsgBox "���� �� ������!", vbExclamation
        Exit Sub
    End If
' ��������� ��������� ����
Dim Ob As Workbook
Set Ob = Workbooks.Open(filePath)

' ���������, ���������� �� ���� "������ ��� ""�-�����"""
Dim targetSheet As Worksheet
On Error Resume Next
Set targetSheet = Ob.Worksheets("������ ��� ""�-�����""")
On Error GoTo 0

If targetSheet Is Nothing Then
    MsgBox "���� ""������ ��� ""�-�����"""" �� ������!", vbExclamation
    Exit Sub
End If
' ���������� ���������� � ������� ������/������� �� ���� ������ ����� Ob
Dim ws As Worksheet
For Each ws In Ob.Worksheets
    ' ���������� ����������, ���� �� �������
    If ws.AutoFilterMode And ws.FilterMode Then
        ws.ShowAllData ' ���� ���� ������ � �� ��� ������ �����, ���������� ��
    End If
    ' ��������� � ���������� ������� ������
    If ws.Cells.EntireRow.Hidden Then
        ws.Cells.EntireRow.Hidden = False ' ���������� ��� ������
    End If
    ' ��������� � ���������� ������� �������
    If ws.Cells.EntireColumn.Hidden Then
        ws.Cells.EntireColumn.Hidden = False ' ���������� ��� �������
    End If
Next ws
' ���������� ������� ������ � ������� �� ���� ������ ������� ����� (Ob)
Dim currentSheet As Worksheet
For Each currentSheet In Ob.Sheets
    If currentSheet.Cells.EntireRow.Hidden Then
        currentSheet.Cells.EntireRow.Hidden = False ' ���������� ��� ������
    End If
    If currentSheet.Cells.EntireColumn.Hidden Then
        currentSheet.Cells.EntireColumn.Hidden = False ' ���������� ��� �������
    End If
Next currentSheet
    ' ���������� ���� "������ ��� ""�-�����"""
    Ob.Worksheets("������ ��� ""�-�����""").Activate
    ' �������� ��� ����� (��� ���� � ����������)
    fileName = Left(Ob.Name, InStrRev(Ob.Name, ".") - 1)
    ' ��������� ��������� ��� ������� (�����) �� ����� ����� ��� ���������� Den
    Den = Right(fileName, 2)
    ' ��������� 5-� � 4-� ����� ������ �� ����� ����� ��� ���������� Mesyac
    If Len(fileName) >= 5 Then
        Mesyac = Mid(fileName, Len(fileName) - 4, 2)
    Else
        Mesyac = "" ' ���� ����� ����� ����� ������ 5 ��������
    End If
    ' ����� ������ ������, ���������� "�� �� ��������������"
    Set Rngt = Cells.Find("�� �� " & Den & "." & Mesyac, , xlFormulas, xlWhole)
    If Rngt Is Nothing Then
        MsgBox "�������� ����� ������ ������������� �� ��������������, ����:" & Chr(10) & " 12.30" & Chr(10) & _
               " ���� �� ����� ������ ��� ""�-�����"" � ������ 8" & vbCrLf & " ����������� ������: �� �� " & Den & "." & Mesyac
        Exit Sub
    End If
    lRow = Rngt.Row ' ������, ��� ���� ������� �����
    lCol = Rngt.Column ' �������, ��� ���� ������� �����
    LR = Cells(Rows.Count, lCol).End(xlUp).Row
    sAddress = Rngt.Address ' ����� ������, ��� ���� ������� �����
     ''''�������� ������ �� ���� �� �������''''
     Application.Calculation = xlCalculationAutomatic '������� ��������� � ���� �����
      ' ���������, ���� �� �������� � ������
   volumeValue = Ob.Worksheets("������ ��� ""�-�����""").Cells(9, lCol).Value
    If IsEmpty(volumeValue) Or volumeValue = 0 Then
        ' ���� �������� ����������� ��� ����� 0, ������� ���������
        MsgBox "�� �������� ���� ����������� ������. ��������� ������������ ����� ""������..."" (��������� 5 �������� - ��� ""����� ����� ����"").", vbExclamation, "������"
        Exit Sub
    Else
        ' ���� �������� ����, ���������� ��� � ������ �� ����� "��������"
        Vb.Worksheets("��������").Cells(9, 15).Value = volumeValue
    End If
      Vb.Worksheets("��������").Cells(8, 15) = "�������� ������ �� �������"
      Application.Calculation = xlCalculationManual '������� ��������� � ������ �����
    ' ���������� ��� ������ �������� � ��� ����� � ������� � ������� �� ����
    Ob.Worksheets("������ ��� ""�-�����""").Range(Cells(lRow + 2, lCol), "FF" & LR).AutoFilter Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
       '  '  '  Range(Columns(15), Columns(lCol - 1)).Hidden = True ' ������ �������
    ' ����� ������ ������, ���������� "�������������"
    Set Rngt2 = Cells.Find("�������������", , xlFormulas, xlWhole)
    If Rngt2 Is Nothing Then
        MsgBox "�� ����� ������ ��� ""�-�����"" � ������ 8" & vbCrLf & "����������� ������: ������������� (��� ������ �������� � ������ ����������)"
        Exit Sub
    End If
    
    lCol2 = Rngt2.Column ' �������, ��� ���� ������� �����
    Ob.Worksheets("������ ��� ""�-�����""").Range(Columns(lCol + 1), Columns(lCol2 - 1)).Hidden = True ' ������ �������
    Ob.Worksheets("������ ��� ""�-�����""").Range(Columns(15), Columns(lCol - 1)).Hidden = True ' ������ �������
   Ob.Worksheets("������ ��� ""�-�����""").Columns("A:D").EntireColumn.Hidden = True ' ������ �������
    ' �������� ������ � ����� �����
    Ob.Worksheets("������ ��� ""�-�����""").Range(Cells(8, 5), Cells(LR, lCol2 + 1)).Copy
    ' ��������� ������������� ������ � ����� Vb �� ���� "��������" ��� ��������
        
    Vb.Worksheets("��������").Cells(8, 1).PasteSpecial Paste:=xlPasteValues
    ' ������� ����� ������
    Application.CutCopyMode = False
        Dim Prov As Worksheet
        Set Prov = Vb.Worksheets("��������")
    Vb.Worksheets("��������").Activate
    Vb.Worksheets("��������").Columns("L:M").Cut
    Vb.Worksheets("��������").Columns("A:A").Insert Shift:=xlToRight
    Vb.Worksheets("��������").Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    LRN = Cells(Rows.Count, 14).End(xlUp).Row
    Vb.Worksheets("��������").Range("A11:A" & LRN) = Den
    Vb.Worksheets("��������").Columns("C:C").ColumnWidth = 40
    Vb.Worksheets("��������").Columns("C:C").ReadingOrder = xlContext
    Vb.Worksheets("��������").Columns("C:C").WrapText = True
     Vb.Worksheets("��������").Rows("1:10").RowHeight = 11
    Vb.Worksheets("��������").Rows("8:8").RowHeight = 25
    Vb.Worksheets("��������").Rows("8:10").HorizontalAlignment = xlCenter
    Vb.Worksheets("��������").Rows("8:10").VerticalAlignment = xlCenter
    Vb.Worksheets("��������").Rows("8:8").WrapText = True
    Vb.Worksheets("��������").Rows("9:40").EntireRow.AutoFit
    Vb.Worksheets("��������").Columns("D:M").ColumnWidth = 1
    Vb.Worksheets("��������").Columns("N:R").ColumnWidth = 17
    Vb.Worksheets("��������").Columns("A:A").ColumnWidth = 3
    Vb.Worksheets("��������").Columns("B:B").ColumnWidth = 11
    



Vb.Worksheets("��������").Range("Q1:Q9") = Ob.Worksheets("���� �� �� (�)").Range("F18:F25").Value  '  ������, ���� ��������� ���������� ���
Vb.Worksheets("��������").Range("R1:R9") = Ob.Worksheets("���� �� �� (�)").Range("K18:K25").Value '  ������, ���� ��������� ���������� ���

Vb.Worksheets("��������").Range("R10") = Ob.Worksheets("���� �� �� (�)").Range("K3").Value ' ���������  ������
Vb.Worksheets("��������").Range("Q10") = "��������� �����" '
        ' �������� �������� ������� �� Alex_ST
        Range("C11:C" & LRN).Select
        If TypeName(Selection) <> "Range" Then Exit Sub
        If Intersect(Selection, ActiveSheet.UsedRange) Is Nothing Then Exit Sub
        Dim rCell As Range, i4%, ASCII%, iColor%
        Application.ScreenUpdating = False
        For Each rCell In Intersect(Selection, ActiveSheet.UsedRange)
        For i4 = 1 To Len(rCell)
        ASCII = Asc(Mid(rCell, i4, 1))
        If (ASCII >= 192 And ASCII <= 255) Or ASCII = 168 Or ASCII = 184 Then iColor = 1   ''���� �������� ��� ������
        If (ASCII >= 65 And ASCII <= 90) Or (ASCII >= 97 And ASCII <= 122) Then iColor = 3   ''���� �������� LAT �������
        rCell.Characters(Start:=i4, Length:=1).Font.ColorIndex = iColor
        Next i4
        Next rCell
        Application.ScreenUpdating = True
        Intersect(Selection, ActiveSheet.UsedRange).Select
    ' ��������� ����� Ob
    Ob.Close SaveChanges:=True
    Range("A8:N8", "A3:N3").Interior.Color = RGB(183, 222, 232)
    Range("A9:N9").Interior.Color = RGB(218, 238, 243) '.Interior.Color = RGB(200, 138, 143)
    Range("A8") = "����"
    Range("O9").FormulaLocal = "=�������������.�����(9;N4:N" & LRN & ")"
'
    Vb.Worksheets("��������").Range("R1:R10").NumberFormat = "#,##0.00"
    Vb.Worksheets("��������").Range("N9:O9").NumberFormat = "#,##0.00"
    ' �������� �� ���������� ��� ��������� � ������� � ����� � �� ����
    If Range("N9") <> Range("O9") Then
    Range("N9").Interior.Color = RGB(200, 138, 143)
    End If
    Range("A11:N" & LRN).Select

        ' ������� ������ ������� C, ���� � ������ ������ ����� ���.
        Dim itm() As Item, A()
        Vb.Worksheets("��������").Activate
        Set Prov = Vb.Worksheets("��������")
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
                ReDim A(1 To UBound(itm), 1 To 12)  ' ������� 12 ��� �����
                For i = 1 To UBound(itm)
                    A(i, 1) = itm(i).fio
                    If Val(Replace(itm(i).number, ",", ".")) <> 0 Then
                     A(i, 12) = Val(Replace(itm(i).number, ",", "."))  ' ������� 12 ��� �����
                    End If
                Next
                 Prov.Rows(LastRow1 & ":" & lastRow2).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove  '�������� ������ ���� ����� ������ � ���
                 Prov.Range("C" & LastRow1).Resize(UBound(itm), 12) = A   ' ������� 12 � , 12) = A  ������� 12 ��� �����
                 Prov.Range("N" & LastRow1).Resize(UBound(itm), 1).Interior.Color = RGB(255, 250, 235) '������� ������-������ ������ ������-������
                 Prov.Cells(n, 3).Interior.Color = RGB(255, 246, 221) '������� ������-������ ������ ������
                 Prov.Cells(n, 14).Interior.Color = RGB(255, 246, 221) '������� ������-������ ������ ������
                 Prov.Cells(n, 16) = "������ ������ ��������"
                  Prov.Cells(n, 14).Copy
                  Prov.Cells(n, 15).PasteSpecial Paste:=xlPasteValues
                  Prov.Cells(n, 15).PasteSpecial Paste:=xlPasteFormats
                  
                   'Prov.Cells(n, 15).Interior.Color = RGB(255, 249, 231) '������� ������-�����������  ������ ������ -1
            End If
        Next
    Prov.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    LastRow3 = Prov.Cells(Prov.Rows.Count, "C").End(xlUp).Row
    Prov.Range("A11:A" & LastRow3).Formula = "=ROW()-10"

    ' ������ ����������� ������ �� ������� ������� � ���� (�� �������� ������-������) � ���������� ��� .Interior.Color = 13431551
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
                i = Application.Round(Range("O" & g).Value, 4) ' ������� �� 4 ������ ����� �������
                j = Application.Round(Application.Sum(Range("O" & g + 1 & ":O" & h)), 4) ' ������� �� 4 ������ ����� �������
                If i <> j Then Range("O" & g + 1 & ":O" & h).Interior.Color = RGB(219, 179, 182) ' ������ �������
                End If
                Next
                
' ���������, ���� �� ������ ���� nx � ������� ������ ���� ��������� �����������
Dim nx As Long
' Worksheets("��������").UsedRange  '�������� ��������� � ��������� �������, �������
If LastRow3 < Prov.Rows.Count Then
    Prov.Rows(LastRow3 + 1 & ":" & Prov.Rows.Count).Delete ' ������� ��� ������ ���� ��������� �����������
End If
                
                
    Dim LastRow5 As Long  ' ��� �������� ������ ������ ��������� �������� ������
    ActiveSheet.UsedRange  '�������� ��������� � ��������� �������, �������
    LastRow5 = Cells.SpecialCells(xlLastCell).Row  '����������� ��������� ����������� ������ ��� ����������� �� �������
' ��������� ���� � ������� 2 ����� ������� B4  lastRow = Prov.Cells(Prov.Rows.Count, "C").End(xlUp).Row ������ ����������� ���� ������ ������,
'' �� � ������ ������ ������ ������� 2 �������� ������� �� ������ ������ ������� 2
'' �������� �� ������� �� 11 �� ��������� ����������� ������
   If LastRow5 >= 11 Then
    ' �������� �� ������� �� LastRow �� 11
    Dim iw As Long
    For iw = LastRow5 To 11 Step -1
        ' ���������, ������ �� ������ � ������� B
        If IsEmpty(Prov.Cells(iw, "B").Value) Or Prov.Cells(iw, "B").Value = "" Then
            ' ���� ������ ������, ���� ������ �������� ������ ����
            Dim ji As Long
            For ji = iw - 1 To 11 Step -1
                If Not IsEmpty(Prov.Cells(ji, "B").Value) And Prov.Cells(ji, "B").Value <> "" Then
                    ' �������� �������� �� ������ �������� ������ ����
                    Prov.Range("B" & iw & ":C" & iw).Value = Prov.Range("B" & ji & ":C" & ji).Value
                    Prov.Range("E" & iw & ":N" & iw).Value = Prov.Range("E" & ji & ":N" & ji).Value
                    Exit For
                End If
            Next ji
        Else
            ' ���� ������ �� ������, ���������� �� ����� ������
            LastRow5 = iw
        End If
    Next iw
End If
    
   ' 2 ��������...
     LastRow7 = Cells.SpecialCells(xlLastCell).Row  '����������� ��������� ����������� ������ ��� ����������� �� �������
 ' ������ ����� �����,
                     ''    '  [P2].FormulaLocal = "=����(O4:O" & LastRow7 & ")-����(P4:P" & LastRow7 & ")"
    Prov.[P2].FormulaLocal = "=������(����(O11:O" & LastRow7 & ")-����(P11:P" & LastRow7 & ");4)"
    Range("O9").Value = Round(Range("O9").Value, 4)
    Range("P9").Value = Round(Range("P9").Value, 4)
RoundedValue = Round(Worksheets("��������").Range("O9").Value, 5)
Worksheets("��������").Range("O9").Value = RoundedValue
RoundedValue = Round(Worksheets("��������").Range("P9").Value, 5)
Worksheets("��������").Range("P9").Value = RoundedValue
    If Prov.Range("O9").Value <> Range("P9").Value Then
    Prov.Range("O9").Interior.Color = RGB(219, 179, 182)
    End If
 ' ������� �����, �.�. ��� ������ ����� ��� ��������
  Prov.[P9].FormulaLocal = "=������(����(O11:O" & LastRow7 & ")-����(P11:P" & LastRow7 & ");4)"

     ' ����� ������ � ��������� �������
    For Row = 11 To LastRow7
    Do While Right(Cells(Row, "D").Value, 1) = " " ' ��� ������� � ����� ������ ������� 2
    Cells(Row, "D").Value = Left(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Left(Cells(Row, "D").Value, 1) = " " ' ��� ������� � ����� ������ ������� 2
    Cells(Row, "D").Value = Right(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Right(Cells(Row, "D").Value, 1) = Chr(10) ' ������ �������� �������+�������� �� ���� ������
    Cells(Row, "D").Value = Left(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Do While Left(Cells(Row, "D").Value, 1) = Chr(10) ' ������ �������� �������+�������� �� ���� ������
    Cells(Row, "D").Value = Right(Cells(Row, "D").Value, Len(Cells(Row, "D").Value) - 1)
    Loop
    Next
    ' ��� � ������� ������� ����� ������������  �������������
    Range("D11:D" & LastRow7).Replace "������� ����� ������������", "���-���-� ����� ������������", xlPart
    Range("D11:D" & LastRow7).Replace "�������������", "�������������", xlPart
    Prov.Range("A10:P" & LastRow7).AutoFilter 'Field:=lCol, Criteria1:="<>", Operator:=xlAnd, Criteria2:="<>0"
    
    
    Prov.Activate
  
'    ' ���������� �������� ����� � 1 �� 10
'    Prov.Rows("11:16").Select
'    ActiveWindow.FreezePanes = True
 
Dim ik As Long
Dim cell7 As Range
Dim Value As String
Dim foundInvalidChar As Boolean
Dim invalidChars As String
Dim char As Variant
' ������ ������������� ��������
invalidChars = "!@#$%^&*(){}[]<>?|/~+=`"
' �������� �� ���� ������� � ������� O, ������� � O4
For ik = 11 To LastRow7
Set cell7 = Prov.Cells(ik, "O")
Value = cell7.Value
' ���������, �������� �� �������� ������
If Not IsEmpty(Value) And Value <> "" Then
foundInvalidChar = False
' �������� �� ������� ����� ������ �������
If InStr(Value, ".") > 0 Then
cell7.Interior.Color = RGB(219, 179, 182) ' ��������� ������ ������-�������
MsgBox "���������� ����� ������ ������� � ������ " & cell7.Address
End If
' �������� �� ������ �������
If Trim(Value) <> Value Then
cell7.Interior.Color = RGB(219, 179, 182) ' ��������� ������ ������
MsgBox "���������� ������ ������� � ������ " & cell7.Address
End If
' �������� �� ������� ������������� ��������
For irt = 1 To Len(invalidChars)
char = Mid(invalidChars, irt, 1) ' �������� ���� ������ �� ������
If InStr(Value, char) > 0 Then
cell7.Interior.Color = RGB(219, 179, 182) ' ��������� ������ ������-�������
MsgBox "��������� ������������ ������ '" & char & "' � ������ " & cell7.Address
foundInvalidChar = True
Exit For
End If
Next irt
' ���� ������� ������������� �������, ���������� ���������� �������� ��� ���� ������
If foundInvalidChar Then GoTo NextCell
End If
NextCell:
Next ik
Application.GoTo Reference:=Prov.Range("A1"), Scroll:=True
MsgBox "������ �� ����� " & fileName & " ����������� �� ���� ""��������""." & Chr(10) _
& "��������� ��������� �� ��������� ��� � ������." & Chr(10) & "� ������ ����������� ��������� ������"
' ���� ��������� �������
Application.ScreenUpdating = True '�������� ���������� ������ ����� ������� �������
Application.Calculation = xlCalculationAutomatic '������� ������ - ����� � �������������� ������
Application.EnableEvents = True '�������� �������
If Workbooks.Count Then
ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True '���������� ������� �����
End If
Application.ScreenUpdating = True '�������� ���������� ������ ����� ������� �������
Application.Calculation = xlCalculationAutomatic '������� ������ - ����� � �������������� ������
Application.EnableEvents = True '�������� �������
If Workbooks.Count Then
ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True '���������� ������� �����
End If
Application.DisplayStatusBar = True '���������� ��������� ������
Application.DisplayAlerts = True '���������
End Sub

' ������� ��� �������� ����������� ���� ������ ����� � ��������� ��������� �����
Function OpenFileDialog1(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "�������� ���� ""������ ��� �-����� �3 ���� ��� _ ..."" �� �������� ������"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' ��������� ��������� �����, ���� ��� ������
        If InitialFolder <> "" Then
            .InitialFileName = InitialFolder & "\"
        End If
        
        If .Show = -1 Then
            OpenFileDialog1 = .SelectedItems(1) ' ���������� ��������� ����
        Else
            OpenFileDialog1 = False ' ���� ���� �� ������
        End If
    End With
End Function


