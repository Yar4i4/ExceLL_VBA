Attribute VB_Name = "Module4"
Sub �_�3_����_���()
Application.ScreenUpdating = False '������ �� ��������� �������� ����� ������� ��������
    Application.Calculation = xlCalculationManual '������� ��������� � ������ �����
    Application.EnableEvents = False '��������� �������
    If Workbooks.Count Then
        ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False '�� ���������� ������� �����
    End If
    Application.DisplayStatusBar = False '��������� ��������� ������
    Application.DisplayAlerts = False '��������� ��������� Excel

    ' ����������� ������� ����� ���������� Vb
    Set Vb = ThisWorkbook
    Vb.Worksheets("��������").Activate
    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
    End If

    ' �������� ������ �� ������� O
    Vb.Sheets("��������").Range("O:O").Copy
    ' ��������� ������ ��� ��������
    Vb.Sheets("��������").Range("O:O").PasteSpecial Paste:=xlPasteValues

    ' ��������� � ������� Q ������� ������ "������ ������ ��������" � ������� ������
    Set Prov = Vb.Sheets("��������")
    ActiveSheet.UsedRange  '�������� ��������� � ��������� �������, �������
    LastRow = Prov.Cells.SpecialCells(xlLastCell).Row '����������� ��������� ����������� ������ ��� ����������� �� �������
    For i = LastRow To 1 Step -1
        If Prov.Cells(i, "Q").Value = "������ ������ ��������" Then
            Prov.Rows(i).Delete
        End If
    Next i

    ' ������������ ���������� ����� �� 11 �� LastRow
    Dim KolProv As Long
    ActiveSheet.UsedRange  '�������� ��������� � ��������� �������, �������
    LastRowPD = Prov.Cells.SpecialCells(xlLastCell).Row '����������� ��������� ����������� ������ ��� ����������� �� �������
    If KolProv = 0 Then
        KolProv = LastRowPD - 10 ' �������� 10, ��� ��� �������� � 11 ������
    End If

    ' ��������� ���������� ���� ��� ������ �����
    Dim filePath As Variant
    filePath = OpenFileDialog(Vb.Path) ' ��������� �����, ��� ��������� ������� �����

    ' ���������, ��� �� ������ ����
    If filePath = False Then
        MsgBox "���� �� ������!", vbExclamation
        Exit Sub
    End If
    

    ' ��������� ��������� ���� � ����������� ��� ���������� P3
    Set P3 = Workbooks.Open(filePath)
    
    
    
    ' ������ ���� ������ ��� ��������
    sheetNames = Array("������� �� ���", "������� �� ��������", "���� �� �� ������� ����")
    missingSheets = "" ' ���������� ��� �������� ������������� ������
    
    ' ��������� ������������� ������� �����
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = P3.Sheets(sheetName)
        On Error GoTo 0
        
        If ws Is Nothing Then
            ' ���� ���� �� ������, ��������� ��� ��� � ������ �������������
            missingSheets = missingSheets & sheetName & vbCrLf
        Else
            ' ���� ���� ����������, ������� ������ �� ����
            Set ws = Nothing
        End If
    Next sheetName
    
    ' ������� ��������� ��������
    If missingSheets = "" Then
        ' MsgBox "��� ����������� ����� ����������.", vbInformation
    Else
        MsgBox "����������� ��������� �����:" & vbCrLf & missingSheets, vbExclamation
        P3.Close SaveChanges:=False ' ��������� ���� ��� ���������� ���������, ���� ����� �����������
        Exit Sub
    End If
    
    ' ���������, ���� �� ������� ������ � ������� �� ������ "������� �� ���", "������� �� ��������", "���� �� �� ������� ����"
    sheetNames = Array("������� �� ���", "������� �� ��������", "���� �� �� ������� ����")
    
    ' �������� �� ������� �����
    For i2 = LBound(sheetNames) To UBound(sheetNames)
        Set ws = P3.Sheets(sheetNames(i2))
        
        ' ��������� ������� ������� �����
        hasHiddenRows = False
        On Error Resume Next
        hasHiddenRows = ws.Cells.SpecialCells(xlCellTypeVisible).Rows.Count < ws.Rows.Count
        On Error GoTo 0
        
        ' ��������� ������� ������� ��������
        hasHiddenColumns = False
        On Error Resume Next
        hasHiddenColumns = ws.Cells.SpecialCells(xlCellTypeVisible).Columns.Count < ws.Columns.Count
        On Error GoTo 0
        
        ' ���� ���� ������� ������, ���������� ��
        If hasHiddenRows Then
            ws.Rows.Hidden = False
            ' MsgBox "�� ����� '" & ws.Name & "' ���� ������� ������. ��� ��������."
        End If
        
        ' ���� ���� ������� �������, ���������� ��
        If hasHiddenColumns Then
            ws.Columns.Hidden = False
            ' MsgBox "�� ����� '" & ws.Name & "' ���� ������� �������. ��� ��������."
        End If
    Next i2
     
     
     
    ' �� ����� "���� �� �� ������� ����" � ������� B ����� �������� �� ����� Vb ���� "��������" ������ B4, ��� s22 ��� �������� �� ������ B4
    Set wsFact = P3.Sheets("���� �� �� ������� ����")
    s22 = Prov.Range("B11").Value
    ' ���� �������� ������ � ������� B
'    Set FoundCell = wsFact.Columns("B").Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
    ' ���� �������� ������ � ������� B, ������� � ������ B4 � ����
Set foundCell = wsFact.Range("B4:B" & wsFact.Cells(wsFact.Rows.Count, "B").End(xlUp).Row).Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
' ���� �������� �������, ������� ��������� � ���������� ������������
If Not foundCell Is Nothing Then
    response = MsgBox("�������� ���� " & s22 & " � �3 ���� ��� ������������." & Chr(10) & "������� ������ �� ��������� ��� �� ""�3 ���� ���""" & Chr(10) & "���������� ���������� �������?", vbInformation + vbYesNo, "������")
        ' ���� ������������ ����� "���������" (������ "���"), ��������� ���������� �������
    If response = vbNo Then
        Exit Sub
    End If
End If
' ���� �������� �� �������, �������� ������ �� ����� Vb � ��������� � ����� P3
  ActiveSheet.UsedRange  '�������� ��������� � ��������� �������, �������
    lastRow2 = Prov.Cells.SpecialCells(xlLastCell).Row '����������� ��������� ����������� ������ ��� ����������� �� �������
Set copyRange = Prov.Range("A11:O" & lastRow2)
' ������� ������ ��� ������� � ����� P3
insertRow = wsFact.Cells(wsFact.Rows.Count, "D").End(xlUp).Row + 1

                             ' ������ ����� ������
                                                                    wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).NumberFormat = "General"
                                                                With wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).Font
                                                                    .Name = "Calibri"
                                                                    .Size = 11
                                                                    .Bold = False
                                                                    .Italic = False
                                                                    .Color = 0
                                                                End With
                                                                wsFact.Range("A" & insertRow & ":BD" & insertRow).Interior.Color = RGB(255, 255, 255) ' ����� ����
                                                                wsFact.Range("A" & insertRow & ":BD" & insertRow).Interior.Pattern = xlNone ' ��������� �������
                                                                With wsFact.Range("A" & (insertRow) & ":BD" & (insertRow)).Borders
                                                                    .LineStyle = 1
                                                                    .Weight = 1
                                                                    .Color = 6773025
                                                                End With
                                                                With wsFact.Range("A" & insertRow & ":BD" & insertRow)
                                                                .HorizontalAlignment = xlCenter ' �������������� ������������ �� ������
                                                                .VerticalAlignment = xlCenter   ' ������������ ������������ �� ������
                                                                .WrapText = False                ' ������� ������ ����
                                                                .Orientation = 0                ' ���������� ������ (0 ��������)
                                                                .IndentLevel = 0                ' ������� �������
                                                                End With
                                        wsFact.Range("D" & insertRow).HorizontalAlignment = xlLeft
                                        wsFact.Range("H" & insertRow & ":M" & insertRow).HorizontalAlignment = xlLeft
                          wsFact.Range("P" & insertRow & ":BB" & insertRow).NumberFormat = "_-* #,##0.00 _?_-;-* #,##0.00 _?_-;_-* ""-""?? _?_-;_-@_-"
                        wsFact.Range("AO" & insertRow & ":AR" & insertRow).NumberFormat = "_-* #,##0.0000000000000 _?_-;-* #,##0.0000000000000 _?_-;_-* ""-""?? _?_-;_-@_-"
                        wsFact.Range("T" & insertRow).NumberFormat = "_-* #,##0.000000000000 _?_-;-* #,##0.000000000000 _?_-;_-* ""-""?? _?_-;_-@_-"
    ' �������
 wsFact.Range("P" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],�����!C[-14]:C[-6],6,0),0)" '����. �������� �� ������ ����� �� ��. ���.
wsFact.Range("Q" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],�����!C[-15]:C[-7],7,0),0)" '����. �������� �� ������������ ������������ (������������ ����� � ���������� - ���), ���.
wsFact.Range("R" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],�����!C[-16]:C[-8],8,0),0)" '����. �������� �� ���-�� �� ��., ���.
wsFact.Range("S" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],�����!C[-17]:C[-9],9,0),0)" '����. �������� �� ������ �������, ���.
wsFact.Range("T" & insertRow).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],'%��_k ���������'!C[-11]:C[-10],2,0),0)" '%��
wsFact.Range("U" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" '��� ��
wsFact.Range("V" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" '��� ��
wsFact.Range("W" & insertRow).FormulaR1C1 = "=ROUND(RC15*RC[-5],2)" '��� ��
wsFact.Range("X" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-3],RC[-2],RC[-1]),2)" '��
wsFact.Range("Y" & insertRow).FormulaR1C1 = "=ROUND(RC[-1]*RC[-5],2)" '��
wsFact.Range("Z" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%��_k ���������'!R8C3-RC[-5],2)" '��� k
wsFact.Range("AA" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%��_k ���������'!R8C3-RC[-5],2)" '��� k
wsFact.Range("AB" & insertRow).FormulaR1C1 = "=ROUND(RC[-5]*'%��_k ���������'!R8C3-RC[-5],2)" '��� k
wsFact.Range("AC" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-3],RC[-2],RC[-1]),2)" '����� k
wsFact.Range("AD" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" '��� �����
wsFact.Range("AE" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" '��� �����
wsFact.Range("AF" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-9],RC[-4]),2)" '��� �����
wsFact.Range("AG" & insertRow).FormulaR1C1 = "=RC[-8]" '��
wsFact.Range("AH" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-4],RC[-3],RC[-2],RC[-1]),2)" '�����
wsFact.Range("AI" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-5],RC[-4],RC[-2]),2)" '���
wsFact.Range("AJ" & insertRow).FormulaR1C1 = "=RC[-4]" '���
wsFact.Range("AK" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '����� � ��������-��������� �� 2025 �.
wsFact.Range("AL" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '��� � ��������-��������� �� 2025 �.
wsFact.Range("AM" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '��� � ��������-��������� �� 2025 �.
wsFact.Range("AN" & insertRow).FormulaR1C1 = "" '�����
wsFact.Range("AO" & insertRow).FormulaR1C1 = "3.5869670184836" '���� �� ���
wsFact.Range("AP" & insertRow).FormulaR1C1 = "9.14237560208042" '���� �� ���
wsFact.Range("AQ" & insertRow).FormulaR1C1 = "1.96204565821346" '���� �� ���
wsFact.Range("AR" & insertRow).FormulaR1C1 = "3.59699312139276" '� �� ���
wsFact.Range("AS" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" '��� �����
wsFact.Range("AT" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" '��� �����
wsFact.Range("AU" & insertRow).FormulaR1C1 = "=ROUND(RC[-24]*RC[-4],2)" '��� �����
wsFact.Range("AV" & insertRow).FormulaR1C1 = "=ROUND(RC[-23]*RC[-4],2)" '��
wsFact.Range("AW" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-4],RC[-3],RC[-2],RC[-1]),2)" '�����
wsFact.Range("AX" & insertRow).FormulaR1C1 = "=ROUND(SUM(RC[-5],RC[-4],RC[-2]),2)" '���
wsFact.Range("AY" & insertRow).FormulaR1C1 = "=RC[-4]" '���
wsFact.Range("AZ" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '����� � ��������-��������� �� 2025 �.
wsFact.Range("BA" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '��� � ��������-��������� �� 2025 �.
wsFact.Range("BB" & insertRow).FormulaR1C1 = "=ROUND(RC[-3]*1.091*1.078,2)" '��� � ��������-��������� �� 2025 �.

        ' �������� �������� "P & (insertRow - 1) : BD & (insertRow - 1)" � ��������� ��� KolProv ���
If KolProv > 0 Then
    ' ���������� �������� ��� �����������
    Dim sourceRange As Range
    Set sourceRange = wsFact.Range("A" & (insertRow) & ":BD" & (insertRow))
        ' �������� ��������
    sourceRange.Copy
           ' ��������� ������������� �������� KolProv ���
    wsFact.Range("A" & (insertRow) & ":BD" & (insertRow + KolProv - 1)).PasteSpecial Paste:=xlPasteAll
End If
' ����� ��� Range("A" & insertRow & ":O" & insertRow + copyRange.Rows.Count - 1)
  wsFact.Range("A" & (insertRow) & ":O" & (insertRow + KolProv - 1)).ClearContents
    
' ��������� ������ ��� �������� ��� ��������������
wsFact.Range("A" & insertRow & ":O" & insertRow + copyRange.Rows.Count - 1).Value = copyRange.Value
                                                       
' ������� ����� ������
Application.CutCopyMode = False
wsFact.Range("O2").FormulaLocal = "=�������������.�����(9;O4:O" & CStr(insertRow + KolProv - 1) & ")"
       
' ���� �� wsFact. � ������� E ���� �������� "2", �� � ��������������� ������ ������� Z � AA �������� "0", �.�. �� 2 ������ ����. 0 (�� 1).
Dim checkRangeFact As Range
Set checkRangeFact = wsFact.Range("E" & insertRow & ":E" & (insertRow + KolProv - 1))
Dim irr As Long
For irr = 1 To checkRangeFact.Rows.Count
    If checkRangeFact.Cells(irr, 1).Value = "2" Then
        wsFact.Cells(insertRow + irr - 1, "Z").Value = "0"
        wsFact.Cells(insertRow + irr - 1, "AA").Value = "0"
    End If
Next irr






    ' ��������� ������� � ������� ������� ������������ �������
    Set formulaRange = wsFact.Range("A" & insertRow & ":A" & insertRow + copyRange.Rows.Count - 1)
    formulaRange.Formula = "=ROW()-3"
    
    ' ������� ����� ������ � ������� Q �� ����� "������� �� ���", ��� ��������� �������� s22
    Set wsSvodSMU = P3.Sheets("������� �� ���")
    Set FoundCell2 = wsSvodSMU.Columns("Q").Find(What:=s22, LookIn:=xlValues, LookAt:=xlWhole)
        
    ' ���� �������� �������, �������� ������ �� ������� S �� ����� "��������" � ��������� �� ���� "������� �� ���"
    If Not FoundCell2 Is Nothing Then
        ' ���������� �������� ��� ����������� �� ������� S
        Set copyRangeS = Prov.Range("S1:S8") '  ������, ���� ��������� ���������� ���   S8
        ' ���������� �������� ��� ������� �� ���� "������� �� ���" ����������� �� ��������� �� ������� �����
        Set insertRangeS = wsSvodSMU.Range("W" & FoundCell2.Row + 3 & ":W" & FoundCell2.Row + 10)  '  ������, ���� ��������� ���������� ���   +10
        ' �������� � ��������� ������
        copyRangeS.Copy
        insertRangeS.PasteSpecial xlPasteValues
        ' ������� ����� ������
        Application.CutCopyMode = False
         ' �������� �� �������� ������� R ���� "��������� �����" + 1 ������� ������
'         � �� ���� ��������� � ���� wsSvodSMU � ������� AC ���� "���� (�) ��" + 1 ������� ������, �.�. �������� � ���� ���������
    Else
        MsgBox "�������� ���� " & s22 & " �� ������ �� ����� '������� �� ���'.", vbExclamation
    End If
          
        
        ' ����� ������ ������ � 1 ������� � �������� ��� ���� 1�� ���. ���������� ��� �������� ������ ������ � ��������
Dim filledRow As Long
filledRow = 0
' ���������� ��� �������� ������� �������
Dim hasFilledCell As Boolean
hasFilledCell = False
' ���������� ��� �������� ����� �������
Dim fillColor As Long
fillColor = RGB(218, 238, 243) '�����
' ������� ��������� ����������� ������ � ������� A
P3.Sheets("������� �� ��������").Activate
Set wsPro = P3.Sheets("������� �� ��������")
lastRowPro = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row
' �������� �� ������� ������� A � 6 ������ �� ��������� �����������
For i = 6 To lastRowPro
    ' ��������� ������� ������
    If wsPro.Cells(i, "A").Interior.Color = fillColor Then
        filledRow = i
        hasFilledCell = True
        Exit For
    End If
Next i
' ���� ������� �� �������, ������� ��������� � ��������� ����������
If Not hasFilledCell Then
    MsgBox "�� ����� '������� �� ��������' ����������� ������� ����� RGB(218, 238, 243) � ������� �. ������� ������� ����������.", vbExclamation
    Exit Sub
End If
'' � ����������� �� ������ ������ � �������� ��������� ��������������� ��������
Select Case filledRow
    Case 6
        ' ��������� ��� ������ ������ ����� 5 ������
        wsPro.Rows("6:7").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' ������� ������� ����������� �����
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case 7
        ' ��������� ���� ������ ������ ����� 5 ������
        wsPro.Rows("6:6").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ' ������� ������� ����������� ������
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case 8
        ' ������� ������� �����
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
    Case Is >= 9
        ' ������� ������ � 6 �� (filledRow - 3)
        wsPro.Rows("6:" & (filledRow - 3)).Delete Shift:=xlUp
        ' ������� ������� �����
        wsPro.Rows("6:7").Interior.ColorIndex = xlNone
End Select

'' ������� ���������� ����� A6:B6
wsPro.Range("A6:B7").ClearContents
'
' ��������� ������� � ������ D6
wsPro.Range("D6").FormulaLocal = "=����������(����������(D$9:D$" & lastRowPro + 999 & ";$A$9:$A$" & lastRowPro + 999 & ";$A6;$B$9:$B$" & lastRowPro + 999 & ";$B6);0)"
 wsPro.[D6].Copy
 wsPro.Range("D6:I6").PasteSpecial Paste:=xlPasteFormulas
' ��������� ������� � ������ D6
wsPro.Range("K6").FormulaLocal = "=����������(����������(K$9:K$" & lastRowPro + 999 & ";$A$9:$A$" & lastRowPro + 999 & ";$A6;$B$9:$B$" & lastRowPro + 999 & ";$B6);0)"
 wsPro.[K6].Copy
 wsPro.Range("K6:P6").PasteSpecial Paste:=xlPasteFormulas
 
' ���������� �������� D4:I4 � C ����� ����
    With wsPro.Range("D4:I4").Interior
        .Color = RGB(218, 238, 243) ' � c���� ����
        .Pattern = xlSolid ' �������� �������
    End With
 
 
 
' �� ����� ���� �� �� ������� ���� ������� � ������ C4 �� ������ ����������� ������ � ������� D ���������� ������ �� ������� �������
' � ����������� ��� ������ � ����������. ����� �� ����� ���� �� �� ������� ���� ������ ������ �� �����, � �������� ���� � ����������, �� ������� ��������� ������.
' ����� � �������, ������������ ���������� ������� ���������, � ������ ���� �������� � �������. ���������� ���������� �����, ���������� ����� �������� ���������� �
'  ����������� ���������� ������ 6  �� ���� ������� �� �������� �  �������� � ����� �� ���������� ��� �� ���� ������� �� �������� ����� ������ 6, ������� ��� �� �� ���������� ����� �������� ����������.
'����� ����� �������� ������ �� �������, ��� ������� ����� ������� ������� ����� A6 �� ����� ������� �� ��������
' ���������� �������� ������ �� ����� "���� �� �� ������� ����"
Set DataRange = wsFact.Range("B4:D" & wsFact.Cells(wsFact.Rows.Count, "C").End(xlUp).Row)
dataArray = DataRange.Value ' ��������� ������ � ������

' ������� ��������� ��� �������� ���������� ������
Set uniqueData = New Collection

' ������� ���������, �������� ��� ������� (C � D)
On Error Resume Next
For i = 1 To UBound(dataArray, 1)
    ' ������� ���������� ���� �� �������� C � D (������� 2 � 3 � �������, ��� ��� B4:D ���������� � B)
    Key = dataArray(i, 2) & "|" & dataArray(i, 3) ' ���������� ���� �� �������� C � D
    uniqueData.Add Key, Key ' ��������� ���� � ��������� (��������� ����� ��������������)
Next i
On Error GoTo 0

' ������������ ���������� ���������� �����
insertRowsCount = uniqueData.Count

' �������� ������ 6 �� ����� "������� �� ��������"
wsPro.Rows(6).Copy
 wsPro.Rows("6:" & insertRowsCount + 5 - 1).Insert Shift:=xlDown ' ��������� ������� ���, ������� ���������� �����

' ������� ����� ������
Application.CutCopyMode = False

' ��������� ���������� ������ �� ���� "������� �� ��������", ������� � A6
For i = 1 To uniqueData.Count
    ' ��������� ���� ������� �� ��� �������
    Dim splitData() As String
    splitData = Split(uniqueData(i), "|")

    ' ��������� ������ � ������ A � B
    wsPro.Cells(6 + i - 1, "A").Value = splitData(0) ' ������� C (������ ����� �����)
    wsPro.Cells(6 + i - 1, "B").Value = splitData(1) ' ������� D (������ ����� �����)
Next i
  '��������� ������ �� A6 �� (B insertRowsCount + 5), ��� insertRowsCount +5 ��� ������ ������ ������ ������� ��� ���������� �� ������� ������� �� � �� �.
   ' ����� ��������� ������ �� ������� ������� �� � �� �.+
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

                                                                ' �������� ������
                                                                wsPro.Range("A6:B" & insertRowsCount + 5).Select
    ' doober �������� ������ ����� ������� �������� ����������� ���
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
    ' ��������� ���� � ��������
insertRowsCount = wsPro.Cells(wsPro.Rows.Count, "B").End(xlUp).Row - 5 ' ���������� ���������� ����� (�������� ���������)
' �������� ������ �� ��������� � ������
dataArr = wsPro.Range("B6:B" & insertRowsCount + 5).Value ' �������� ��� �������� (������� � B6)
' ���� �� ������� � �������
For yzy = 1 To UBound(dataArr, 1) ' ���������� ������ �������
    ' ���������, ������ �� ������� ������ ������ 13431551
    If wsPro.Cells(yzy + 5, 2).Interior.Color = 13431551 Then ' ��������� �������� �� B6
        ' ���������, ���� �� ������� ������ ������ ��� �����
        Dim hasNeighbor As Boolean
        hasNeighbor = False
        ' �������� ������ (���� ��� �� ������ ������)
        If yzy > 1 Then
            If wsPro.Cells(yzy + 4, 2).Interior.Color = 13431551 Then ' ������ ������
                hasNeighbor = True
            End If
        End If
        ' �������� ����� (���� ��� �� ��������� ������)
        If yzy < UBound(dataArr, 1) Then
            If wsPro.Cells(yzy + 6, 2).Interior.Color = 13431551 Then ' ������ �����
                hasNeighbor = True
            End If
        End If
        ' ���� ��� �������� ������� �����, ������������� ������� � ���������� ���
        If Not hasNeighbor Then
            wsPro.Cells(yzy + 5, 2).Interior.ColorIndex = xlNone ' ���������� ���
        End If
    End If
Next yzy
   
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' ����� ������ ������ � ������� A, ���������� ������� fillColor
Dim filledRow2 As Long
filledRow2 = 0
' ���������� ��� �������� ������� �������
Dim hasFilledCell2 As Boolean
hasFilledCell2 = False
' ���������� ��� �������� ����� �������
Dim fillColor2 As Long
fillColor2 = RGB(218, 238, 243) '�����

' ������� ��������� ����������� ������ � ������� A
P3.Sheets("������� �� ��������").Activate
'Set wsPro = P3.Sheets("������� �� ��������")
LastRowPro2 = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row

' �������� �� ������� ������� A � 6 ������ �� ��������� �����������
For i = 6 To LastRowPro2
    ' ��������� ������� ������
    If wsPro.Cells(i, "A").Interior.Color = fillColor Then
        filledRow2 = i
        hasFilledCell2 = True
    End If
Next i

' ���� ������� �� �������, ������� ��������� � ��������� ����������
If Not hasFilledCell2 Then
    MsgBox "�� ����� '������� �� ��������' ����������� ������� ����� RGB(218, 238, 243) � ������� �. ������� ������� ����������.", vbExclamation
    Exit Sub
End If
    
   ' �������� ������ � filledRow2 �� filledRow2 + 3 �� ����� "������� �� ��������"
wsPro.Rows(filledRow2 & ":" & filledRow2 + 1).Copy
wsPro.Rows(LastRowPro2 + 2).Insert Shift:=xlDown
Application.CutCopyMode = False

 ' ������� �:�  �  LastRowPro2 + 3 �� LastRowPro2 + 3 �� ����� "������� �� ��������"
 wsPro.Range("A" & LastRowPro2 + 3 & ":C" & LastRowPro2 + 3 & "").ClearContents
 
'  �� ����� ���� �� �� ������� ���� ������� � ������ B4 �� ������ ����������� ������ � ������� D ���������� ������ �� ������� �������
' � ����������� ��� ������ � ����������. ����� �� ����� ���� �� �� ������� ���� ������ ������ �� �����, � �������� ���� � ����������, �� ������� ��������� ������.
'����� � �������, ������������ ���������� �������� ��� ���������� ������ �� ������,
'������� �������� � ������ ������� ������� (������� B) ��������, ������� ����� � ���� �� ��������� ���������� s22.
'����� ����� ������� ���������, � ������ ���� �������� � � D � �������. ���������� ���������� �����, ���������� ����� �������� ���������� �
'  ����������� ���������� ������� �� A �� P ��������  �� ����� ������� �� �������� � ������  LastRowPro + 2 �  �������� ��� ����� �� ���������� ��� �� ���� ������� �� ��������
'����� ������  LastRowPro + 2, ������� ��� �� �� ���������� ����� �������� ����������.
'����� ����� �������� ������ �� �������, ��� ������� ����� ������� ������� ����� LastRowPro + 3 �� ����� ������� �� ��������
' �� ������� dataArray �������� ������ �� ������, ������� ����������� �������� �������� s22
' �� ����� "���� �� �� ������� ����" ������� � ������ B4 �� ������ ����������� ������ � ������� D ���������� ������ �� ������� �������
Set DataRange = wsFact.Range("B4:D" & wsFact.Cells(wsFact.Rows.Count, "D").End(xlUp).Row)
dataArray = DataRange.Value ' ��������� ������ � ������

' ������� ����� ������ ��� �������� �����, ��� � ������ ������� (������� B) ���������� �������� s22
Dim filteredData() As Variant
Dim filteredRowCount As Long
filteredRowCount = 0
'
' �������� �� ������� � �������� ������, ��� � ������ ������� (������� B) ���������� �������� s22
For i = 1 To UBound(dataArray, 1)
    If dataArray(i, 1) = s22 Then
        filteredRowCount = filteredRowCount + 1
        ReDim Preserve filteredData(1 To 3, 1 To filteredRowCount)
        filteredData(1, filteredRowCount) = dataArray(i, 1) ' ������� B
        filteredData(2, filteredRowCount) = dataArray(i, 2) ' ������� C
        filteredData(3, filteredRowCount) = dataArray(i, 3) ' ������� D
    End If
Next i

' ������� ��������� ��� �������� ���������� ������
Set uniqueData = New Collection

' ������� ���������, �������� ��� ������� (C � D)
On Error Resume Next
For i = 1 To filteredRowCount
    ' ������� ���������� ���� �� �������� C � D (������� 2 � 3 � �������)
    Key = filteredData(2, i) & "|" & filteredData(3, i) ' ���������� ���� �� �������� C � D
    uniqueData.Add Key, Key ' ��������� ���� � ��������� (��������� ����� ��������������)
Next i
On Error GoTo 0

' ������������ ���������� ���������� �����
insertRowsCount = uniqueData.Count

'' �������� ������ � ��������� � insertRowsCount ���
    wsPro.Rows(LastRowPro2 + 3).Copy
'    ��������� ������������� ������ insertRowsCount ���
wsPro.Rows(LastRowPro2 + 3 & ":" & LastRowPro2 + 3 + insertRowsCount - 2).Insert Shift:=xlDown


' ��������� ���������� ������ �� ���� "������� �� ��������", ������� � LastRowPro2 + 3
For i = 1 To uniqueData.Count
    ' ��������� ���� ������� �� ��� �������
    Dim splitData5() As String
    splitData5 = Split(uniqueData(i), "|")

    ' ��������� ������ � ������ A � B
    wsPro.Cells(LastRowPro2 + 3 + i - 1, "A").Value = splitData5(0) ' ������� C (������ ����� �����)
    wsPro.Cells(LastRowPro2 + 3 + i - 1, "B").Value = splitData5(1) ' ������� D (������ ����� �����)
Next i
  ' ��������� ������ � ������ �
 wsPro.Range("C" & LastRowPro2 + 3 & ":C" & LastRowPro2 + 3 + insertRowsCount - 1) = s22
 ' ���������  ������ wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1) �� � �� � ������� �� 2 �������, ����� �� 1 �������
' ��������� ������ �� A �� � ������� �� ������� ������� (B), ����� �� ������� ������� (A)
With wsPro.Sort
    .SortFields.Clear
    ' ���������� �� ������� ������� (B)
    .SortFields.Add2 Key:=wsPro.Range("B" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     ' ������������� �������� ��� ����������
    .SetRange wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1)
    .Header = xlNo ' ���������, ��� ���������� ���
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
        .SortFields.Clear
    ' ���������� �� ������� ������� (A)
    .SortFields.Add2 Key:=wsPro.Range("A" & LastRowPro2 + 3 & ":A" & LastRowPro2 + 3 + insertRowsCount - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' ������������� �������� ��� ����������
    .SetRange wsPro.Range("A" & LastRowPro2 + 3 & ":B" & LastRowPro2 + 3 + insertRowsCount - 1)
    .Header = xlNo ' ���������, ��� ���������� ���
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

wsPro.UsedRange  '�������� ��������� � ��������� �������, �������
Dim LastRow5 As Long
LastRow5 = Cells(Rows.Count, "A").End(xlUp).Row
' ������������� ������� ������
'    ������� ������: ��������� - ��������� ������, ����������� - ������� ������� h
    LastRow5 = wsPro.UsedRange.Row + wsPro.UsedRange.Rows.Count - 1
    wsPro.PageSetup.PrintArea = wsPro.Range(Cells(1, 1), Cells(LastRow5, 16)).Address

    ' ������� ��� ������ ���� LastRow5
If LastRow5 < Rows.Count Then
     wsPro.Rows(LastRow5 + 1 & ":" & Rows.Count).Delete
End If


'������ ����. � ������ wsPro.Range("A" & LastRowPro2 + 2) ��������� �����. ����� ��� ����� ���������, ��������� � ���� 1 � ��������� ������� � ��� �� ������
' ���� � ������ wsPro.Range("A" & LastRowPro + 2) ��������� ����, ����� ��������� ��, �������� ����������, �����
'������ ������ ��� ����� ����� �� �������� �� ���������� s22, ��� s22 ��� ����� ��� �����
' ���������� ������� ������
Dim targetCell As Range
Set targetCell = wsPro.Range("A" & LastRowPro2 + 2)

' ���������� ������� �������� (����) � ����������
Dim originalDate As String
originalDate = CStr(targetCell.Value)
' ���������, ��� ����� ������ ��������� �������� ������ ��� �������
If Len(originalDate) >= 2 Then
    ' ����������� s22 � ����� ��� ��������
    Dim s22Value As Long
    s22Value = CLng(s22)
    ' ���������, ��������� �� �������� � ��������� �� 1 �� 9
    If s22Value >= 1 And s22Value <= 9 Then
        ' ����������� s22 ��� ����� � ������� ���� (��������, "01", "02", ...)
        Dim formattedS22 As String
        formattedS22 = Format(s22Value, "00")
        ' �������� ������ ��� ����� �� ����������������� ��������
        Dim modifiedDate As String
        modifiedDate = formattedS22 & Mid(originalDate, 3)
        ' ��������� �������� ������
        targetCell.Value = modifiedDate
    Else
        ' ���� �������� �� � ��������� �� 1 �� 9, ������ �������� ������ ��� �������
        Dim modifiedDate4 As String
        modifiedDate4 = s22 & Mid(originalDate, 3)
        ' ��������� �������� ������
        targetCell.Value = modifiedDate4
    End If
Else
    MsgBox "���� ������� �������� ��� ������ ������ ���� ��������!", vbExclamation
End If
  
    ' ������� � ����� �������� ��� ������������� �����
     wsPro.Range("D" & LastRowPro2 + 2).FormulaLocal = "=�������������.�����(9;D" & LastRowPro2 + 3 & ":D" & LastRowPro2 + 3 + insertRowsCount - 1 & ")"
     wsPro.Range("D" & LastRowPro2 + 2).Copy
     wsPro.Range("D" & LastRowPro2 + 2 & ":I" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas
     wsPro.Range("K" & LastRowPro2 + 2 & ":P" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas

    
 ' ����� ������������� ������ � 1 ������� � �������� ��� ���� ���������� ���
filledRow3 = 0
' ���������� ��� �������� ������� �������
hasFilledCell3 = False
Dim fillColor3 As Long
fillColor3 = RGB(218, 238, 243) '�����
' �������� �� ������� ������� A � 6 ������ �� ��������� �����������
For i = LastRowPro2 To 6 Step -1
    ' ��������� ������� ������
    If wsPro.Cells(i, "A").Interior.Color = fillColor3 Then
        filledRow3 = i
        hasFilledCell3 = True
        Exit For
    End If
Next i


' �����  ' ����� ������ ������ � 1 ������� � �������� ��� ���� 1�� ���. ���������� ��� �������� ������ ������ � ��������
filledRow5 = 0
' ���������� ��� �������� ������� �������
hasFilledCell5 = False
Dim fillColor5 As Long
fillColor5 = RGB(218, 238, 243) '�����
' �������� �� ������� ������� A � 6 ������ �� ��������� �����������
For i = 6 To LastRow5
    ' ��������� ������� ������
    If wsPro.Cells(i, "A").Interior.Color = fillColor5 Then
        filledRow5 = i
        hasFilledCell5 = True
        Exit For
    End If
Next i







        
     wsPro.Range("D4").FormulaLocal = "=�������������.�����(9;D6:D" & filledRow5 - 2 & ")"
     wsPro.Range("D4").Copy
     wsPro.Range("D4:I4").PasteSpecial Paste:=xlPasteFormulas
     wsPro.Range("K4:P4").PasteSpecial Paste:=xlPasteFormulas
'    ' ������� �������� �� ������ ������� U ����� ���� �����
wsPro.Range("U" & LastRowPro2 + 2 & ":Z" & LastRowPro2 + 3 + insertRowsCount - 1).ClearContents

wsPro.Range("U" & LastRowPro2 + 2).Formula = "=D" & (LastRowPro2 + 2)
wsPro.Range("U" & LastRowPro2 + 2).Copy
wsPro.Range("U" & LastRowPro2 + 2 & ":Z" & LastRowPro2 + 2).PasteSpecial Paste:=xlPasteFormulas

wsPro.Range("U" & LastRowPro2 + 3).Formula = "='������� �� ���'!C" & (FoundCell2.Row + 11)  '  ������, ���� ��������� ���������� ��� +11
wsPro.Range("U" & LastRowPro2 + 3).Copy
wsPro.Range("U" & LastRowPro2 + 3 & ":Z" & LastRowPro2 + 3).PasteSpecial Paste:=xlPasteFormulas

wsPro.Range("U" & LastRowPro2 + 4).Formula = "=U" & (LastRowPro2 + 3) & "=U" & (LastRowPro2 + 2) ' ������ ����
wsPro.Range("U" & LastRowPro2 + 4).Copy ' ������ ����
wsPro.Range("U" & LastRowPro2 + 4 & ":Z" & LastRowPro2 + 4).PasteSpecial Paste:=xlPasteFormulas ' ������ ����

' ����� ����� ��� �������� ������������� ����� �� ���������� ����, ������� ��������� � �������� U:Z
'��� ����� �� ���� ����� ����� ������ ��������� ��������   fillColor4 = RGB(218, 238, 243)
Dim filledRow4 As Long ' ��������� ���������� ����
Dim fillColor4 As Long
fillColor4 = RGB(218, 238, 243) ' ���� ������� '�����
Dim foundCount As Long
foundCount = 0 ' ������� ��������� ����� � ��������
' ������� ��������� ����������� ������ � ������� A
LastRowPro7 = wsPro.Cells(wsPro.Rows.Count, "A").End(xlUp).Row
' �������� �� ������� ������� A ����� �����
For i = LastRowPro7 To 1 Step -1
    ' ��������� ������� ������
    If wsPro.Cells(i, "A").Interior.Color = fillColor4 Then
        foundCount = foundCount + 1 ' ����������� ������� ��������� �����
        If foundCount = 2 Then ' ���� ��� ������ ��������� ��������
            filledRow4 = i ' ����������� ����� ������ ����������
            Exit For ' ������� �� �����
        End If
    End If
Next i
' ���������, ������� �� ������ ��������
If foundCount < 2 Then
    MsgBox "������ �������� � �������� RGB(218, 238, 243) � ������� A �� �������!", vbExclamation
'    Exit Sub
End If

                                                      ' ����� �����, ����� ������� ���������� ������� ������ � ������ ����������� �� ����������� ���� �������  Application.Calculation = xlCalculationAutomatic
                                                       Application.Calculation = xlCalculationAutomatic
                                                    wsPro.Range("U" & LastRowPro2 + 5).Formula = "=U" & (LastRowPro2 + 2) & "+U" & (filledRow4 + 3) '
                                                    wsPro.Range("U" & LastRowPro2 + 5).Copy '
                                                    wsPro.Range("U" & LastRowPro2 + 5 & ":Z" & LastRowPro2 + 5).PasteSpecial Paste:=xlPasteFormulas '
                                                    
                                                    wsPro.Range("U" & LastRowPro2 + 6 & ":Z" & LastRowPro2 + 6).Value = wsSvodSMU.Range("C13:H13").Value   '  ������, ���� ��������� ���������� ��� C13:H13
                                                    wsPro.Range("U" & LastRowPro2 + 7).Formula = "=U" & (LastRowPro2 + 6) & "=U" & (LastRowPro2 + 5) ' ������ ����
                                                    wsPro.Range("U" & LastRowPro2 + 7).Copy ' ������ ����
                                                    wsPro.Range("U" & LastRowPro2 + 7 & ":Z" & LastRowPro2 + 7).PasteSpecial Paste:=xlPasteFormulas ' ������ ����




'
' ���� "��������� �����" �� ����� "��������" � ������� R
Set FoundCell5 = Vb.Sheets("��������").Columns("R").Find(What:="��������� �����", LookIn:=xlValues, LookAt:=xlPart)
' ���� ������� "��������� �����", �������� �������� �� �������� ������ (�� 1 ������� ������)
If Not FoundCell5 Is Nothing Then
    ' �������� �������� �� ������ ������ �� "��������� �����"
    copyValue = FoundCell5.Offset(0, 1).Value
    ' ���� "���� (�) ��" �� ����� "������� �� ���" � ������� AC
'    Set wsSvodSMU = P3.Sheets("������� �� ���")
    Set FoundCell55 = wsSvodSMU.Columns("AC").Find(What:="���� (�) ��", LookIn:=xlValues, LookAt:=xlPart)
    ' ���� ������� "���� (�) ��", ��������� ������������� �������� � �������� ������ (�� 1 ������� ������)
    If Not FoundCell55 Is Nothing Then
        FoundCell55.Offset(0, 1).Value = copyValue    ' ��������� �������� � ������ ������ �� "���� (�) ��"
'        FoundCell55.Offset(1, 1).Value = Delta        ' ��������� �������� Delta � ������ ���� � ������ �� "���� (�) ��"
        ' ���������, ���� �������� � FoundCell55.Offset(1, 1) ������ -10 ��� ������ 10
        
        
        
        
        
        
        
        If FoundCell55.Offset(1, 1).Value < -10 Or FoundCell55.Offset(1, 1).Value > 10 Then
                ' ��������� ������� ��� (RGB(219, 179, 182))
                 wsPro.Cells(filledRow5 - 1, "B").Interior.Color = RGB(219, 179, 182) ' �������
            ' ��������� ��������� � ������
          wsPro.Cells(filledRow5 - 1, "B") = "��������� �� ����� '������� �� ���' �������� � ������� W '���� (�) ��' �� �����������"
'             "��������� �� ����� '������� �� ���' �������� �� ������������� � ������� AD '���� (�) ��'"
        End If
    Else
        MsgBox "�������� '���� (�) ��' �� ������� �� ����� '������� �� ���'.", vbExclamation
    End If
Else
    MsgBox "�������� '��������� �����' �� ������� �� ����� '��������'.", vbExclamation
End If
' ���������� �������� ��� ��������
Set checkRange = wsSvodSMU.Range("W" & FoundCell2.Row + 3 & ":W" & FoundCell2.Row + 9).Offset(0, 1)
Dim hasExceeded As Boolean
hasExceeded = False ' ���� ��� �������� ������� �������� > 1 ��� < -1

' ���������� ������ ������ � ���������
For Each cell In checkRange
    If Not IsEmpty(cell) And IsNumeric(cell.Value) Then
        Dim cellValue As Double
        cellValue = CDbl(cell.Value) ' ����������� �������� � �����
        If cellValue > 1 Or cellValue < -1 Then
            hasExceeded = True ' ���� ������� �������� > 1 ��� < -1, ������������� ����
            Exit For ' ���������� ����, ��� ��� ������� ���������
        End If
    End If
Next cell
















' ���� ���� �� ���� �������� ������ 1 ��� ������ -1, ��������� ��������
If hasExceeded Then
    ' �������� ������ �� ����� "������� �� ��������" � ������� D, ������ filledRow - 1
  ' ��������� ������� ��� (RGB(219, 179, 182))
                 wsPro.Cells(filledRow5 - 1, "D").Interior.Color = RGB(219, 179, 182) ' �������
            ' ��������� ��������� � ������
          wsPro.Cells(filledRow5 - 1, "D") = "��������� �� ����� '������� �� ���' �������� � ������� W '���� (�) ��' �� �����������"
End If
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
   ' ���������� ������ ������ � ��������� C12:H12 �� ����� wsSvodSMU                ��������
For Each CellsMU In wsSvodSMU.Range("C13:H13")
    ' ������� ��������������� ������ � ��������� D4:I4 �� ����� wsPro
    Dim cellProrabs As Range
    Set cellProrabs = wsPro.Range("D4").Offset(0, CellsMU.Column - wsSvodSMU.Range("C13").Column)
        ' �������� �������� �� ����� � ��������� �� �� 2 ������ ����� �������
    If IsNumeric(CellsMU.Value) Then
        valueSMU = Round(CellsMU.Value, 2)
    Else
        valueSMU = 0 ' ���� �������� �� �����, ����������� 0
    End If
        If IsNumeric(cellProrabs.Value) Then
        valueProrabs = Round(cellProrabs.Value, 2)
    Else
        valueProrabs = 0 ' ���� �������� �� �����, ����������� 0
    End If
        ' ���������� ����������� ��������
    If valueSMU <> valueProrabs Then
        ' ���� �������� �� �����, ���������� ��������������� ������ �� ����� "������� �� ��������"
        cellProrabs.Interior.Color = RGB(219, 179, 182)
    End If
Next CellsMU
       
       
       
       
       
       
       
       
       
' ������������� ���� �������
wsPro.Range("J6:J" & filledRow5 - 2).Interior.Color = RGB(146, 205, 220)
' ������������� ������ �������
wsPro.Range("J6:J" & filledRow5 - 2).Interior.Pattern = xlSolid



        ' �������� ������ ����� � ��������� ��������
                                                                                Dim checkRow As Long
                                                                                Dim emptyCellsMsg As String
                                                                                emptyCellsMsg = ""
                                                                                For checkRow = insertRow To insertRow + KolProv - 1
                                                                                    ' �������� ������� O (�����)
                                                                                    If IsEmpty(wsFact.Range("O" & checkRow)) Then
                                                                                        wsFact.Range("O" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ����� � ������ O" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� C (���)
                                                                                    If IsEmpty(wsFact.Range("C" & checkRow)) Then
                                                                                        wsFact.Range("C" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ��� � ������ C" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� D (���)
                                                                                    If IsEmpty(wsFact.Range("D" & checkRow)) Then
                                                                                        wsFact.Range("D" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ��� � ������ D" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� E (����� ������)
                                                                                    If IsEmpty(wsFact.Range("E" & checkRow)) Then
                                                                                        wsFact.Range("E" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ����� ������ � ������ E" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� F (����)
                                                                                    If IsEmpty(wsFact.Range("F" & checkRow)) Then
                                                                                        wsFact.Range("F" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ���� � ������ F" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� G (�����)
                                                                                    If IsEmpty(wsFact.Range("G" & checkRow)) Then
                                                                                        wsFact.Range("G" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ����� � ������ G" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� L (��)
                                                                                    If IsEmpty(wsFact.Range("L" & checkRow)) Then
                                                                                        wsFact.Range("L" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� �� � ������ L" & checkRow & vbCrLf
                                                                                    End If
                                                                                    ' �������� ������� N (��. ���.)
                                                                                    If IsEmpty(wsFact.Range("N" & checkRow)) Then
                                                                                        wsFact.Range("N" & checkRow).Interior.Color = RGB(219, 179, 182)
                                                                                        emptyCellsMsg = emptyCellsMsg & "�� ����� ""���� ��..."" ����������� ��. ���. � ������ N" & checkRow & vbCrLf
                                                                                    End If
                                                                                Next checkRow
                                                                                ' ����� ��������� ������������, ���� ������� ������ ������
                                                                                If emptyCellsMsg <> "" Then
                                                                                    MsgBox emptyCellsMsg, vbExclamation, "������ ������"
                                                                                End If

' ���� �� wsFact. � �������� P Q R S  �������, ������� ��������� � ������ ������� ����� ��������� �������� �������� 0, ��
' ������ ��������� "�� ����� ""���� ��..."" �� �� ������������ � ������ � ������ " � ������� � ��������� ����� ������ (�����) � �������� ������ RGB(219, 179, 182) � ������� L
'' ���� �� wsFact. � �������� P Q R S  �������, ������� ��������� � ������ ������� ����� ��������� �������� �������� 0, ��
' ������ ��������� "�� ����� ""���� ��..."" �� �� ������������ � ������ � ������ " � ������� � ��������� ����� ������ (�����) � �������� ������ RGB(219, 179, 182) � ������� L
' ���������� ����������
Dim checkRowFormula As Long
Dim zeroFormulaCellsMsg As String ' ���������� ��� �������� ��������� �� �������
zeroFormulaCellsMsg = "" ' ������������� ����������

For checkRowFormula = insertRow To insertRow + KolProv - 1
    ' ��������� ������ � �������� P, Q, R, S
    Dim cellFactP As Range, cellFactQ As Range, cellFactR As Range, cellFactS As Range
    Set cellFactP = wsFact.Cells(checkRowFormula, "P")
    Set cellFactQ = wsFact.Cells(checkRowFormula, "Q")
    Set cellFactR = wsFact.Cells(checkRowFormula, "R")
    Set cellFactS = wsFact.Cells(checkRowFormula, "S")
    
    ' ���������, ���� ��� ������ �������� ������� � �� �������� ����� 0
    If cellFactP.HasFormula And cellFactQ.HasFormula And cellFactR.HasFormula And cellFactS.HasFormula Then
        If cellFactP.Value = 0 And cellFactQ.Value = 0 And cellFactR.Value = 0 And cellFactS.Value = 0 Then
            ' ��������� ���������� �� ������
            zeroFormulaCellsMsg = zeroFormulaCellsMsg & "�� ����� ""���� ��..."" �� �� ������������ � ������ � ������� P" & checkRowFormula & ", Q" & checkRowFormula & ", R" & checkRowFormula & ", S" & checkRowFormula & vbCrLf
            
            ' �������� ��������������� ������ � ������� L ������-������� ������
            wsFact.Cells(checkRowFormula, "L").Interior.Color = RGB(219, 179, 182)
        End If
    End If
Next checkRowFormula

' ����� ��������� ������������, ���� ������� ������
If zeroFormulaCellsMsg <> "" Then
    MsgBox zeroFormulaCellsMsg, vbExclamation, "���������� �� � ������"
End If




    ' � ������ A6
    Application.GoTo Range("A6"), True
'
    Application.ScreenUpdating = True '�������� ���������� ������ ����� ������� �������
    Application.Calculation = xlCalculationAutomatic '������� ������ - ����� � �������������� ������
    Application.EnableEvents = True  '�������� �������
    If Workbooks.Count Then
    ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True '���������� ������� �����
    End If
    Application.DisplayStatusBar = True '���������� ��������� ������
    Application.DisplayAlerts = True '���������

    MsgBox "������ �� ����� ""��������"" ��������� � " & filePath & " �� ���� ""���� �� �� ������� ����"". ������� ������ �� ����� ""������� �� ��������"", ������� �������� ��������� � ���� ������ (��� �������� �� ������������ ��� � ��������� ������� 40. ����� ����� ������� ����������� � ������ 4 � ����� 01 ������. ��� �������� �� ���������� �������� � ������ ""������� �� ���"""
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
' ������� ��� �������� ����������� ���� ������ ����� � ��������� ��������� �����
Function OpenFileDialog(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "�������� ���� ""�3 ���� ��� _ ����..."" �� �������� ������"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' ��������� ��������� �����, ���� ��� ������
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


