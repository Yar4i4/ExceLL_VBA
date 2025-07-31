Attribute VB_Name = "Module5"
  Sub ��_P3_�_����_��()
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
    
    ' ��������� ���������� ���� ��� ������ �����
    Dim filePath As Variant
    filePath = OpenFileDialog2(Vb.Path) ' ��������� �����, ��� ��������� ������� �����
    
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
    If missingSheets <> "" Then
        MsgBox "����������� ��������� �����:" & vbCrLf & missingSheets, vbExclamation
        P3.Close SaveChanges:=False ' ��������� ���� ��� ���������� ���������
        Exit Sub
    End If
    
    ' ������� ����� �����
    Dim newBook As Workbook
    Set newBook = Workbooks.Add
       
    
    ' �������� ����� � ����� �����
    For Each sheetName In sheetNames
        P3.Sheets(sheetName).Copy After:=newBook.Sheets(newBook.Sheets.Count)
    Next sheetName
    
     ' ������� ���� "����1" �� ����� �����, ���� �� ����������
    On Error Resume Next ' ���������� ������, ���� ���� �� ����������
    Application.DisplayAlerts = False ' ��������� ��������������
    newBook.Sheets("����1").Delete ' ������� ����
    Application.DisplayAlerts = True ' �������� �������������� �������
    On Error GoTo 0 ' ���������� ����������� ��������� ������
    
    
    ' �������� � ������ "���� �� �� ������� ����" � ����� �����
    Dim ������ As Worksheet
    Set ������ = newBook.Sheets("���� �� �� ������� ����")
    
    ' ���������� �������� ��� ��������� (������� P, Q, R, S, T, Z, AA, AB)
    Dim LastRow As Long
    LastRow = ������.Cells(������.Rows.Count, "P").End(xlUp).Row ' ������� ��������� ����������� ������ � ������� P
    
    ' ������������ ������ �������
    Dim col As Variant
    For Each col In Array("P", "Q", "R", "S", "T", "Z", "AA", "AB")
        ' �������� ������� �� ��������
        ������.Range(col & "4:" & col & LastRow).Value = ������.Range(col & "4:" & col & LastRow).Value
    Next col
    
  




   
   
   
   
    ' �������� � ������ "������� �� ���" � ����� �����
    Dim ���������� As Worksheet
    Set ���������� = newBook.Sheets("������� �� ���")
    
    ' ���������� ������� ��� ���������
    Dim columnsToProcess As Variant
    columnsToProcess = Array("C", "D", "F", "G", "J", "K", "M", "N") ' ��������� ������� ��� ���������
    
    ' ������������ ������ �������
    Dim colSMU As Variant
    For Each colSMU In columnsToProcess
        ' ���������� �������� ��� ��������� (������� �������)
        Dim lastRow2 As Long
        lastRow2 = ����������.Cells(����������.Rows.Count, colSMU).End(xlUp).Row ' ������� ��������� ����������� ������ � �������
        
        ' ������������ ������ � �������
        Dim cell5 As Range
        For Each cell5 In ����������.Range(colSMU & "4:" & colSMU & lastRow2)
            If InStr(1, cell5.Formula, "[") > 0 Then
                ' ���� � ������� ���� ������ �� ������� ���� (���������� ������), ������� �
                Dim formulaText As String
                formulaText = cell5.Formula
                
                ' ������� ��� ������ �� ������� ����� � �������
                Do While InStr(1, formulaText, "[") > 0
                    Dim startPos As Long
                    Dim endPos As Long
                    startPos = InStr(1, formulaText, "[") ' ������� ������ ������
                    endPos = InStr(startPos, formulaText, "]") ' ������� ����� ������
                    
                    If startPos > 0 And endPos > 0 Then
                        ' ������� ���� � ����� � ��������� ������ ��� ����� � ������
                        formulaText = Left(formulaText, startPos - 1) & Mid(formulaText, endPos + 1)
                    Else
                        Exit Do
                    End If
                Loop
                
                ' ��������� ������� � ������
                cell5.Formula = formulaText
            End If
        Next cell5
    Next colSMU
   









' �������� � ������ "������� �� ��������" � ����� �����
    Dim �������������� As Worksheet
    Set �������������� = newBook.Sheets("������� �� ��������")
    
    ' ���������� ������� ��� ��������� (�� D �� Z)
    Dim columnsToProcessPro As Variant
    columnsToProcessPro = Array("D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "U", "V", "W", "X", "Y", "Z")
    
    ' ������������ ������ �������
    Dim colPro As Variant
    For Each colPro In columnsToProcessPro
        ' ���������� �������� ��� ��������� (������� �������)
        Dim lastRowPro As Long
        lastRowPro = ��������������.Cells(��������������.Rows.Count, colPro).End(xlUp).Row ' ������� ��������� ����������� ������ � �������
        
        ' ������������ ������ � �������
        Dim cellPro As Range
        For Each cellPro In ��������������.Range(colPro & "4:" & colPro & lastRowPro)
            If InStr(1, cellPro.Formula, "[") > 0 Then
                ' ���� � ������� ���� ������ �� ������� ���� (���������� ������), ������� �
                Dim formulaTextPro As String
                formulaTextPro = cellPro.Formula
                
                ' ������� ��� ������ �� ������� ����� � �������
                Do While InStr(1, formulaTextPro, "[") > 0
                    Dim startPosPro As Long
                    Dim endPosPro As Long
                    startPosPro = InStr(1, formulaTextPro, "[") ' ������� ������ ������
                    endPosPro = InStr(startPosPro, formulaTextPro, "]") ' ������� ����� ������
                    
                    If startPosPro > 0 And endPosPro > 0 Then
                        ' ������� ���� � ����� � ��������� ������ ��� ����� � ������
                        formulaTextPro = Left(formulaTextPro, startPosPro - 1) & Mid(formulaTextPro, endPosPro + 1)
                    Else
                        Exit Do
                    End If
                Loop
                
                ' ��������� ������� � ������
                cellPro.Formula = formulaTextPro
            End If
        Next cellPro
    Next colPro



' ��������� ������ ���� ��� ���������� �����
Dim folderPath As String
folderPath = Vb.Path ' ���������� ���� ������� ����� (ThisWorkbook)

' ���������, ���������� �� �����
If Dir(folderPath, vbDirectory) = "" Then
    MsgBox "����� ��� ���������� �� ����������: " & folderPath, vbExclamation
    Exit Sub
End If

' ��������� ��� �����
Dim fileName As String
With P3.Sheets("������� �� ��������")
    ' ��������� ����� ���� �� ������ A1
    Dim YearPart As String
    Dim monthPart As String
    Dim dayPart As String

    ' ���: 4 ������� ������
    YearPart = Right(.Range("A1").Value, 4) ' ��������� 4 �������

    ' �����: 7� � 6� ������� ������
    monthPart = Mid(.Range("A1").Value, Len(.Range("A1").Value) - 6, 2)

    ' ����: 10� � 9� ������� ������
    dayPart = Mid(.Range("A1").Value, Len(.Range("A1").Value) - 9, 2)

    ' ��������� ��� �����
    fileName = "���� �� �� ���� _ " & YearPart & "." & monthPart & ".01-" & dayPart & ".xlsb"
End With

' ��������� ������ ����
Dim fullPath As String
fullPath = folderPath & "\" & fileName


' ��������� ��� ������� ����� � ����� �����
Dim link As Variant
For Each link In newBook.LinkSources(xlExcelLinks)
    newBook.BreakLink Name:=link, Type:=xlLinkTypeExcelLinks
Next link


' ��������� ����� �����
Application.DisplayAlerts = False ' ��������� ��������������
newBook.SaveAs fileName:=fullPath, FileFormat:=xlExcel12
Application.DisplayAlerts = True ' �������� �������������� �������



    Application.ScreenUpdating = True '�������� ���������� ������ ����� ������� �������
    Application.Calculation = xlCalculationAutomatic '������� ������ - ����� � �������������� ������
    Application.EnableEvents = True  '�������� �������
    If Workbooks.Count Then
    ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True '���������� ������� �����
    End If
    Application.DisplayStatusBar = True '���������� ��������� ������
    Application.DisplayAlerts = True '���������
  End Sub
' ������� ��� �������� ����������� ���� ������ ����� � ��������� ��������� �����
Function OpenFileDialog2(Optional InitialFolder As String = "") As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker) ' ���������� FileDialog
    
    With fd
        .Title = "�������� ���� ""�3 ���� ��� _ ����..."" �� �������� ������"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
        
        ' ��������� ��������� �����, ���� ��� ������
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

