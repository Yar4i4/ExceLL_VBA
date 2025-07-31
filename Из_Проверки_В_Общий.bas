Attribute VB_Name = "Module2"
Sub ��_��������_�_�����()
Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    If Workbooks.Count Then ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ' ��������� � �������������� ��������� ��� �������� ��������� �� �������
    Dim errorMessages As Collection
    Set errorMessages = New Collection
Dim Vb As Workbook
    Set Vb = ThisWorkbook
    ' ��������� ���������� ���� ��� ������ �����
    Dim filePath As Variant
    filePath = OpenFileDialog5(Vb.Path) ' ��������� �����, ��� ��������� ������� �����
    ' ���������, ��� �� ������ ����
    If filePath = False Then
        MsgBox "���� �� ������!", vbExclamation
        Exit Sub
    End If
' ��������� ��������� ����
Dim Ob As Workbook
Set Ob = Workbooks.Open(filePath)
'    ' ���������, ������� �� �������� ����
'    If Ob Is Nothing Then
'        errorMessages.Add "�� ������� ������� ����: " & filePath
'        Exit Sub
'    End If
' ���������, ���������� �� ���� "������ ��� ""�-�����"""
Dim targetSheet As Worksheet
On Error Resume Next
Set targetSheet = Ob.Worksheets("������ ��� ""�-�����""")
On Error GoTo 0
If targetSheet Is Nothing Then
    MsgBox "���� ""������ ��� ""�-�����"""" �� ������!", vbExclamation
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

    
                                                                                  ' ���������, ���� �� �������� � ������
                                                                                    Dim volumeValue As Variant
                                                                                    volumeValue = Ob.Worksheets("������ ��� ""�-�����""").Cells(9, lCol).Value
                                                                                    If Not IsEmpty(volumeValue) And volumeValue <> 0 Then
                                                                                    ' �������� ��������� ����������� �������
                                                                                    Dim columnLetter As String
                                                                                    columnLetter = Split(Ob.Worksheets("������ ��� ""�-�����""").Cells(1, lCol).Address, "$")(1)
                                                                                    ' ���� �������� ����, ���������� ������������, ���������� �� ����������
                                                                                    Dim userResponse As VbMsgBoxResult
                                                                                    userResponse = MsgBox("� ""������ ��� ""�-�����"" � ������� " & columnLetter & " (��) ��� ���� �������� � ��������." & vbCrLf & _
                                                                                                    "���������� ���������� �������?", vbYesNo + vbExclamation, "��������")
                                                                                    ' ���� ������������ ������ "���", ��������� ����������
                                                                                    If userResponse = vbNo Then
                                                                                    Ob.Close SaveChanges:=False
                                                                                    Exit Sub
                                                                                    End If
                                                                                    End If

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

  ' ����� ������ ������, ���������� "��� ������� / �������"
   Dim Rngt3 As Range
    Set Rngt3 = Cells.Find("��� ������� / �������", , xlFormulas, xlWhole)
    If Rngt3 Is Nothing Then
        errorMessages.Add "� ����� " & fileName & " ����������� ������: ""��� ������� / �������""."
        Ob.Close SaveChanges:=False
        Exit Sub
    End If
    Dim lCol3 As Long
    lCol3 = Rngt3.Column ' �������, ��� ���� ������� ����� "��� ������� / �������"

''� Prov.Range("D") ������� D ������� �� ������ 11 � �� ��������� ������ (LastRowDPQ) � ������� D ��������� ��������, ���������� ��������� ��� ������.
'����� ����� � ���������� ����� � Ob.Worksheets("������ ��� ""�-�����""").Range("D") ������� D ������� �� ������ 11 � �� ��������� ������ (LastRowDPQ2), ��� �� �������� �������� ������ �� Prov.Range("D"), � ���� ����� ����  �������� �������, ��
'1. �� ��������������� ������ �������� �������� (����������� � ������� Prov.Range("O11:O" & LastRowDPQ)) ����� �� ������� O ����������� �������� � ��������  �  Ob.Worksheets("������ ��� ""�-�����""").  �  ������� lCol (� ������� ������� �� ������ 11 � �� ��������� ������ (LastRowDPQ2)) � ��������������� ������, � ������� ��� ������ ������� �������� �� Prov.Range("D").
'2.  �� ��������������� ������ �������� �������� (����������� � ������� Prov.Range("P11:P" & LastRowDPQ)) ����� �� ������� P ����������� �������� � ��������  �  Ob.Worksheets("������ ��� ""�-�����""").  �  ������� lCol2 (� ������� ������� �� ������ 11 � �� ��������� ������ (LastRowDPQ2)) � ��������������� ������, � ������� ��� ������ ������� �������� �� Prov.Range("D").
'3.  �� ��������������� ������ �������� �������� (����������� � ������� Prov.Range("Q11:Q" & LastRowDPQ)) ����� �� ������� Q ����������� �������� � ��������  �  Ob.Worksheets("������ ��� ""�-�����""").  �  ������� lCol3 (� ������� ������� �� ������ 11 � �� ��������� ������ (LastRowDPQ2)) � ��������������� ������, � ������� ��� ������ ������� �������� �� Prov.Range("D").
'
'� ���� ����� ����  �������� �� �������, �� �������� ������ Prov.Range("A:D") � ��������������� ������, �� ������� �� ��� ������ ������� �������� �� Prov.Range("D") � �������� ���  RGB(219, 179, 182). � � ����� ������� �������� ��� ���� �� ��������� ��������� �  ����� ���� "�� ���� ���������� ��������� ����� � "������ ��� ""�-�����""
'
'� ���� ����� �������� ������� �� ����, �� �������� ������ Prov.Range("A:D") � ��������������� ������, �� ������� ���� ������� ����� 1 �������� �������� �� Prov.Range("D") � �������� ���  RRGB(129, 131, 143). � � ����� ������� �������� ��� ���� ��������� ��������� ����� ������ �  ����� ��������� ���� "���� ���������� ��� ��� ����� ������ � "������ ��� ""�-�����""
' ���������� ��������� ������ �� ����� "��������" � ������� D
    Dim Prov As Worksheet
    Set Prov = Vb.Worksheets("��������")
    Vb.Worksheets("��������").Activate
    Dim LastRowDPQ As Long
    LastRowDPQ = Prov.Cells(Prov.Rows.Count, "D").End(xlUp).Row
    
     

' ���������� ��������� ������ �� ����� "������ ��� ""�-�����""" � ������� D
Ob.Worksheets("������ ��� ""�-�����""").Activate
Dim LastRowDPQ2 As Long
LastRowDPQ2 = Ob.Worksheets("������ ��� ""�-�����""").Cells(Ob.Worksheets("������ ��� ""�-�����""").Rows.Count, "D").End(xlUp).Row


' ������� ��������� ��� �������� ��������� �� �������
Dim notFoundKeys As Collection
Dim duplicateKeys As Collection
Set notFoundKeys = New Collection
Set duplicateKeys = New Collection

' ���������� ��� �������� ��������� ����������
Dim foundCount As Long





' ���� �� ������� �� ����� "��������"
Dim i As Long
For i = 11 To LastRowDPQ

Application.ScreenUpdating = True
' ������������ ���������� ����������� ����� �� ������� ������ i
Dim filledRowsCount As Long
filledRowsCount = WorksheetFunction.CountA(Prov.Range("D11:D" & i))  ' �������� 1, ����� ��������� ������� ������ i
' ��������� �������� � ������ D1
Vb.Worksheets("�������").Cells(1, "D").Value = "���������� " & filledRowsCount & " ����� �� " _
                                                  & WorksheetFunction.CountA(Prov.Range("D11:D" & LastRowDPQ))
Application.ScreenUpdating = False
                                                      
    ' �������� �������� �� ������� D �� ����� "��������"
    Dim searchValue As Variant
    searchValue = Prov.Cells(i, "D").Value
    
    ' ���������� ������� ��������� ����������
    foundCount = 0
    
    ' ���� �� ������� �� ����� "������ ��� ""�-�����"""
    Dim j As Long
    For j = 11 To LastRowDPQ2
        ' �������� �������� �� ������� D �� ����� "������ ��� ""�-�����"""
        Dim targetValue As Variant
        targetValue = Ob.Worksheets("������ ��� ""�-�����""").Cells(j, "D").Value
        
        ' ���� �������� ���������
        If searchValue = targetValue Then
            foundCount = foundCount + 1
            
            ' ���� ������� ������ ����������, �������� ������ �� �������� O, P � Q
            If foundCount = 1 Then
                ' �������� �������� �� ������� O
                Ob.Worksheets("������ ��� ""�-�����""").Cells(j, lCol).Value = Prov.Cells(i, "O").Value
                ' �������� �������� �� ������� Q
                Ob.Worksheets("������ ��� ""�-�����""").Cells(j, lCol2).Value = Prov.Cells(i, "Q").Value
                ' �������� �������� �� ������� P
                Ob.Worksheets("������ ��� ""�-�����""").Cells(j, lCol3).Value = Prov.Cells(i, "P").Value
            End If
        End If
    Next j
    
     ' ���� ���������� �� �������
    If foundCount = 0 Then
        ' �������� ������ �� ����� "��������" � ���������� � � ���� RGB(219, 179, 182)
        Prov.Range("A" & i & ":D" & i).Interior.Color = RGB(219, 179, 182)
        Prov.Cells(i, "T").Value = Prov.Cells(i, "O").Value    ' �������� �������� �� ������ O � ��������� ��� � ������ T
        notFoundKeys.Add "���� �� ������: " & searchValue   ' ��������� ��������� �� ������ � ���������
    ' ���� ������� ����� ������ ����������
    ElseIf foundCount > 1 Then
        ' �������� ������ �� ����� "��������" � ���������� � � ���� RGB(129, 131, 143)
        Prov.Range("A" & i & ":D" & i).Interior.Color = RGB(255, 242, 204)
        ' ��������� ��������� �� ������ � ���������
        duplicateKeys.Add "������� ����� ������ �����: " & searchValue
    End If
Next i
 Prov.Range("T8") = "�� ������������� �����"
 Prov.Range("U8") = "������"
 Prov.Range("U9") = "=O9-T9"
  Prov.Range("T9").FormulaLocal = "=�������������.�����(9;T11:T" & LastRowDPQ & ")"
 ' ������ �������� �� ��������������� ������
    Dim sCell As Range
    For Each sCell In Prov.Range("S11:S" & LastRowDPQ)    ' �������� �������� � ������� S � ����������� �����
        If Not IsEmpty(sCell.Value) And Trim(sCell.Value) <> "" Then
            Prov.Range("A" & sCell.Row & ":R" & sCell.Row).Interior.Color = RGB(218, 238, 243)   '�����
        End If
    Next sCell


 
 












' ������� ��������� �� �������
If notFoundKeys.Count > 0 Then
    Dim notFoundMessage As String
    notFoundMessage = "�� ���� ���������� ��������� ����� � ""������ ��� ""�-�����"": " & vbCrLf
    For i = 1 To notFoundKeys.Count
        notFoundMessage = notFoundMessage & notFoundKeys(i) & vbCrLf
    Next i
    MsgBox notFoundMessage, vbExclamation, "������"
End If

If duplicateKeys.Count > 0 Then
    Dim duplicateMessage As String
    duplicateMessage = "���� ���������� ��� ��� ����� ������ � ""������ ��� ""�-�����"": " & vbCrLf
    For i = 1 To duplicateKeys.Count
        duplicateMessage = duplicateMessage & duplicateKeys(i) & vbCrLf
    Next i
    MsgBox duplicateMessage, vbExclamation, "������"
End If

' ��������������� ��������� Excel
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

'' ��������� ���� Ob ��� ���������� ���������
'Ob.Close SaveChanges:=False

End Sub



' ������� ��� �������� ����������� ���� ������ ����� � ��������� ��������� �����
Function OpenFileDialog5(Optional InitialFolder As String = "") As Variant
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
            OpenFileDialog5 = .SelectedItems(1) ' ���������� ��������� ����
        Else
            OpenFileDialog5 = False ' ���� ���� �� ������
        End If
    End With
End Function


