Attribute VB_Name = "OJW__5"
Option Explicit

Sub OJW_5()
Dim sh0 As Worksheet, sh1 As Worksheet, sh2 As Worksheet, sh3 As Worksheet, sh4 As Worksheet, sh5 As Worksheet, sh6 As Worksheet, sh6_1 As Worksheet, sh6_2 As Worksheet, _
    sh7 As Worksheet, sh7_1 As Worksheet, sh7_2 As Worksheet, sh8 As Worksheet, sh8_1 As Worksheet, sh8_2 As Worksheet, sh9 As Worksheet, sh9_1 As Worksheet, sh9_2 As Worksheet, _
    sh10 As Worksheet, sh11 As Worksheet, sh12 As Worksheet, sh13 As Worksheet, sh13_1 As Worksheet, sh13_2 As Worksheet, sh14 As Worksheet, sh15 As Worksheet, sh16 As Worksheet, _
    sh17 As Worksheet, sh18 As Worksheet, sh19 As Worksheet, sh20 As Worksheet
Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long, o As Long, p As Long, t As Long
Dim lLastRow1 As Long, lLastRow2 As Long, lLastRow3 As Long, lLastRow4 As Long, lLastRow5 As Long, lLastRow7 As Long, lLastRow9 As Long, lLastRow10 As Long, lLastRow11 As Long, lLastRow12 As Long
Dim MinDatePP As Date, MinDateGS As Date, MinDateOS As Date, MinDateFS As Date, MaxDatePP As Date, MaxDateGS As Date, MaxDateOS As Date, MaxDateFS As Date, D_ICJ_Min As Date, D_ICJ_Max As Date, MaxDateADG As Date, MaxDate As Date
Dim TU1 As String, TU2 As String, TU_K1 As String, TU_K2 As String, TU_K3 As String ' ��/���� �� ��������. ��������� �������� Empty ������ ��� ����� �������
Dim SP As String ' ������� ��������
Dim r As Long, q As Long, s As Long


Set sh0 = ActiveWorkbook.Sheets("UNIQ")
Set sh1 = ActiveWorkbook.Sheets("DRAFT")
Set sh2 = ActiveWorkbook.Sheets("Titul")
Set sh3 = ActiveWorkbook.Sheets("�� ������")
Set sh4 = ActiveWorkbook.Sheets("���������")
Set sh5 = ActiveWorkbook.Sheets("������")
Set sh6 = ActiveWorkbook.Sheets("����_��") '���� ���������� �����������
Set sh7 = ActiveWorkbook.Sheets("����_��") '���� ��������� ������������� ����
Set sh8 = ActiveWorkbook.Sheets("����_��") '���� ��������� ��������� ����. ������������ ������ �� ������� 4
Set sh9 = ActiveWorkbook.Sheets("����_��") '���� ��������� ��������� ����
Set sh19 = ActiveWorkbook.Sheets("������ 5") '������ 5

lLastRow1 = sh1.Cells(Rows.Count, 10).End(xlUp).Row
MaxDateFS = Application.WorksheetFunction.Max(sh1.Range("DB3:DB" & lLastRow1)) '������������ ���� �������� ����
MaxDateADG = Application.WorksheetFunction.Max(sh1.Range("V3:V" & lLastRow1)) '������������ ���� ��������� �������

MaxDate = IIf(MaxDateFS > MaxDateADG, MaxDateFS, MaxDateADG)

sh19.Cells(5, 1).Value = "1"
sh19.Cells(5, 2).Value = "��� ������������������� ������� ����� " & sh6.Cells(34, 2).Value & _
        ". ���������� ����������� ������������� ��� ���������������� ������ (�������, �������������, �������������)�� ������� ��������: " & sh0.Cells(2, 4).Value & _
            ". " & sh4.Cells(15, 6).Value
For o = 2 To 4 '�������� ����������� ��� ��� ������ 5
    If sh6.Cells(33, 25).Value >= sh2.Cells(o, 10).Value And sh6.Cells(33, 25).Value <= sh2.Cells(o, 11).Value Then
        sh19.Cells(5, 3).Value = sh6.Cells(33, 25).Value & ". " & sh2.Cells(o, 12).Value
        Exit For
    Else
    End If
Next o

sh19.Cells(6, 1).Value = "2"
sh19.Cells(6, 2).Value = "��� ������������������� ������� ����� " & sh7.Cells(34, 2).Value & _
        ". ��������� ������������� ���� ����������������� �������� �� ����������� �������� ������������� �� ������� ��������: " & sh0.Cells(2, 4).Value & _
            ". " & sh4.Cells(15, 6).Value
For o = 2 To 4 '�������� ����������� ��� ��� ������ 5
    If sh7.Cells(33, 25).Value >= sh2.Cells(o, 10).Value And sh7.Cells(33, 25).Value <= sh2.Cells(o, 11).Value Then
        sh19.Cells(6, 3).Value = sh7.Cells(33, 25).Value & ". " & sh2.Cells(o, 12).Value
        Exit For
    Else
    End If
Next o

SP = sh0.Cells(2, 4).Value
If SP = "4" Then
    sh19.Cells(7, 1).Value = "3"
    sh19.Cells(7, 2).Value = "��� ������������������� ������� ����� " & sh8.Cells(34, 2).Value & _
            ". ��������� �������������� ���� ����������������� �������� �� ����������� �������� ������������� �� ������� ��������: " & sh0.Cells(2, 4).Value & _
            ". " & sh4.Cells(15, 6).Value
    For o = 2 To 4 '�������� ����������� ��� ��� ������ 5
        If sh8.Cells(33, 25).Value >= sh2.Cells(o, 10).Value And sh8.Cells(33, 25).Value <= sh2.Cells(o, 11).Value Then
            sh19.Cells(7, 3).Value = sh8.Cells(33, 25).Value & ". " & sh2.Cells(o, 12).Value
            Exit For
        Else
        End If
    Next o
    
    sh19.Cells(8, 1).Value = "4"
    sh19.Cells(8, 2).Value = "��� ������������������� ������� ����� " & sh9.Cells(34, 2).Value & _
            ". ��������� ��������� ���� ����������������� �������� �� ����������� �������� ������������� �� ������� ��������: " & sh0.Cells(2, 4).Value & _
            ". " & sh4.Cells(15, 6).Value
    For o = 2 To 4 '�������� ����������� ��� ��� ������ 5
        If sh9.Cells(33, 25).Value >= sh2.Cells(o, 10).Value And sh9.Cells(33, 25).Value <= sh2.Cells(o, 11).Value Then
            sh19.Cells(8, 3).Value = MaxDate & ". " & sh2.Cells(o, 12).Value
            Exit For
        Else
        End If
    Next o
    
'    sh19.Cells(9, 1).Value = "5"
'    sh19.Cells(9, 2).Value = "��� ������� ��������� �������� " & "� ���-" & sh0.Cells(2, 3).Value & "-" & "AKZ02.03-" & sh0.Cells(2, 4).Value & "-" & Format(sh1.Cells(3, 44).Value, "000") & _
'            " �� ������� ��������: " & sh0.Cells(2, 4).Value & ". " & sh4.Cells(15, 6).Value
'    For o = 9 To 10 '�������� ����������� ��� ��� ������ 5
'        If MaxDate >= sh2.Cells(o, 10).Value And MaxDate <= sh2.Cells(o, 11).Value Then
'            sh19.Cells(9, 3).Value = MaxDate & ". " & sh2.Cells(o, 12).Value
'            Exit For
'        Else
'        End If
'    Next o
    
    lLastRow2 = sh0.Cells(Rows.Count, 7).End(xlUp).Row
    For i = 2 To lLastRow2
        lLastRow3 = sh19.Cells(Rows.Count, 2).End(xlUp).Row
        sh19.Cells(lLastRow3 + 1, 1).Value = 2 + i
        sh19.Cells(lLastRow3 + 1, 2).Value = "�������� ����������� ������� �������� ������������� �������� ������� ������ �" & sh0.Cells(i, 7).Value
        sh19.Cells(lLastRow3 + 1, 3).Value = sh0.Cells(i, 8).Value & ". ������� ������������� ����������� ����� �. �., ��������� ������������� ����������� �������� �. �." '�������� ����������� ��� ��� ������ 5
    Next i
Else
    sh19.Cells(7, 1).Value = "3"
    sh19.Cells(7, 2).Value = "��� ������������������� ������� ����� " & sh9.Cells(34, 2).Value & _
            ". ��������� ��������� ���� ����������������� �������� �� ����������� �������� ������������� �� ������� ��������: " & sh0.Cells(2, 4).Value & _
                    ". " & sh4.Cells(15, 6).Value
    For o = 2 To 4 '�������� ����������� ��� ��� ������ 5
        If MaxDate >= sh2.Cells(o, 10).Value And MaxDate <= sh2.Cells(o, 11).Value Then
            sh19.Cells(7, 3).Value = MaxDate & ". " & sh2.Cells(o, 12).Value
            Exit For
        Else
        End If
    Next o
    
'    sh19.Cells(8, 1).Value = "4"
'    sh19.Cells(8, 2).Value = "��� ������� ��������� �������� " & "� ���-" & sh0.Cells(2, 3).Value & "-" & "AKZ02.03-" & sh0.Cells(2, 4).Value & "-" & Format(sh1.Cells(3, 44).Value, "000") & _
'            " �� ������� ��������: " & sh0.Cells(2, 4).Value & ". " & sh4.Cells(15, 6).Value
'    For o = 9 To 10 '�������� ����������� ��� ��� ������ 5
'        If MaxDate >= sh2.Cells(o, 10).Value And MaxDate <= sh2.Cells(o, 11).Value Then
'            sh19.Cells(8, 3).Value = MaxDate & ". " & sh2.Cells(o, 12).Value
'            Exit For
'        Else
'        End If
'    Next o
    
    lLastRow2 = sh0.Cells(Rows.Count, 7).End(xlUp).Row
    For i = 2 To lLastRow2
        lLastRow3 = sh19.Cells(Rows.Count, 2).End(xlUp).Row
        sh19.Cells(lLastRow3 + 1, 1).Value = 2 + i
        sh19.Cells(lLastRow3 + 1, 2).Value = "�������� ����������� ������� �������� ������������� �������� ������� ������ �" & sh0.Cells(i, 7).Value
        sh19.Cells(lLastRow3 + 1, 3).Value = sh0.Cells(i, 8).Value & ". ������� ������������� ����������� ����� �. �., ��������� ������������� ����������� �������� �. �." '�������� ����������� ��� ��� ������ 5
    Next i
End If

End Sub
