Attribute VB_Name = "ģ��3"
Sub �����ֺ�Ϊð��()
Attribute �����ֺ�Ϊð��.VB_Description = "�����ֺ�Ϊð�ţ����������屨���ж���Ԥ����Ϣ�����������"
Attribute �����ֺ�Ϊð��.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' �����ֺ�Ϊð�� ��
'

    
    Cells.Replace What:=";", Replacement:=":", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

'    ���ļ�����ʾ��
'    MsgBox ThisWorkbook.Path
'    If Len(Dir(FN)) = 0 Then
'    MsgBox "�Ҳ����ļ���" & vbCrLf & FN, vbExclamation, "����"
'    Else
'    Workbooks.Open Filename:=ThisWorkbook.Path & "\����.xls"
'    End If

'�����ļ���·��FN
    Dim FN As String
    Dim idate As Date
    idate = Format(Now, "yyyy/m/d")
    FN = ThisWorkbook.Path & "\��ʯ�ֹ�˾ƽ̨����" & Month(idate) & "." & Day(idate) - 1 & "\"
'    MsgBox "FN : " & FN
'    MsgBox Dir(FN)


'    Dim jsc As String
'    Dim bnddz As String
'    Dim lj As String
'    Dim yj As String
'    Dim dq As String

'�����屨����file_lzw
    Dim file_lzw
'ģ�������������ļ�·��FileAddress_lzw
    Dim FileAddress_lzw
'�����屨��·��lzw
    Dim lzw
    
    FileAddress_lzw = FN & "*����GPS��������Ŧվ*.xlsx"
    file_lzw = Dir(FileAddress_lzw)
    lzw = FN & file_lzw
'    MsgBox "lzw: " + lzw
'    MsgBox "file_lzw: " + file_lzw
'    jsc = FN & "����GPS��ʯ��.xlsx"
'    bnddz = FN & "����GPS���ϴ����.xlsx"
'    lj = FN & "����GPS¹��.xlsx"
'    yj = FN & "����GPSԶ��.xlsx"
'    dq = FN & "����GPS��Ȫ.xlsx"
'    MsgBox Dir(FN & lzw)
'    Dim a
'    a = Dir(FN & "*����GPS*.xlsx")
'    Workbooks.Open FN + a
'    Do
'    a = Dir
'    If a <> "" Then
'    Workbooks.Open FN + a
'    End If
'    Loop
    
''�������屨��
'    If Len(file_lzw) = 0 Then
'    MsgBox "�Ҳ����ļ���" & vbCrLf & FileAddress_lzw, vbExclamation, "����"
'    Else
'    Workbooks.Open filename:=lzw
'    End If
    
    Call open_file(FN, "����GPS��������Ŧվ", "xlsx")

''�򿪶���GPS����
'    If Len(Dir(FN & "����GPS��ʯ.xlsx")) = 0 Then
'    MsgBox "�Ҳ����ļ���" & vbCrLf & FN & "����GPS��ʯ.xlsx", vbExclamation, "����"
'    Else
'    Workbooks.Open file_lzw:=FN + "����GPS��ʯ.xlsx"
'    End If

    Call open_file(FN, "����GPS��ʯ", "xlsx")

'-------------------------------------
'���Ƹ���������λ�ʴ���������ܵ��������� (δ���)
'    Windows(file_lzw).Activate
'    Sheets("186").Select
'    Dim a
'
'    a = Range("I10").Value
'
''    MsgBox "a: " + a
'
'    Windows("����GPS��ʯ.xlsx").Activate
'    Range("I21") = a + ","
'    End
'----------------------------------------

'������GPS�����屨����Ԥ����Ϣ������1��
'    Windows("����GPS��ʯ.xlsx").Activate
'    Range("A2:W2").Select
'    ActiveCell.FormulaR1C1 = _
'        "��λ����ʯ�ֹ�˾                         ��̨����139̨                                                           ʱ�䣺2019��" & Month(idate) & "��" & (Day(idate) - 1) & "��"
'    Range("J6:W6").Select

'    Range("I4").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R5C19"
'    Range("J4:W4").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R5C20"
'    Range("I5").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R9C19"
'    Range("J5:W5").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R9C20"
'    Range("I6").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R12C19"
'    Range("J6:W6").Select
'    ActiveCell.FormulaR1C1 = "='[" + file_lzw + "]393'!R12C20"

''����I4:W6��ֵ��ճ����I4:W6
'    Range("I4:W6").Select
'    Selection.Copy
'    Range("I4").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("I20").Select

'
''����
'    ActiveWorkbook.Save
''    ActiveWindow.Close
'
'
''ɾ������GPS�����屨����393�Ķ���Ԥ����Ϣ
'    Windows(file_lzw).Activate
'    Sheets("393").Select
'    Range("S5:T12").Select
'    Selection.Delete Shift:=xlToLeft
'    Range("I10:N12").Select
'    ActiveWorkbook.Save
''    ActiveWindow.Close

'������GPS�����屨����Ԥ����Ϣ(����2)
    Windows(file_lzw).Activate
    Sheets("393").Select
    Dim time1, time2, time3
    Dim msg1, msg2, msg3
    time1 = Range("s5").Value
    time2 = Range("s8").Value
    time3 = Range("s12").Value
    msg1 = Range("t5").Value
    msg2 = Range("t8").Value
    msg3 = Range("t12").Value

'ɾ������GPS�����屨����393�Ķ���Ԥ����Ϣ
    Range("S5:T12") = ""
    ActiveWorkbook.Save

    Windows("����GPS��ʯ.xlsx").Activate
    Range("i4") = time1
    Range("j4") = msg1
    Range("i5") = time2
    Range("j5") = msg2
    Range("i6") = time3
    Range("j6") = msg3
    Range("a2") = "��λ����ʯ�ֹ�˾                         ��̨����139̨                                                           ʱ�䣺2019��" & Month(idate) & "��" & (Day(idate) - 1) & "��"
    ActiveWorkbook.Save
'    ActiveWindow.Close

End Sub

'Sub open_all_files(����)
'Dim a
'a = Dir("C:\Users\Administrator\Desktop\�½��ļ���\*.txt") '��txt��β�������ļ��򿪣�����������ֻ�򿪵�һ�����ϵ��ļ������������ļ���doѭ�������δ�
'Workbooks.Open "C:\Users\Administrator\Desktop\�½��ļ���\" + a
'Do '����Ŀ¼�µ�����ָ����ʽ���ļ���
'a = Dir '֮ǰdir()���Ѿ����˶���ļ�������Ͳ�����д�ϣ���ʾ���δ򿪷��ϸ�ʽ���ļ�
'If a <> "" Then
'Workbooks.Open "C:\Users\Administrator\Desktop\�½��ļ���\" + a '��ÿһ�����ϸ�ʽ���ļ�
'Else
'Exit Sub
'End If
'Loop

'End Sub

Sub open_file(FN, key, extension)
    'open_file(FN, key, extension)��·�����ļ��ؼ��֡��ļ���׺�����ļ�
    '�ļ���·��FN���ļ��ؼ���key���ļ���׺��extension
    
'    MsgBox "FN:" & FN & ",file:" & file
    Dim file
    file = "*" & key & "*." & extension
'    MsgBox file
    If Len(Dir(FN & file)) = 0 Then
'   MsgBox (Dir(FN & file))
        MsgBox "�Ҳ����ļ���" & vbCrLf & FN & file, vbExclamation, "����"
'End�����ֹ��������
        End
    Else
    Workbooks.Open filename:=FN & Dir(FN & file)
    End If
    
End Sub
