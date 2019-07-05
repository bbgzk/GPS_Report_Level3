Sub 更正分号为冒号()
'
' 更正分号为冒号 宏
'


    Cells.Replace What:=";", Replacement:=":", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

'    打开文件代码示例
'    MsgBox ThisWorkbook.Path
'    If Len(Dir(FN)) = 0 Then
'    MsgBox "找不到文件：" & vbCrLf & FN, vbExclamation, "错误"
'    Else
'    Workbooks.Open Filename:=ThisWorkbook.Path & "\报表.xls"
'    End If

'报表文件夹路径FN
    Dim FN As String
    Dim idate As Date
    idate = Format(Now, "yyyy/m/d")
    idate = idate - 1
    
    FN = ThisWorkbook.Path & "\界石分公司平台报表" & Month(idate) & "." & Day(idate) & "\"
'    MsgBox "FN : " & FN
'    MsgBox Dir(FN)


'    Dim jsc As String
'    Dim bnddz As String
'    Dim lj As String
'    Dim yj As String
'    Dim dq As String

'龙洲湾报表名file_lzw
    Dim file_lzw
'模糊查找龙洲湾文件路径FileAddress_lzw
    Dim FileAddress_lzw
'龙洲湾报表路径lzw
    Dim lzw

    FileAddress_lzw = FN & "*三级GPS龙洲湾枢纽站*.xlsx"
    file_lzw = Dir(FileAddress_lzw)
    lzw = FN & file_lzw
'    MsgBox "lzw: " + lzw
'    MsgBox "file_lzw: " + file_lzw
'    jsc = FN & "三级GPS界石场.xlsx"
'    bnddz = FN & "三级GPS巴南大道中.xlsx"
'    lj = FN & "三级GPS鹿角.xlsx"
'    yj = FN & "三级GPS远郊.xlsx"
'    dq = FN & "三级GPS东泉.xlsx"
'    MsgBox Dir(FN & lzw)
'    Dim a
'    a = Dir(FN & "*三级GPS*.xlsx")
'    Workbooks.Open FN + a
'    Do
'    a = Dir
'    If a <> "" Then
'    Workbooks.Open FN + a
'    End If
'    Loop

''打开龙洲湾报表
'    If Len(file_lzw) = 0 Then
'    MsgBox "找不到文件：" & vbCrLf & FileAddress_lzw, vbExclamation, "错误"
'    Else
'    Workbooks.Open filename:=lzw
'    End If

    Call open_file(FN, "三级GPS龙洲湾枢纽站", "xlsx")

''打开二级GPS报表
'    If Len(Dir(FN & "二级GPS界石.xlsx")) = 0 Then
'    MsgBox "找不到文件：" & vbCrLf & FN & "二级GPS界石.xlsx", vbExclamation, "错误"
'    Else
'    Workbooks.Open file_lzw:=FN + "二级GPS界石.xlsx"
'    End If

    Call open_file(FN, "二级GPS界石", "xlsx")

'-------------------------------------
'复制各三级报表定位率处理情况汇总到二级报表 (未完成)
'遍历各个三级报表中各个sheet（各线路）的定位率处理情况，并取值（取值判断：是否是5位数字加汉字，每个值后面加逗号）加入到队列，将队列的值赋给二级报表的定位率处理情况格子
'    Windows(file_lzw).Activate
'    Sheets("186").Select
'    Dim a
'
'    a = Range("I10").Value
'
''    MsgBox "a: " + a
'
'    Windows("二级GPS界石.xlsx").Activate
'    Range("I21") = a + ","
'    End
'----------------------------------------

'从三级GPS龙洲湾报表复制预警信息（方法1）
'    Windows("二级GPS界石.xlsx").Activate
'    Range("A2:W2").Select
'    ActiveCell.FormulaR1C1 = _
'        "单位：界石分公司                         车台数：140台                                                           时间：2019年" & Month(idate) & "月" & (Day(idate) - 1) & "日"
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

''复制I4:W6的值并粘贴到I4:W6
'    Range("I4:W6").Select
'    Selection.Copy
'    Range("I4").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("I20").Select

'
''保存
'    ActiveWorkbook.Save
''    ActiveWindow.Close
'
'
''删除三级GPS龙洲湾报表中393的多余预警信息
'    Windows(file_lzw).Activate
'    Sheets("393").Select
'    Range("S5:T12").Select
'    Selection.Delete Shift:=xlToLeft
'    Range("I10:N12").Select
'    ActiveWorkbook.Save
''    ActiveWindow.Close

'从三级GPS龙洲湾报表复制预警信息(方法2)
    Windows(file_lzw).Activate
    Sheets("393").Select
'错误捕获（未完成）
'    If (Sheets("394").Select = flase) Then
'        MsgBox ("表格393找不到错误")
'        End
'    End If
    
'把5、8、11行改成5、6、7行的数据
'    Dim time1, time2, time3
'    Dim msg1, msg2, msg3
'    time1 = Range("s5").Value
'    time2 = Range("s8").Value
'    time3 = Range("s11").Value
'    msg1 = Range("t5").Value
'    msg2 = Range("t8").Value
'    msg3 = Range("t11").Value

    Dim time1, time2, time3
    Dim msg1, msg2, msg3
    time1 = Range("s5").Value
    time2 = Range("s6").Value
    time3 = Range("s7").Value
    msg1 = Range("t5").Value
    msg2 = Range("t6").Value
    msg3 = Range("t7").Value


'删除三级GPS龙洲湾报表中393的多余预警信息
'    Range("S5:T12") = ""
    Range("S5:T12").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWorkbook.Save
'    ActiveWindow.Close

    Windows("二级GPS界石.xlsx").Activate
    Range("i4") = time1
    Range("j4") = msg1
    Range("i5") = time2
    Range("j5") = msg2
    Range("i6") = time3
    Range("j6") = msg3
    Dim a2
    a2 = Range("a2").Value
    Dim a22
    a22 = InStrRev(a2, "时间：") + 2
'    MsgBox (InStrRev(a2, "时间："))
'    MsgBox (Left(a2, a22) + "2019年" & Month(idate) & "月" & (Day(idate) - 1) & "日")
'    Range("a2") = "单位：界石分公司                         车台数：139台                                                           时间：2019年" & Month(idate) & "月" & (Day(idate) - 1) & "日"
    
'只更改最后的的日期，暂保持原表车台数不变，求和车台数功能以后添加
    Range("a2") = Left(a2, a22) & Year(idate) & "年" & Month(idate) & "月" & Day(idate) & "日"
    ActiveWorkbook.Save
'    ActiveWindow.Close

End Sub

'Sub open_all_files(参数)
'Dim a
'a = Dir("C:\Users\Administrator\Desktop\新建文件夹\*.txt") '将txt结尾的所有文件打开，但是在这里只打开第一个符合的文件，接下来的文件在do循环里依次打开
'Workbooks.Open "C:\Users\Administrator\Desktop\新建文件夹\" + a
'Do '遍历目录下的所有指定格式的文件名
'a = Dir '之前dir()下已经打开了多个文件，这里就不用在写上，表示依次打开符合格式的文件
'If a <> "" Then
'Workbooks.Open "C:\Users\Administrator\Desktop\新建文件夹\" + a '打开每一个符合格式的文件
'Else
'Exit Sub
'End If
'Loop

'End Sub

Sub open_file(FN, key, extension)
    'open_file(FN, key, extension)按路径、文件关键字、文件后缀名打开文件
    '文件夹路径FN，文件关键字key，文件后缀名extension
    
'    MsgBox "FN:" & FN & ",file:" & file
    Dim file
    file = "*" & key & "*." & extension
'    MsgBox file
    If Len(Dir(FN & file)) = 0 Then
'   MsgBox (Dir(FN & file))
        MsgBox "找不到文件：" & vbCrLf & FN & file, vbExclamation, "错误"
'End语句终止整个程序
        End
    Else
    Workbooks.Open filename:=FN & Dir(FN & file)
    End If
    
End Sub
