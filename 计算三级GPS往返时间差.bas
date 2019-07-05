Sub 计算三级GPS往返时间差()
'
' 计算三级GPS往返时间差 宏
'
' 快捷键: Ctrl+q
'
    
    
    Range("S15").Select
    ActiveCell.FormulaR1C1 = "=RC[-15]-RC[-16]"
    Range("S15").Select
    Selection.AutoFill Destination:=Range("S15:S20"), Type:=xlFillDefault
    Range("S15:S20").Select
    Selection.AutoFill Destination:=Range("S15:AA20"), Type:=xlFillDefault
    Range("S15:AA20").Select
    Selection.AutoFill Destination:=Range("S15:AC20"), Type:=xlFillDefault
    Range("S15:AC20").Select
    Selection.NumberFormatLocal = "h:mm;@"
    Range("T15:T20,V15:V20,X15:X20,Z15:Z20,AB15:AB20").Select
    Range("AB15").Activate
    Selection.ClearContents
    
    Range("s15:s20,u15:u20,w15:w20,y15:y20,Aa15:Aa20,ac15:ac20").Select
    
End Sub
