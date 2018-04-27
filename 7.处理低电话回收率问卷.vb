Sub 处理低电话回收率问卷()
'
' 处理低电话回收率问卷 宏
'
' 快捷键: Ctrl+j
'
  
    Columns("L:M").Select
    Selection.NumberFormatLocal = "0.00%"
	
    Columns("C").Select
    Selection.NumberFormatLocal = "0_ "
	
	maxrow = ActiveSheet.UsedRange.Rows.Count 
	Range("A2:M" & maxrow).Select
    Selection.Copy
End Sub
