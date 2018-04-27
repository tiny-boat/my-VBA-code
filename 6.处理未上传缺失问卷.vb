Sub 处理未上传缺失问卷()
'
' 处理未上传缺失问卷 宏
'
' 快捷键: Ctrl+j
'
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
	
    Columns("F:F").Select
    Selection.NumberFormatLocal = "0_ "
	
	Columns("G:G").Select
    Columns("G:G").EntireColumn.AutoFit
	
    Columns("I:I").Select
    Columns("I:I").EntireColumn.AutoFit
	
    Columns("K:O").Select
    Selection.NumberFormatLocal = "0.00%"
	
    Columns("Z:AA").Select
    Selection.NumberFormatLocal = "0_ "
	
	maxrow = ActiveSheet.UsedRange.Rows.Count 
	Range("A2:AA" & maxrow).Select
    Selection.Copy
End Sub
