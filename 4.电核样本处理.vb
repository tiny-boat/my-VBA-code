Sub 宏3()
'
' 宏3 宏
'
' 快捷键: Ctrl+y
'
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").Select
    Selection.NumberFormatLocal = "0_ "
    Columns("C:C").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("P:P").Select
    Selection.NumberFormatLocal = "0_ "
    Columns("CC:CC").Select
    Selection.NumberFormatLocal = "0_ "
	
	maxrow = ActiveSheet.UsedRange.Rows.Count 
	Range("A2:CF" & maxrow).Select
    Selection.Copy

End Sub
