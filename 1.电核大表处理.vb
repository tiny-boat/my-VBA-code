
Sub 宏1()
'
' 宏1 宏
'
' 快捷键: Ctrl+t
'
    Application.DisplayAlerts = False
    a = "D:\CGSS调查问卷统计\数据质控测试文档\"
    b = Format$(Now, "yyyy-mm-dd")
    
    Sheets("CGSS2017电核样本更新及进度管理表").Select
    Columns("B:B").Select
    Selection.NumberFormatLocal = "0_ "
    ActiveWorkbook.SaveAs a & b & " dhdb.csv", xlCSV
    
    Sheets("废卷情况汇总").Select
    Columns("B:B").Select
    Selection.NumberFormatLocal = "0_ "
    ActiveWorkbook.SaveAs a & b & " fjbh.csv", xlCSV
    
    Sheets("高缺失率问卷").Select
    Columns("F:F").Select
    Selection.NumberFormatLocal = "0_ "
    ActiveWorkbook.SaveAs a & b & " gqslwj.csv", xlCSV
    
    Sheets("电话回收率低访员").Select
    ActiveWorkbook.SaveAs a & b & " hsldfy.csv", xlCSV
    
    ActiveWorkbook.Close SaveChanges:=False
End Sub

	
End Sub


