Sub 问卷数据另存为csv()
'
' 问卷数据另存为csv 宏
'
' 快捷键: Ctrl+u
'
    Application.DisplayAlerts = False
    a = "D:\CGSS调查问卷统计\数据质控测试文档\"
    b = Format$(Now, "yyyy-mm-dd")
	
	Sheets("2017年中国综合社会调查（CGSS）").Select
	ActiveWorkbook.SaveAs a & b & " wjda.csv", xlCSV
	ActiveWorkbook.Close SaveChanges:=False
End Sub