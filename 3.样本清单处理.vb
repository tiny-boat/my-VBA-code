Sub SaveToCSV()
    Dim fDir As String
    Dim fPath As String
    Dim sPath As String
    fPath = "D:\CGSS调查问卷统计\样本清单信息\" '输入路径
    sPath = "D:\CGSS调查问卷统计\样本清单信息\" '输出路径
    fDir = Dir(fPath) '路径下的第一个文件名
    Do While (fDir <> "")
        Workbooks.Open (fPath & fDir) '打开文件
        Sheets("正选样本").Select '定位到需要的sheet
        ActiveWorkbook.SaveAs sPath & Mid(ActiveWorkbook.Name, 1, 7) & ".csv", xlCSV '另存为与源文件同名的csv文件
        ActiveWorkbook.Close SaveChanges:=False '关闭文件，不保存
        fDir = Dir '路径下的下一个文件名
    Loop
End Sub