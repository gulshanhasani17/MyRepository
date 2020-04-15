Set oServ = GetObject("winmgmts:")
Set cProc = oServ.ExecQuery("Select * from Win32_Process")

For Each oProc In cProc

    'Rename EXCEL.EXE in the line below with the process that you need to Terminate. 
    'NOTE: It is 'case sensitive

    If oProc.Name = "chromedriver.exe" Then
      'MsgBox "KILL"   ' used to display a message for testing pur
      oProc.Terminate()
    End If
	 
	 
Next

