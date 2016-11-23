Attribute VB_Name = "LogHandler"
Public Sub SaveLogInFile(strLog As String)
    Dim logPath As String

    logPath = App.Path & "\" & "Logs\"
    If Right(logPath, 1) <> "\" Then logPath = logPath & "\"
    
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    
    Open (logPath & Format(Date, "YYYY-MM-DD") & ".log") For Append As #1
    Print #1, CStr(Time) & "> " & strLog
    Close #1
End Sub

Public Sub Log_Info(strLog As String)
    Form1.tbLogInfo.Text = Form1.tbLogInfo.Text + strLog + vbCrLf
    Form1.tbLogInfo.SelStart = Len(Form1.tbLogInfo)

    SaveLogInFile strLog
End Sub

Public Sub Log_Clear()
    Form1.tbLogInfo.Text = ""
End Sub
