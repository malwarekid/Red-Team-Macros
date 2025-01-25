Sub AutoOpen()
    Dim objShell As Object
    Dim command As String
    command = "powershell.exe -WindowStyle Hidden -NoP -NonI -ExecutionPolicy Bypass -Command ""Get-Process | Out-File -FilePath C:\Temp\processes.txt"""
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run command, 0, False
End Sub

