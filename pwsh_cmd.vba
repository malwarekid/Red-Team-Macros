Sub Auto_Open()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "powershell.exe -NoP -NonI -ExecutionPolicy Bypass -Command ""Get-Process | Out-File -FilePath C:\Temp\processes.txt"""
End Sub
