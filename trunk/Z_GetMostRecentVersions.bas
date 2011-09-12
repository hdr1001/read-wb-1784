Attribute VB_Name = "Z_GetMostRecentVersions"
Option Explicit

Private Function sDownloadTextFile(url As String) As String
    Dim oHTTP As WinHttp.WinHttpRequest
    Set oHTTP = New WinHttp.WinHttpRequest
    oHTTP.Open Method:="GET", url:=url, async:=False
    oHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHTTP.setRequestHeader "Content-Type", "multipart/form-data; "
    oHTTP.Option(WinHttpRequestOption_EnableRedirects) = True
    oHTTP.send

    If Not oHTTP.waitForResponse() Then Err.Raise Number:=1025, Description:="Error while waiting for response"
    
    sDownloadTextFile = oHTTP.responseText
End Function

Private Sub WriteContent2TextFile(sFile As String, sContent As String)
    Dim FileNum As Integer
    FileNum = FreeFile 'Get file number available for open
    Open sFile For Output As #FileNum 'Creates the file in case it doesn't exist else overwrites
    Print #FileNum, sContent 'Write the content
    Close #FileNum 'Now close the file
End Sub

Private Function bComponentExists(sComponent As String) As Boolean
    On Error GoTo ErrHandler
    Dim s As String: s = ThisWorkbook.VBProject.VBComponents(sComponent).Name
    bComponentExists = True
    Exit Function
    
ErrHandler:
    bComponentExists = False
End Function

Public Sub ReadCodeAndRefTables()
    On Error GoTo ErrHandler
    
    Dim sUrlGoogleCode As String: sUrlGoogleCode = "http://read-wb-1784.googlecode.com/svn/trunk/"
    
    Dim sGlobalsBas As String: sGlobalsBas = "A_Globals.bas"
    Dim sEventHandlersBas As String: sEventHandlersBas = "B_EventHandlers.bas"
    Dim sPublicFuncsBas As String: sPublicFuncsBas = "C_PublicFunctions.bas"
    
    Dim str As String
    
    'Retrieve the latest version of A_Globals.bas
    str = sDownloadTextFile(sUrlGoogleCode & sGlobalsBas)
    WriteContent2TextFile ThisWorkbook.Path & "\" & sGlobalsBas, str
    
    'Retrieve the latest version of B_EventHandlers.bas
    str = sDownloadTextFile(sUrlGoogleCode & sEventHandlersBas)
    WriteContent2TextFile ThisWorkbook.Path & "\" & sEventHandlersBas, str
    
    'Retrieve the latest version of C_PublicFunctions.bas
    str = sDownloadTextFile(sUrlGoogleCode & sPublicFuncsBas)
    WriteContent2TextFile ThisWorkbook.Path & "\" & sPublicFuncsBas, str
    
    'For the following code to work Excel must have trusted access to VBProject!
    'Go to the Tools menu, first choose Options then Security
    'Check the "Trust Access To VBProject" on the Trusted Sources tab
    
    'Import A_Globals.bas but first remove the existing version of this module if needed
    str = Left(sGlobalsBas, Len(sGlobalsBas) - 4)
    If bComponentExists(str) Then
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(str)
    End If
    ThisWorkbook.VBProject.VBComponents.Import Filename:=ThisWorkbook.Path & "\" & sGlobalsBas
    
    'Import B_EventHandlers.bas but first remove the existing version of this module if needed
    str = Left(sEventHandlersBas, Len(sEventHandlersBas) - 4)
    If bComponentExists(str) Then
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(str)
    End If
    ThisWorkbook.VBProject.VBComponents.Import Filename:=ThisWorkbook.Path & "\" & sEventHandlersBas
    
    'Import sPublicFuncsBas.bas but first remove the existing version of this module if needed
    str = Left(sPublicFuncsBas, Len(sPublicFuncsBas) - 4)
    If bComponentExists(str) Then
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(str)
    End If
    ThisWorkbook.VBProject.VBComponents.Import Filename:=ThisWorkbook.Path & "\" & sPublicFuncsBas
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error (" & Err.Number & ") occured. " & Err.Description
End Sub
