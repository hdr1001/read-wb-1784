Attribute VB_Name = "Z_GetMostRecentVersions"
Option Explicit

'Check if a worksheet exists
Private Function Z_WsExists(sWsName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sDummy As String
    sDummy = ThisWorkbook.Worksheets(sWsName).Name
    
    Z_WsExists = True
    Exit Function
    
ErrHandler:
    Z_WsExists = False
End Function

'Check if VBA module exists
Private Function bComponentExists(sComponent As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sDummy As String
    sDummy = ThisWorkbook.VBProject.VBComponents(sComponent).Name
    
    bComponentExists = True
    Exit Function
    
ErrHandler:
    bComponentExists = False
End Function

'Use functionality WinHttpRequest object to access the code on the Internet
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

'Write a text file from a string
Private Sub WriteContent2TextFile(sFile As String, sContent As String)
    Dim FileNum As Integer
    FileNum = FreeFile
    
    Open sFile For Output As #FileNum

    Print #FileNum, sContent

    Close #FileNum
End Sub

'Create a workbook for reading the WB1784 based on the latest code and reference tables
Public Sub ReadCodeAndRefTables()
    On Error GoTo ErrHandler
    
    'Some global settings
    Dim sUrlGoogleCode As String: sUrlGoogleCode = "http://read-wb-1784.googlecode.com/svn/trunk/"

    Application.ScreenUpdating = False 'No screen updates at this stage
    Application.DisplayAlerts = False 'No alerts

    'Insert the Visual Basic modules
    Dim collBasFiles As New Collection
    Dim sFileName As Variant, sFileContents As String, sModuleName As String
    
    'Create a collection of the modules to be inserted
    collBasFiles.Add "A_Globals.bas", "Globals"
    collBasFiles.Add "B_EventHandlers.bas", "EventHandlers"
    collBasFiles.Add "C_PublicFunctions.bas", "PublicFunctions"
    
    For Each sFileName In collBasFiles
        'Get the latest version of the code and create a local copy of the file
        sFileContents = sDownloadTextFile(sUrlGoogleCode & sFileName)
        WriteContent2TextFile ThisWorkbook.Path & "\" & sFileName, sFileContents
        
        'Remove, if necessary, possibly outdated code from the workbook
        sModuleName = Left(sFileName, Len(sFileName) - 4)
        If bComponentExists(sModuleName) Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(sModuleName)
            Debug.Print "Removed module " & sModuleName
        End If
        
        'Import the most recent version of the code
        ThisWorkbook.VBProject.VBComponents.Import Filename:=ThisWorkbook.Path & "\" & sFileName
        Debug.Print "Imported module " & sModuleName
    Next

ErrHandler:
    Application.DisplayAlerts = True 'Enable application alerts
    Application.ScreenUpdating = True 'Update the screen

    'Give the user feedback about what went wrong
    If Err.Number <> 0 Then MsgBox "Error (" & Err.Number & ") occured. " & Err.Description
End Sub
