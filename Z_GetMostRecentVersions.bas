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
    oHTTP.setRequestHeader "Content-Type", "text/plain;charset=UTF-8"
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

'Return an array of array's to import all the columns in a text file as text
Private Function vFldInfo(lNumCols As Long) As Variant
    Dim arr() As Variant, i As Integer
    ReDim arr(lNumCols - 1)
    
    For i = 0 To lNumCols - 1
        arr(i) = Array(i + 1, 2)
    Next
    
    vFldInfo = arr
End Function

'Create a workbook for reading the WB1784 based on the latest code and reference tables
Public Sub ReadCodeAndRefTables()
    On Error GoTo ErrHandler
    
    'Some global settings
    Dim sUrlGoogleCode As String: sUrlGoogleCode = "http://read-wb-1784.googlecode.com/svn/trunk/"

    Application.ScreenUpdating = False 'No screen updates at this stage
    Application.DisplayAlerts = False 'No alerts

    'Insert the Visual Basic modules
    Dim collBasFiles As New Collection
    Dim vFileName As Variant, sFileContents As String, sModuleName As String
    
    'Create a collection of the modules to be inserted
    collBasFiles.Add "A_Globals.bas", "Globals"
    collBasFiles.Add "B_EventHandlers.bas", "EventHandlers"
    collBasFiles.Add "C_PublicFunctions.bas", "PublicFunctions"
    
    For Each vFileName In collBasFiles
        'Get the latest version of the code and create a local copy of the file
        sFileContents = sDownloadTextFile(sUrlGoogleCode & vFileName)
        WriteContent2TextFile ThisWorkbook.Path & "\" & vFileName, sFileContents
        
        'Remove, if necessary, the possibly outdated code from the workbook
        sModuleName = Left(vFileName, Len(vFileName) - 4)
        If bComponentExists(sModuleName) Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(sModuleName)
            Debug.Print "Removed module " & sModuleName
        End If
        
        'Import the most recent version of the code
        ThisWorkbook.VBProject.VBComponents.Import Filename:=ThisWorkbook.Path & "\" & vFileName
        Debug.Print "Imported module " & sModuleName
    Next

    'Insert the reference tables
    Const FILE_NAME = 0
    Const WS_NAME = 1
    Const NUM_COL = 2
    Const RANGE_NAME = 3
    Dim collRefTables As New Collection
    Dim arr As Variant, sFileName As String, sWsName As String, FldInfoLayout1784 As Variant
    
    'Create a collection of reference tables to imported. Also import layout specification of 1784
    collRefTables.Add Array("Subsidiary code.csv", "Subsidiary code (935 - 935)", 2, "Subsidiary_Codes"), "SubsCode"
    collRefTables.Add Array("Status code.csv", "Status code (934 - 934)", 2, "Status_Codes"), "StatusCode"
    collRefTables.Add Array("Legal status.csv", "Legal status (930 - 932)", 2, "LegalStatus_Codes"), "LegalStatusCode"
    collRefTables.Add Array("Imp-Exp code.csv", "Imp.-Exp. code (929 - 929)", 2, "ImportExport_Codes"), "ImpExpCode"
    collRefTables.Add Array("Currency code.csv", "Currency code (908 - 911)", 7, "Curr_Codes"), "CurrCode"
    collRefTables.Add Array("Activity ind.csv", "Activity ind. (867 - 869)", 2, "LocalAct_Codes"), "ActCode"
    collRefTables.Add Array("SIC87.csv", "SIC87 (835 - 838)", 6, "SIC87_Codes"), "Sic87Code"
    collRefTables.Add Array("National ID code.csv", "National ID code (617 - 621)", 2, "NatID_Codes"), "NatIDCode"
    collRefTables.Add Array("Continent code.csv", "Continent code (429 - 429)", 2, "Cont_Codes"), "ContinentCode"
    collRefTables.Add Array("Country code.csv", "Country code (417 - 419)", 3, "Ctry_Codes"), "CountryCode"
    collRefTables.Add Array("State-Province abbr.csv", "State-Province abbr. (413 -416)", 3, "StateProv_Codes"), "StateProvCode"
    collRefTables.Add Array("Layout1784.csv", "Layout1784", 6, ""), "Layout1784"

    For Each arr In collRefTables 'Import the reference tables
        'Save a local copy of the reference table as a ".txt" file
        sFileName = Left(arr(FILE_NAME), Len(arr(FILE_NAME)) - 3) & "txt"
        sWsName = arr(WS_NAME)
        
        'Get the latest version of the reference table and create a local copy of the file
        sFileContents = sDownloadTextFile(sUrlGoogleCode & arr(FILE_NAME))
        WriteContent2TextFile ThisWorkbook.Path & "\" & sFileName, sFileContents

        'Import the csv into an Excel workbook as a new worksheet
        If Not (sWsName = "Layout1784") Then
            Workbooks.OpenText Filename:=ThisWorkbook.Path & "\" & sFileName, Origin:=65001, StartRow:=1, _
                DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, Comma:=True, FieldInfo:=vFldInfo(CLng(arr(NUM_COL)))
        Else
            FldInfoLayout1784 = Array(Array(1, 2), Array(2, 1), Array(3, 2), Array(4, 2), Array(5, 1), Array(6, 1))
            Workbooks.OpenText Filename:=ThisWorkbook.Path & "\" & sFileName, Origin:=1252, StartRow:=1, _
                DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, Comma:=True, FieldInfo:=FldInfoLayout1784
        End If

        'Remove, if necessary, the possibly outdated reference table from the workbook
        With ActiveSheet 'Just imported reference table
            .Name = sWsName

            'Remove, if necessary, a possibly outdated reference table from the workbook
            If Z_WsExists(sWsName) Then
                ThisWorkbook.Worksheets(sWsName).Delete
                Debug.Print "Removed worksheet " & sWsName
            End If
            
            .Move before:=ThisWorkbook.Sheets(1)
            Debug.Print "Imported table " & sWsName
        End With
    Next
    
    For Each arr In collRefTables 'Add names and hide the worksheet
        If Not (arr(WS_NAME) = "Layout1784") Then
            'Add a name to refer to the reference table
            ThisWorkbook.Names.Add Name:=arr(RANGE_NAME), RefersTo:="='" & arr(WS_NAME) & "'!" & ThisWorkbook.Worksheets(arr(WS_NAME)).Cells(1, 1).CurrentRegion.Address
    
            'Hide the reference table in the workbook
            ThisWorkbook.Worksheets(arr(WS_NAME)).Visible = False
        End If
    Next

ErrHandler:
    Application.DisplayAlerts = True 'Enable application alerts
    Application.ScreenUpdating = True 'Update the screen

    'Give the user feedback about what went wrong
    If Err.Number <> 0 Then MsgBox "Error (" & Err.Number & ") occured. " & Err.Description
End Sub

'Remove the downloaded modules, reference tables & names before save
Private Sub Z_CleanUpWb()
    On Error Resume Next
    
    Application.DisplayAlerts = False
    
    Dim v As Variant
    For Each v In ThisWorkbook.Names
        ThisWorkbook.Names(v).Delete
    Next
    
    Dim collBasFiles As New Collection
    
    collBasFiles.Add "A_Globals", "Globals"
    collBasFiles.Add "B_EventHandlers", "EventHandlers"
    collBasFiles.Add "C_PublicFunctions", "PublicFunctions"
    
    For Each v In collBasFiles
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(v)
    Next
    
    Dim collRefTables As New Collection
    
    collRefTables.Add "Subsidiary code (935 - 935)", "SubsCode"
    collRefTables.Add "Status code (934 - 934)", "StatusCode"
    collRefTables.Add "Legal status (930 - 932)", "LegalStatusCode"
    collRefTables.Add "Imp.-Exp. code (929 - 929)", "ImpExpCode"
    collRefTables.Add "Currency code (908 - 911)", "CurrCode"
    collRefTables.Add "Activity ind. (867 - 869)", "ActCode"
    collRefTables.Add "SIC87 (835 - 838)", "Sic87Code"
    collRefTables.Add "National ID code (617 - 621)", "NatIDCode"
    collRefTables.Add "Continent code (429 - 429)", "ContinentCode"
    collRefTables.Add "Country code (417 - 419)", "CountryCode"
    collRefTables.Add "State-Province abbr. (413 -416)", "StateProvCode"
    collRefTables.Add "Layout1784", "Layout1784"
    
    For Each v In collRefTables
        ThisWorkbook.Worksheets(v).Delete
    Next

    Application.DisplayAlerts = True
End Sub

