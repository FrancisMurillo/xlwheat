Attribute VB_Name = "ChipLib"
Public gVerbose As Boolean ' Sets the debug print option

'===========================
'Internal Functions
'===========================

Public Sub AttachChip(ChipBookPath As String, _
        Optional Verbose As Boolean = True)
    gVerbose = Verbose ' Set global output flag
    Dim ChipBook As Workbook, MyBook As Workbook
    Dim ProjectReferences As Variant
    ProjectReferences = ChipLib.ListProjectReferences
    
On Error GoTo Cleanup
    WriteLine "Opening Chip Book"
    Set MyBook = ActiveWorkbook
    Set ChipBook = Workbooks.Open(ChipBookPath)

    WriteLine "Reading ChipInfo"
    CopyModule "ChipInfo", ChipBook, MyBook
    
    ChipReadInfo.ClearInfo
    Application.Run MyBook.Name & "!ChipInfo.WriteInfo"
    
    WriteLine "Checking depedencies"
    Dim ChipReference As Variant, ReferencesSatisfied As Boolean
    ReferencesSatisfied = True
    For Each ChipReference In ChipReadInfo.References
        If Not InLikeArr(ChipReference, ProjectReferences) Then
            ReferencesSatisfied = False
            Exit Sub
        End If
    Next
    
    If Not ReferencesSatisfied Then
        WriteLine "Chip references not satisfied. Include these references in the project and try again: "
        For Each ChipReferene In ChipReadInfo.References
            WriteLine "* " & ChipReference
        Next
    End If
    
    WriteLine "Installing modules:"
    Dim ChipModule As Variant
    For Each ChipModule In ChipReadInfo.Modules
        WriteLine "* " & ChipModule
        CopyModule CStr(ChipModule), ChipBook, MyBook
    Next
Cleanup:
On Error Resume Next
    If Err.Number <> 0 Then
        WriteLine _
            "Whoops! An error occured with the Chip book." & _
            "Check if the book is a certified Chip book and there are no residual modules installed."
    End If
    Err.Clear

    WriteLine "Closing Chip book"
    DoEvents
    ChipBook.Close SaveChanges:=False
    DoEvents
End Sub


'===========================
' Helper Functions
' Taken from ChipLib, now more comprehensive
'===========================

'# Checks if a pattern matches any one in an array.
'# Sort of InArr with a Like twist
Public Function InLikeArr(Pattern As Variant, Arr As Variant) As Boolean
    InLikeArr = False
    Dim Elem As Variant
    For Each Elem In Arr
        InLikeArr = (Elem Like Pattern)
        If InLikeArr Then Exit Function
    Next
End Function

'# Copies one module from one book to the other
Public Sub CopyModule(ModuleName As String, SourceBook As Workbook, TargetBook As Workbook, _
            Optional ShouldOverwrite As Boolean = True)
    If Not HasModule(ModuleName, SourceBook) Then
        Err.Raise 10001, _
            Description:=SourceBook.Name & " does not have the module " & ModuleName
    End If
    
    Dim Module As VBComponent, ModulePath As String
    Set Module = SourceBook.VBProject.VBComponents(ModuleName)
    ModulePath = CreateUniqueModulePath()
    
    Module.Export ModulePath
    
    If Not ShouldOverwrite And HasModule(ModuleName, TargetBook) Then
        Err.Raise 10002, _
            Description:=TargetBook.Name & " already has the module " & ModuleName
    Else
        DeleteModule ModuleName, TargetBook
    End If
    
    TargetBook.VBProject.VBComponents.Import ModulePath
    DeleteFile ModulePath
End Sub

'# Creates a pseudo unique path for the exporting modules
Public Function CreateUniqueModulePath() As String
    CreateUniqueModulePath = "~" & Format(Now(), "yyyymmddhhmmss") & "tmp"
End Function

'# Writes to the standard output when the verbose option is set
Public Sub WriteLine(LineMsg As String)
    If gVerbose Then Debug.Print LineMsg
End Sub


'# Clears the intermediate screen
Public Sub ClearScreen()
    Application.SendKeys "^g ^a {DEL}"
    DoEvents
End Sub

'# Removes a module whether it exists or not
'# Used in making sure there are no duplicate modules
Public Sub DeleteModule(ModuleName As String, Book As Workbook)
On Error Resume Next
    Dim CurProj As VBProject, Module As VBComponent
    Set CurProj = Book.VBProject
    Set Module = CurProj.VBComponents(ModuleName)
    CurProj.VBComponents.Remove Module
    DoEvents
    Err.Clear
End Sub

'# Checks if an module exists
Public Function HasModule(ModuleName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasModule = False
    HasModule = Not Book.VBProject.VBComponents(ModuleName) Is Nothing  ' This fails if the module does not exists thus defaulting to False
    Err.Clear
End Function

'# Lists the modules of an workbook
'# Primarily used to get all Chip modules
'@ Return: An array of VB Components
Public Function ListWorkbookModuleObjects(Book As Workbook) As Variant
    Dim Comp As VBComponent, Modules As Variant, Index As Long
    Modules = Array()
    ReDim Modules(0 To Book.VBProject.VBComponents.Count - 1)
    For Each Comp In Book.VBProject.VBComponents
        Set Modules(Index) = Comp
        Index = Index + 1
    Next
    ListWorkbookModuleObjects = Modules
End Function

'# This browses a file using the Open File Dialog
'@ Return: The absolute path of the selected file, an empty string if none was selected
Public Function BrowseFile() As String
    BrowseFile = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
        FileFilter:="Excel Macro Enabled Files *.xlsm (*.xlsm),")
    BrowseFile = IIf(BrowseFile = "False", "", BrowseFile) ' This is to normalize the result
End Function

'# Checks if an file exists
Public Function DoesFileExists(FilePath As String) As Boolean
    With New FileSystemObject
        DoesFileExists = .FileExists(FilePath)
    End With
End Function

'# This downloads a file from the internet using the HTTP GET method
'@ Return: The absolute path of the downloaded file, if path was not provided else the path itself
Public Function DownloadFile(URL As String, Optional Path As String = "")
On Error Resume Next
    If Path = "" Then ' Create pseudo unique path
        Path = ActiveWorkbook.Path & Application.PathSeparator & "~" & Format(Now(), "yyyymmddhhmmss")
    End If

    Dim FileNum As Long
    Dim FileData() As Byte
    Dim MyFile As String
    Dim WHTTP As Object
    
    Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
    If Err.Number <> 0 Then
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
    End If
    Err.Clear
    
    WHTTP.Open "GET", URL, False
    WHTTP.Send
    FileData = WHTTP.responseBody
    Set WHTTP = Nothing
    
    FileNum = FreeFile
    Open Path For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    
    If Err.Number <> 0 Then Path = ""
    Err.Clear
    
    DownloadFile = Path
End Function

'# Deletes a file forcibly, it does not check whether it is a folder or the path does not exists
'# This is used to delete a temp file whether it still exists or not
Public Sub DeleteFile(FilePath As String)
    With New FileSystemObject
        If .FileExists(FilePath) Then
            .DeleteFile FilePath
        End If
    End With
End Sub

'# This returns an string array of the references used in this VBA Project
'# The strings are the name of the references, not the filename or path
'@ Return: A zero-based array of strings
Public Function ListProjectReferences() As Variant
    Dim VBProj As VBIDE.VBProject
    Set VBProj = Application.VBE.ActiveVBProject
    
    Dim ReferenceLength As Integer, Index As Long
    Dim References As Variant
    
    ReferenceLength = VBProj.References.Count
    If ReferenceLength = 0 Then
        ListProjectReferences = Array()
        Exit Function
    End If
    
    References = Array()
    ReDim References(1 To ReferenceLength)
    For Index = 1 To VBProj.References.Count
        With VBProj.References.Item(Index)
            References(Index) = .Description
        End With
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferences = References
End Function

'# This returns an array of project references, the objects themselves for use
'# This is used for setting up the test workbook to have the correct references
'@ Return: A zero-based array of references
Public Function ListProjectReferenceObjects() As Variant
    Dim VBProj As VBIDE.VBProject
    Set VBProj = Application.VBE.ActiveVBProject
    
    Dim ReferenceLength As Integer, Index As Long
    Dim References As Variant
    
    ReferenceLength = VBProj.References.Count
    If ReferenceLength = 0 Then
        ListProjectReferences = Array()
        Exit Function
    End If
    
    References = Array()
    ReDim References(1 To ReferenceLength)
    For Index = 1 To VBProj.References.Count
        Set References(Index) = VBProj.References.Item(Index)
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferenceObjects = References
End Function

'# Checks if the refrence exists for a workbook given its name
Public Function HasReference(ReferenceName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasReference = False
    HasReference = Not Book.VBProject.References(ReferenceName) Is Nothing
    Err.Clear
End Function


