Attribute VB_Name = "WheatLib"
'# Wheat Files and Directory names to be used
Public Const SHEET_EXTENSION As String = "shcls"
Public Const FORM_EXTENSION As String = "frm"
Public Const CLASS_EXTENSION As String = "cls"
Public Const MODULE_EXTENSION As String = "bas"

'# Project Directories
Public Const SHEET_DIR As String = "Sheets"
Public Const FORM_DIR As String = "Forms"
Public Const MODULE_DIR As String = "Modules"
Public Const CLASS_DIR As String = "Classes"

'# Imports the wheat repo into the current project
Public Function ImportProject(WheatRepo As String, _
    Optional ShowImportedModules As Boolean = WheatConfig.SHOW_IMPORTED_MODULES, _
    Optional ShowPassedModules As Boolean = WheatConfig.SHOW_PASSED_MODULES, _
    Optional ShowPassedExceptModules As Boolean = WheatConfig.SHOW_PASSED_EXCEPT_MODULES)
    Dim SUBFOLDERS As Variant, EXTENSIONS As Variant
    SUBFOLDERS = Array(SHEET_DIR, FORM_DIR, MODULE_DIR, CLASS_DIR)
    EXTENSIONS = Array(SHEET_EXTENSION, FORM_EXTENSION, CLASS_EXTENSION, MODULE_EXTENSION)
    
    Dim SubDir As Variant, SubPath As String
    Dim Files As Variant, File As Variant, ModulePath As String, ModuleName As String
    Dim Module As VBComponent, Modules As VBComponents, Ext As String, TmpModule As VBComponent, FName As String
    Set Modules = ActiveWorkbook.VBProject.VBComponents
    Dim PassedModuleCol As New Collection, ImportedModuleCol As New Collection
    For Each SubDir In SUBFOLDERS
        SubPath = Join_(WheatRepo, CStr(SubDir))
        Files = ListFiles(AsDirectory(SubPath))
        For Each File In Files
            ModulePath = Join_(SubPath, CStr(File))
            Ext = AsFileExtension(GetFile(ModulePath))
            ModuleName = AsFileName(GetFile(ModulePath))
            If Ext = SHEET_EXTENSION Then
                Set Module = GetModuleFromSheetName(ModuleName)
            Else
                Set Module = GetModule(ModuleName)
            End If
            
            If InArray(EXTENSIONS, Ext) Then
                If InLikeArray(ModuleName, WheatConfig.PassImportModules) And _
                    Not InLikeArray(ModuleName, WheatConfig.PassExceptImportModules) _
                Then
                    PassedModuleCol.Add ModuleName
                Else
                    ImportModule Modules, ModulePath, ModuleName, Ext, Module
                    ImportedModuleCol.Add ModuleName
                End If
            End If
        Next
    Next
    
    ' Output section
    
End Function

'# This centralizes how modules are imported
Private Sub ImportModule(Modules As VBComponents, ModulePath As String, ModuleNameID As String, ModuleExt As String, ModuleSource As VBComponent)
    If ModuleSource Is Nothing Then ' If the module does not exist, import it right away
        If Ext = SHEET_EXTENSION Then ' If it is sheet code, you can't import normally
            ' You have to create the sheet and copy the code
            Dim NewSheet As Worksheet, SheetModule As VBComponent, TempModule As VBComponent
            Set NewSheet = Worksheets.Add
            
            NewSheet.Name = ModuleName
            
            Set SheetModule = GetModule(NewSheet.CodeName)
            Set TmpModule = Modules.Import(ModulePath)
                
            CopyCode TmpModule, SheetModule
            Modules.Remove TmpModule
        Else ' All others import normally
            Modules.Import ModulePath
        End If
    Else
        ' The idea is to remove the module and import it
        ' But this does not work on forms and sheet, copy pasting the code is the way to go
        Select Case ModuleSource.Type
            Case vbext_ct_Document
                ' We import it first as a class module
                ' Then we copy the codes to the target sheet
                Set TmpModule = Modules.Import(ModulePath)
                
                CopyCode TmpModule, Module
                Modules.Remove TmpModule
            Case Else
                If Module.Name = "ThisWorkbook" Then ' Special case for the active workbook
                    Set TmpModule = Modules.Import(ModulePath)
                
                    CopyCode TmpModule, ModuleSource
                    Modules.Remove TmpModule
                Else
                    Modules.Remove ModuleSource
                    Modules.Import ModulePath
                End If
        End Select
    End If
End Sub

'# This returns the worksheet associated with a sheet using it's codename, not the sheet name
'# ModuleName: The automatic name of the sheet, not necessarily the sheet name
Public Function GetSheetFromModuleName(ModuleName As String) As Worksheet
    Dim Sheet As Worksheet
    For Each Sheet In ActiveWorkbook.Worksheets
        If Sheet.CodeName = ModuleName Then
            Set GetSheetFromModuleName = Sheet
            Exit Function
        End If
    Next
    Set GetSheetFromModuleName = Nothing
End Function

'# This returns the module associated with the sheet
'@ ModuleName: The name of an module
Public Function GetModuleFromSheetName(SheetName As String) As VBComponent
    Dim Module As VBComponent
On Error Resume Next
    Dim Sheet As Worksheet
    Set Sheet = Nothing
    Set Sheet = Worksheets(SheetName)
    
    If Sheet Is Nothing Then
        Set GetModuleFromSheetName = Nothing
    Else
        Set GetModuleFromSheetName = ActiveWorkbook.VBProject.VBComponents(Sheet.CodeName)
    End If
End Function

'# This returns the VBComponent with the same name or nothing
'@ ModuleName: Module name to be called
Public Function GetModule(ModuleName As String) As VBComponent
    Set GetModule = Nothing
On Error Resume Next
    Set GetModule = ActiveWorkbook.VBProject.VBComponents(ModuleName)
End Function

'# This just adds the file separator to the end of the path
'@ Path: Your usual path, no file separator at the end please
Public Function AsDirectory(Path As String) As String
    AsDirectory = Join_(Path, "")
End Function

'# Get file extension of a file
'@ FileName: A filename with a file extension
Public Function AsFileExtension(File As String) As String
    Const WINDOWS_SEPARATOR As String = "."
    If InStr(File, WINDOWS_SEPARATOR) > 0 Then
        Dim EXTENSIONS As Variant ' A file may have multiple extensions, take the last one
        Extension = Split(File, ".")
        AsFileExtension = Extension(UBound(Extension))
    Else
        AsFileExtension = ""
    End If
End Function

'# Get filename of a file without the extension, used in conjunction with GetFile
'@ FileName: A filename with a file extension
Public Function AsFileName(File As String) As String
    Const WINDOWS_SEPARATOR As String = "."
    If InStr(File, WINDOWS_SEPARATOR) > 0 Then
        AsFileName = Split(File, ".")(0)
    Else
        AsFileName = File
    End If
End Function

'# You're friendly neighborhood UNIX ls for files only
'# This returns an array(base 0) of files
'@ Path: A folder path as usual, the path must also end with the path separator
Public Function ListFiles(Path As String) As Variant
    ' Using Dir as a repeating module method is weird, a wrapper is needed
    Dim FileCol As New Collection
    Dim File As String
    File = Dir(Path, vbNormal)
    
    While File <> vbNullString
        FileCol.Add File
        File = Dir
    Wend
    
    If FileCol.Count = 0 Then
        ListFiles = Array()
    Else
        ListFiles = CollectionToArray(FileCol)
    End If
End Function

'# Converts an array into an collection
'# Taken from another project as well, no need to test
'# Reference: http://en.wikibooks.org/wiki/Visual_Basic/Collections
Public Function CollectionToArray(Col As Collection) As Variant
    Dim MyArray As Variant
    ReDim MyArray(Col.Count - 1)
    Index = 0
    For Each Member In Col
      MyArray(Index) = Member
      Index = Index + 1
    Next
    CollectionToArray = MyArray
End Function

'# Get filename of a filepath
'@ Path: A path to a file
Public Function GetFile(Path As String) As String
    GetFile = Dir(Path, vbNormal)
End Function

'# Export whole project
'@ Path: The wheat repo you want it dumped
Public Sub ExportProject(WheatRepo As String, _
            Optional ShowExportedModules As Boolean = WheatConfig.SHOW_EXPORTED_MODULES, _
            Optional ShowIgnoredModules As Boolean = WheatConfig.SHOW_IGNORED_MODULES, _
            Optional ShowIgnoredExceptModules As Boolean = WheatConfig.SHOW_IGNORED_EXCEPT_MODULES)
    Dim Component As VBComponent
    Dim ExportedModuleCol As New Collection, IgnoredModuleCol As New Collection
    Dim Folder As String, Ext As String, FileName As String
    ' Iterate through each object and export to folder depending on what type
    For Each Component In ActiveWorkbook.VBProject.VBComponents
        If InLikeArray(Component.Name, WheatConfig.IgnoreExportModules) And _
           Not InLikeArray(Component.Name, WheatConfig.IgnoreExceptExportModules) _
        Then
            ' Ignore this module
            IgnoredModuleCol.Add Component.Name
        Else
            Ext = ""
            Select Case Component.Type
                Case vbext_ct_StdModule
                    Folder = MODULE_DIR
                    Ext = MODULE_EXTENSION
                    FileName = Component.Name
                Case vbext_ct_ClassModule
                    Folder = CLASS_DIR
                    Ext = CLASS_EXTENSION
                    FileName = Component.Name
                Case vbext_ct_MSForm
                    Folder = FORM_DIR
                    Ext = FORM_EXTENSION
                    FileName = Component.Name
                Case vbext_ct_Document
                    Folder = SHEET_DIR
                    If Component.Name = "ThisWorkbook" Then ' ThisWorkbook is a module under Excel Objects, weird
                        FileName = Component.Name
                        Ext = CLASS_EXTENSION
                    Else
                        FileName = GetSheetFromModuleName(Component.Name).Name ' Special mention, export by sheet name
                        Ext = SHEET_EXTENSION
                    End If
            End Select
            Component.Export Join_(Join_(WheatRepo, Folder), FileName & "." & Ext)
            ExportedModuleCol.Add Component.Name
        End If
    Next
    
    ' Output section
    If ShowIgnoredModules Then
        Debug.Print "Ignored Modules"
        Debug.Print "---------------"
        
        If IgnoredModuleCol.Count = 0 Then
            Debug.Print "No modules ignored"
        Else
            Dim IgnoredModule As Variant
            For Each IgnoredModule In IgnoredModuleCol
                Debug.Print "* " & IgnoredModule
            Next
            Debug.Print "Total: " & IgnoredModuleCol.Count
        End If
        
        Debug.Print ""
    End If
    
    If ShowExportedModules Then
        Debug.Print "Exported Modules"
        Debug.Print "---------------"
        
        If ExportedModuleCol.Count = 0 Then
            Debug.Print "No modules exported"
        Else
            Dim ExportedModule As Variant, ExceptedModuleCount As Long
            ExceptedModuleCount = 0
            For Each ExportedModule In ExportedModuleCol
                ' Check if it an excepted module
                If ShowIgnoredExceptModules And _
                    InLikeArray(CStr(ExportedModule), WheatConfig.IgnoreExportModules) And _
                    InLikeArray(CStr(ExportedModule), WheatConfig.IgnoreExceptExportModules) _
                Then
                    Debug.Print "! " & ExportedModule
                    ExceptedModuleCount = ExceptedModuleCount + 1
                Else
                    Debug.Print "+ " & ExportedModule
                End If
                
            Next
            Debug.Print "Total: " & ExportedModuleCol.Count & _
                IIf(ExceptedModuleCount > 0, " / Excepted: " & ExceptedModuleCount, "")
        End If
        Debug.Print ""
    End If
End Sub

'# Checks if an element is like in an array
Public Function InLikeArray(Similar As String, Arr As Variant) As Boolean
    InLikeArray = False
    If UBound(Arr) = -1 Then Exit Function
    
    Dim Item As Variant
    For Each Item In Arr
        InLikeArray = CStr(Similar) Like CStr(Item)
        If InLikeArray Then Exit Function
    Next
End Function

'# Checks if an element is in an array, uses Filter()
Public Function InArray(Arr As Variant, Value As String) As Boolean
    InArray = (UBound(Filter(Arr, Value)) > -1)
End Function

'# Copies code from one module to the other
'@ SrcMod: The module source
'@ DstMod: The module to be inserted
Public Function CopyCode(SrcMod As VBComponent, DstMod As VBComponent)
    DeleteCode DstMod
    SetCode DstMod, GetCode(SrcMod)
End Function


'# Deletes the code of an module
'# Used to fill it back up
'@ Component: Any component
Public Sub DeleteCode(Component As VBComponent)
    With Component.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
    End With
End Sub
'# Set the code of an component
'# Works along side GetCode
'@ Component: A sheet component
Public Function SetCode(Component As VBComponent, Code As String) As String
    DeleteCode Component
    With Component.CodeModule
        .AddFromString Code
    End With
End Function
'# Get the code of an component
'# Must be tested visually
'@ Component: A sheet component to be precise but any will do
Public Function GetCode(Component As VBComponent) As String
    With Component.CodeModule
        If .CountOfLines > 0 Then
            GetCode = .Lines(1, .CountOfLines)
        Else
            GetCode = ""
        End If
    End With
End Function

'# Create the basic structure of a wheat repo
'@ Path: A folder where you want to create a wheat repo
Public Sub SetupWheatRepo(Path As String)
    Debug.Print "Setting repository at " & Path
    
    Debug.Print "Creating /" & SHEET_DIR
    SafeMkDir (Join_(Path, SHEET_DIR))
    
    Debug.Print "Creating /" & FORM_DIR
    SafeMkDir (Join_(Path, FORM_DIR))
    
    Debug.Print "Creating /" & MODULE_DIR
    SafeMkDir (Join_(Path, MODULE_DIR))
    
    Debug.Print "Creating /" & CLASS_DIR
    SafeMkDir (Join_(Path, CLASS_DIR))
End Sub

'# Creates a directory without the exception
'@ Path: The usual branch
Public Sub SafeMkDir(Path As String)
On Error Resume Next
    MkDir Path
End Sub

'# Checks if the path is a valid wheat repo dir
'# The check is only if it contains the four directories(Sheets, Forms, Modules, Classes)
'@ Path: An absolute path to an repo
Public Function IsWheatRepo(Path As String) As Boolean
    IsWheatRepo = IsDir(Path) And _
                    IsDir(Path & Application.PathSeparator & SHEET_DIR) And _
                    IsDir(Path & Application.PathSeparator & FORM_DIR) And _
                    IsDir(Path & Application.PathSeparator & MODULE_DIR) And _
                    IsDir(Path & Application.PathSeparator & CLASS_DIR)
End Function

'# Returns an absolute path given a path, be it relative or absolute.
'# Relative by the way to the current dir.
'# If the combined path does not exist, this returns an empty string
'@ Path: Any path
Public Function AsAbsolutePath(Path As String, Optional RelativePath As String = ".") As String
    If RelativePath = "." Then RelativePath = Application.ActiveWorkbook.Path
    If IsDir(Path) Then
        AsAbsolutePath = Path
    Else
        Dim NPath As String
        NPath = Join_(RelativePath, Path)
        If IsDir(NPath) Then
            AsAbsolutePath = NPath
        Else
            AsAbsolutePath = ""
        End If
    End If
End Function



'# Checks if an path is an directory relative to a path
'@ Path: A path to an directory, should be relative(not starting with a path separator)
'@ RelativePath: The base path of the directory, if not stated the current directory is used
Public Function IsRelativeDir(Path As String, Optional RelativePath As String = ".") As Boolean
    If RelativePath = "." Then RelativePath = Application.ActiveWorkbook.Path
    IsRelativeDir = IsDir(Join_(RelativePath, Path))
End Function

'# Joins a patha as in Python's os.path.join
'@ BasePath: The base path to be joined
'@ ExtPath: The path that will join base path
Public Function Join_(BasePath As String, ExtPath As String) As String
    Join_ = BasePath & Application.PathSeparator & ExtPath
End Function


'# Checks if an path is an directory
'@ Path: A path to an directory
Public Function IsDir(Path As String) As Boolean
    IsDir = Dir(Path, vbDirectory) <> vbNullString
End Function

'# Converts a collection to an array
'# Taken from Vase
Public Function ToArray(Col As Collection) As Variant
    If Col Is Nothing Then
        ToArray = Array()
        Exit Function
    End If
    
    If Col.Count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    Dim Arr As Variant, Item As Variant, Index As Integer
    Arr = Array()
    ReDim Arr(0 To Col.Count - 1)
    Index = 0
    For Each Item In Col
        Arr(Index) = Item
        Index = Index + 1
    Next
    
    ToArray = Arr
End Function

