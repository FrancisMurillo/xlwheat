Attribute VB_Name = "Wheat"
'# ----- WHEAT -----
'# Description: An pseudo CVS for Excel since there is none. With two projects in Excel without an CVS,
'#              there must be at least a better way.
'# Author: FVMurillo
'# Copyright: There is none

'# ----- SETUP -----
'# Just import the files Wheat.bas, WheatLib.bas and WheatUtil.bas into your project along with
'# 'Microsoft Visual Basic for Application Extensibility' in your reference(Tools > References)
'# and your good to go.

'# ----- USAGE -----
'# The only module you should care about is Wheat because all of the commands are declared there.
'# The module acts tries to emulate a command line. For example 'git init' is 'Wheat.Setup' and
'# 'git commit' is 'Wheat.Export'. The preferred use is via the immediate window(Ctrl+G) and
'# and executing the Wheat module methods from there with some help from the code completion.
'# You could also use F5 to run the commands but that might be too impractical to scroll.

Public Sub Setup( _
        Optional ProjectRepo As String = WheatConfig.PROJECT_REPO, _
        Optional Verbose As Boolean = True)
    WheatConfig.InitializeVariables
    
    Debug.Print "Wheat Setup"
    Debug.Print "==================="
    
    ' Path error checking
    Dim AbsProjectRepo As String
    AbsProjectRepo = AsAbsolutePath(ProjectRepo)
    If AbsProjectRepo = vbNullString Then
        AbsProjectRepo = WheatLib.Join_(Application.ActiveWorkbook.Path, ProjectRepo)
        WheatLib.SafeMkDir AbsProjectRepo
    End If
    If WheatLib.IsWheatRepo(AbsProjectRepo) Then
        Debug.Print ProjectRepo & " is already an wheat repo."
        Exit Sub
    End If
    
    ' Create the actual repo
    SetupWheatRepo AbsProjectRepo
    Debug.Print "Repo created."
End Sub

Public Sub Export(Optional ProjectRepo As String = WheatConfig.PROJECT_REPO)
    WheatConfig.InitializeVariables

    Debug.Print "Wheat Export"
    Debug.Print "==================="

    ' Path error checking
    Dim AbsProjectRepo As String
    AbsProjectRepo = AsAbsolutePath(ProjectRepo)
    If AbsProjectRepo = vbNullString Then
        Debug.Print "Wheat Export error: " & ProjectRepo & " does not exist."
        Debug.Print "Run Wheat.Init first"
        Exit Sub
    End If
    
    ' Repo checking
    If Not WheatLib.IsWheatRepo(AbsProjectRepo) Then
        Debug.Print "Wheat Export error: " & ProjectRepo & " is not a wheat repo"
        Exit Sub
    End If
    
    WheatLib.ExportProject AbsProjectRepo
    Debug.Print "Modules exported"
End Sub

Public Sub Import(Optional ProjectRepo As String = WheatConfig.PROJECT_REPO)
    WheatConfig.InitializeVariables
    
    Debug.Print "Wheat Import"
    Debug.Print "==================="
    
    ' Path error checking
    Dim AbsProjectRepo As String
    AbsProjectRepo = AsAbsolutePath(ProjectRepo)
    If AbsProjectRepo = vbNullString Then
        Debug.Print "Wheat Import error: " & ProjectRepo & " does not exist."
        Debug.Print "Run Wheat.Init first"
        Exit Sub
    End If
    
    ' Repo checking
    If Not WheatLib.IsWheatRepo(AbsProjectRepo) Then
        Debug.Print "Wheat Import error: " & ProjectRepo & " is not a wheat repo"
        Exit Sub
    End If
    
    WheatLib.ImportProject (AbsProjectRepo)
    Debug.Print "Modules imported"
End Sub

Public Sub Changes(Optional ProjectRepo As String)
    WheatConfig.InitializeVariables
End Sub


