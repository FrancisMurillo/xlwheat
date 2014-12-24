Attribute VB_Name = "Sandbox"
Public Sub ExportWheatCode()
    Dim ExportModules As Variant, Path As String
    Path = "C:\Users\NOBODY\Desktop\Robot\Workspace\Excel\Wheat\wheat-src"
    ExportModules = Array("Wheat", "WheatLib", "WheatConfig", "TestSuite", "Sandbox")
        
    Debug.Print "Exporting Modules"
    Debug.Print "Run on:" & Now
    Debug.Print "Export path: " & Path
    With ActiveWorkbook.VBProject
        Dim Component As Variant
        For Each Component In .VBComponents
            If WheatLib.InArray(ExportModules, Component.Name) Then
                Component.Export Path & Application.PathSeparator & Component.Name & ".bas"
                Debug.Print "Module " & Component.Name
            End If
        Next
    End With
    Debug.Print "Export Success"
End Sub

Public Sub Sandbox()
    Dim r As String
    
    With ActiveWorkbook.VBProject
        .VBComponents.Import "C:\Users\FVMurillo\Desktop\Robot\Workspace\Excel\excel-sketches\wheat\WheatUtil.bas"
        
    End With
    
End Sub
