Attribute VB_Name = "Vase"
'=======================
'--- User Function   ---
'=======================

'# Like Nose, this starts test discovery and runs each test found by it
Public Sub RunTests()
On Error GoTo ErrHandler
    VaseLib.ClearScreen
    
    Debug.Print "Vase Test Framework"
    Debug.Print "Don't break the vase."
    Debug.Print "======================="
    
    VaseLib.RunVaseSuite ActiveWorkbook, Verbose:=True ' The output result is printed out, so no need to capture the output
    
    Debug.Print "Vase was filled"
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print _
            "Whoops! There was error running Vase. Check if you put the water in the vase correctly."
    End If
    Err.Clear
End Sub
