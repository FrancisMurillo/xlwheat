Attribute VB_Name = "TestWheatLib"
Public Sub TestWheatLibAsAbsolutePath()
On Error Resume Next
    ' Setup
    Const RANDOM As String = "Random Dir"
    MkDir RANDOM

    ' Checking
    Dim Path As String
    Path = WheatLib.AsAbsolutePath(RANDOM)
    VaseAssert.AssertTrue Path <> vbNullString
    
    Path = WheatLib.AsAbsolutePath(Path & 1)
    VaseAssert.AssertEqual Path, vbNullString
    
    RmDir RANDOM
End Sub

