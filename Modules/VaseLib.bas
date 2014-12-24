Attribute VB_Name = "VaseLib"
'=======================
'--- Constants       ---
'=======================
Public Const METHOD_HEADER_PATTERN As String = _
    "Public Sub " & VaseConfig.TEST_METHOD_PATTERN

'=======================
'- Internal Functions  -
'=======================

'# Run the test suites with the correct options
'@ Return: A tuple of reporting the test execution
Public Function RunVaseSuite(VaseBook As Workbook, _
        Optional Verbose As Boolean = True, _
        Optional ShowOnlyFailed As Boolean = False) As Variant
    If Verbose Then Debug.Print "" ' Newline
    If Verbose Then Debug.Print "Finding test modules"
    Dim Book As Workbook
    Dim TestMethodCount As Long, TestMethodSuccessCount As Long
    Dim TestModuleCount As Long, TestModuleSuccessCount As Long
    Dim TestMethodLocalPassCount As Long, TestMethodLocalTotalCount As Long
    Dim TestMethodFailedCol As New Collection, TestMethodFailedArr As Variant
    Dim TModules As Variant, TModule As Variant, TMethods As Variant, TMethod As Variant
    Dim Module As VBComponent
    Dim TestResult As Variant
    Set Book = ActiveWorkbook ' Just in case I want to segregate it again
    TModules = FindTestModules(VaseBook)
    
    TestMethodCount = 0
    TestMethodSuccessCount = 0
    TestModuleCount = 0
    TestModuleSuccessCount = 0
    
    For Each TModule In TModules
        If Verbose Then Debug.Print "* " & TModule.Name
        If Verbose Then Debug.Print "=============="
        TestMethodLocalPassCount = 0
        
        Set Module = TModule ' Just type casting it
        TMethods = FindTestMethods(Module)
        For Each TMethod In TMethods
            VaseAssert.InitAssert
            TestResult = RunTestMethod(VaseBook, TModule.Name, CStr(TMethod))
            If TestResult(0) Then
                If Verbose Then Debug.Print vbTab & "+ " & TMethod
                TestMethodLocalPassCount = TestMethodLocalPassCount + 1
            Else
                If Verbose Then Debug.Print vbTab & "- " & TMethod & " >> " & TestResult(1)
                TestMethodFailedCol.Add TModule.Name & "." & TMethod
            End If
        Next
        TestMethodLocalTotalCount = UBound(TMethods) + 1
        TestMethodCount = TestMethodCount + TestMethodLocalTotalCount
        TestMethodSuccessCount = TestMethodSuccessCount + TestMethodLocalPassCount
        
        TestModuleCount = TestModuleCount + 1
        If Verbose And TestMethodLocalTotalCount > 0 Then Debug.Print "-------"  ' Dashes if there was an test method run
        If TestMethodLocalTotalCount = 0 Then
            If Verbose Then Debug.Print "** No test cases to run here"
            TestModuleSuccessCount = TestModuleSuccessCount + 1
        ElseIf TestMethodLocalPassCount = TestMethodLocalTotalCount Then
            If Verbose Then Debug.Print "*+ Total: " & CStr(TestMethodLocalTotalCount)
            TestModuleSuccessCount = TestModuleSuccessCount + 1
        Else
            If Verbose Then Debug.Print "*+ Total: " & CStr(TestMethodLocalTotalCount) & _
                                " / Passed: " & TestMethodLocalPassCount & _
                                " / Failed: " & (TestMethodLocalTotalCount - TestMethodLocalPassCount)
        End If
        
        If Verbose Then Debug.Print "" ' Emptyline
    Next
    
    TestMethodFailedArr = ToArray(TestMethodFailedCol)
    If Verbose And TestModuleCount > 0 Then Debug.Print "--------------"  ' Dashes if there was an test method run
    If TestModuleCount = 0 Then
        If Verbose Then Debug.Print _
            "No test modules were found. Vase is full of air."
    ElseIf TestModuleCount = TestModuleSuccessCount Then
        If Verbose Then Debug.Print _
            "+ Modules: " & CStr(TestModuleCount) & " / Methods: " & CStr(TestMethodCount)
    Else
        If Verbose Then Debug.Print _
            "- Modules: " & CStr(TestModuleCount) & " / Passed: " & CStr(TestModuleSuccessCount) & " / Failed: " & CStr(TestModuleCount - TestModuleSuccessCount) & vbCrLf & _
            "- Methods: " & CStr(TestMethodCount) & " / Passed: " & CStr(TestMethodSuccessCount) & " / Failed: " & CStr(TestMethodCount - TestMethodSuccessCount) & vbCrLf & vbCrLf & _
            "Failed Methods:" & vbCrLf & Join_(TestMethodFailedArr, vbCrLf, Prefix:="* ") & vbCrLf
    End If
    
    Dim Tuple As Variant
    Tuple = Array()
    RunVaseSuite = Tuple
End Function

'# This finds the modules that are deemed as test modules
Public Function FindTestModules(Book As Workbook) As Variant
    Dim Module As VBComponent, Modules As Variant, Index As Integer
    Modules = Array()
    Index = 0
    ReDim Modules(0 To Book.VBProject.VBComponents.Count)
    For Each Module In Book.VBProject.VBComponents
        If Module.Name Like VaseConfig.TEST_MODULE_PATTERN Then
            Set Modules(Index) = Module
            Index = Index + 1
        End If
    Next
    
    ' Fit array
    If Index = 0 Then
        Modules = Array()
    Else
        ReDim Preserve Modules(0 To Index - 1)
    End If
    
    FindTestModules = Modules
End Function

'# Finds the test methods to execute for a module
'@ Return: A zero-based string array of the method names to execute
Public Function FindTestMethods(Module As VBComponent) As Variant
    Dim Methods As Variant, Index As Integer, LineIndex As Integer, CodeLine As String
    Methods = Array()
    ReDim Methods(0 To Module.CodeModule.CountOfLines)
    
    For LineIndex = 1 To Module.CodeModule.CountOfLines
        CodeLine = Module.CodeModule.Lines(LineIndex, 1)
        If CodeLine Like METHOD_HEADER_PATTERN Then
            Dim LeftPos As Integer, RightPos As Integer
            LeftPos = InStr(CodeLine, "Sub") + 4
            RightPos = InStr(LeftPos, CodeLine, "(") - 1
            
            Methods(Index) = Mid(CodeLine, LeftPos, RightPos - LeftPos + 1)
            Index = Index + 1
        End If
    Next
    
    If Index = 0 Then
        Methods = Array()
    Else
        ReDim Preserve Methods(0 To Index - 1)
    End If
    FindTestMethods = Methods
End Function

'# Runs a test method, this assumes just it is a sub with no parameters.
'# This also encloses it in a block for protection
'@ Return: A 2-tuple consisting of a boolean indicating success and a string indicating the assertion where it failed
Public Function RunTestMethod(Book As Workbook, ModuleName As String, MethodName As String) As Variant
On Error GoTo ErrHandler:
    Application.Run Book.Name & "!" & ModuleName & "." & MethodName
ErrHandler:
    Dim HasError As Boolean, Tuple As Variant
    HasError = (Err.Number <> 0)
    If HasError Then
        Tuple = Array(False, "ExceptionRaised(" & Err.Number & "):  " & Err.Description)
    Else
        Tuple = Array(VaseAssert.TestResult, VaseAssert.FirstFailedTestMethod & ": " & VaseAssert.FirstFailedTestMessage)
    End If
    Err.Clear
    RunTestMethod = Tuple
End Function

'=======================
'-- Helper Functions  --
'=======================

'# Clears the intermediate screen
Public Sub ClearScreen()
    Application.SendKeys "^g ^a {DEL}"
    DoEvents
End Sub

'# Simple zip of two arrays, returns an array of 2-tuples
'# This assumes that arrays are zero-indexed
Public Function Zip(LeftArr As Variant, RightArr As Variant) As Variant
    If UBound(LeftArr) = -1 Or UBound(RightArr) = -1 Then
        Zip = Array()
        Exit Function
    End If
    
    Dim ZipArr As Variant, Index As Integer
    ZipArr = Array()
    ReDim ZipArr(0 To IIf(UBound(LeftArr) > UBound(RightArr), UBound(RightArr), UBound(LeftArr))) ' Take minimum of the two

    For Index = 0 To UBound(ZipArr)
        ZipArr(Index) = Array(LeftArr(Index), RightArr(Index))
    Next
    Zip = ZipArr
End Function

'# Finds a value in an array of values, this assumes the elements can be matched using the equality operator
Public Function InArray(Look As Variant, Arr As Variant) As Boolean
    Dim Elem As Variant
    InArray = False
    
    If UBound(Arr) = -1 Then Exit Function ' Nothing to do

    For Each Elem In Arr
        If Elem = Look Then
            InArray = True
            Exit Function
        End If
    Next
End Function

'# Converts a collection to an array
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

'# Joins an array assuming all entries are string
Public Function Join_(Arr As Variant, Delimiter As String, Optional Prefix As String = "") As String
    Dim StrArr() As String, Index As Integer
    ReDim StrArr(0 To UBound(Arr))
        
    For Index = 0 To UBound(StrArr)
        StrArr(Index) = Prefix & CStr(Arr(Index))
    Next
    Join_ = Join(StrArr, Delimiter)
End Function

'# Determines if a string is in an array using the like operator instead of equality
'@ Param: Patterns > An array of strings, not necessarily zero-based
'@ Return: True if the string matches any one of the patterns
Public Function InLike(Source As String, Patterns As Variant) As Boolean
    Dim Pattern As Variant
    InLike = False
    For Each Pattern In Patterns
        InLike = Source Like Pattern
        If InLike Then Exit For
    Next
End Function
