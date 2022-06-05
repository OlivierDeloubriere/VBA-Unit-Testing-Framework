VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestOutputToConsoleUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TestingFramework")
Option Explicit

Public Sub OutputTestResults(ByVal tRunner As TestRunner, cTimer As CustomTimer)
    With New frmConsole
        .Populate consoleModelFromTestRunner(tRunner, cTimer)
        .Height = DEFAULT_CONSOLE_HEIGHT
        .Width = DEFAULT_CONSOLE_WIDTH
        .ScrollHeight = (.consoleListOfLines.Count + 2) * DEFAULT_CONSOLE_LINE_HEIGHT
        .ScrollTop = .ScrollHeight
        .Caption = "Test Output Console"
        .Show
    End With
End Sub
Public Function consoleModelFromTestRunner(ByVal tRunner As TestRunner, cTimer As CustomTimer) As ConsoleModel
    Dim model As New ConsoleModel
    
    model.AddSingleBlockLine "-------- RUNNING TESTS --------", CONSOLE_COLOR_YELLLOW, CONSOLE_COLOR_BLACK
    Dim tCollector As TestCollector
    If tRunner.listOfTestCollectors.Count = 0 Then
        model.AddSingleBlockLine " Found 0 test to run"
    End If
    Dim timeIndex As Integer
    timeIndex = 1
    For Each tCollector In tRunner.listOfTestCollectors
        AddAllLinesFromTestCollector model, tCollector, Round(cTimer.elapsedTimes(timeIndex) * 1000, 0)
        timeIndex = timeIndex + 1
    Next tCollector
    model.AddSingleBlockLine vbCrLf
    model.AddSingleBlockLine "---------- END TESTS ----------", CONSOLE_COLOR_YELLLOW, CONSOLE_COLOR_BLACK
    AddWrapUpLinesToModel tRunner, model
    Set consoleModelFromTestRunner = model
End Function

Public Sub AddAllLinesFromTestCollector(ByVal model As ConsoleModel, tCollector As TestCollector, elapsedTime As Double)
    Dim headerLine As String
    Dim testDescription As String
    Dim numFail As Integer
    Dim numPass As Integer
    Dim numInconclusive As Integer
    
    headerLine = " Running " & tCollector.listOfTestResults.Count & " test(s) in " & tCollector.nameOfCaller
    model.AddSingleBlockLine ""
    model.AddSingleBlockLine headerLine
    Dim tResult As testResult
    Dim line As ConsoleLine
    For Each tResult In tCollector.listOfTestResults
        If tResult.status = Pass Then numPass = numPass + 1
        If tResult.status = Fail Then numFail = numFail + 1
        If tResult.status = Inconclusive Then numInconclusive = numInconclusive + 1
        Set line = New ConsoleLine
        line.AddSingleBlock blockText:="---", fontColor:=CONSOLE_COLOR_BLACK
        line.AddBlock symbolBlock(tResult.status)
        testDescription = " " & tResult.testDescription
        If Not tResult.resultDescription = vbNullString Then testDescription = testDescription & ": " & tResult.resultDescription
        line.AddSingleBlock testDescription, , CONSOLE_COLOR_DARKGREY
        model.AddLine line
    Next tResult
    
    If numFail > 0 Then
        AddSummaryLinesToModel model, numFail, Fail, elapsedTime
        Exit Sub
    End If
    
    If numInconclusive > 0 Then
        AddSummaryLinesToModel model, numInconclusive, Inconclusive, elapsedTime
        Exit Sub
    End If
    
    If numPass > 0 Then
        AddSummaryLinesToModel model, numPass, Pass, elapsedTime
        Exit Sub
    End If
End Sub

Private Function StatusMessage(ByVal status As TestStatus)
    If status = Fail Then StatusMessage = "FAIL"
    If status = Pass Then StatusMessage = "PASS"
    If status = Inconclusive Then StatusMessage = "INCONCLUSIVE"
End Function

Private Function BackColorFromStatus(ByVal status As TestStatus)
    If status = Fail Then BackColorFromStatus = CONSOLE_COLOR_RED
    If status = Pass Then BackColorFromStatus = CONSOLE_COLOR_GREEN
    If status = Inconclusive Then BackColorFromStatus = CONSOLE_COLOR_GREY
End Function
Private Function FontColorFromStatus(ByVal status As TestStatus)
    If status = Fail Then FontColorFromStatus = CONSOLE_COLOR_RED
    If status = Pass Then FontColorFromStatus = CONSOLE_COLOR_GREEN
    If status = Inconclusive Then FontColorFromStatus = CONSOLE_COLOR_GREY
End Function

Private Function symbolBlock(ByVal status As TestStatus) As ConsoleBlock
    Dim symbol As String
    Dim symbolFont As String
    Dim symbolSize As Long
    Dim block As ConsoleBlock
    
    If status = Fail Then
        symbol = Chr(251)
        symbolFont = "Wingdings"
        symbolSize = 13
    End If
    If status = Pass Then
        symbol = Chr(252)
        symbolFont = "Wingdings"
        symbolSize = 13
    End If
    If status = Inconclusive Then
        symbol = "?"
        symbolFont = ""
        symbolSize = 11
    End If
    Set symbolBlock = ConsoleBlock.Create(symbol, FontColorFromStatus(status), , symbolSize, symbolFont)
End Function
Private Sub AddSummaryLinesToModel(ByVal model As ConsoleModel, numTests As Integer, status As TestStatus, elapsedTime As Double)
    Dim line As ConsoleLine
    Set line = New ConsoleLine
    line.AddSingleBlock blockText:="---", fontColor:=CONSOLE_COLOR_BLACK
    line.AddSingleBlock " " & StatusMessage(status), BackColorFromStatus(status), CONSOLE_COLOR_BLACK
    line.AddSingleBlock "-", BackColorFromStatus(status), FontColorFromStatus(status)
    
    Select Case status
        Case Pass
            If numTests > 1 Then
                line.AddSingleBlock " All " & numTests & " tests have passed (" & elapsedTime & "ms)", , CONSOLE_COLOR_WHITE
            Else
                line.AddSingleBlock " The test has passed (" & elapsedTime & "ms)", , CONSOLE_COLOR_WHITE
            End If
        Case Fail
            line.AddSingleBlock " Found " & numTests & " failed test(s) (" & elapsedTime & "ms)", , CONSOLE_COLOR_WHITE
        Case Inconclusive
            line.AddSingleBlock " Found " & numTests & " inconclusive test(s) (" & elapsedTime & "ms)", , CONSOLE_COLOR_WHITE
    End Select
    
    model.AddLine line
End Sub

Private Sub AddWrapUpLinesToModel(ByVal tRunner As TestRunner, ByRef model As ConsoleModel)
    Dim numPass As Integer
    Dim numInconclusive As Integer
    Dim numFail As Integer
    Dim numTests As Integer
    Dim passPercent As String
    Dim failPercent As String
    
    
    numPass = tRunner.TotalNumberOfTestWithStatus(Pass)
    numFail = tRunner.TotalNumberOfTestWithStatus(Fail)
    numInconclusive = tRunner.TotalNumberOfTestWithStatus(Inconclusive)
    numTests = numPass + numFail + numInconclusive
    
    If numTests > 0 Then
        passPercent = "(" & Round(CDbl(numPass) / CDbl(numTests) * 100, 0) & "%)"
        failPercent = "(" & Round(CDbl(numFail) / CDbl(numTests) * 100, 0) & "%)"
    End If
    
    model.AddSingleBlockLine ""
    
    Dim line As ConsoleLine
    Set line = New ConsoleLine
    
    line.AddSingleBlock " | Results:         "
    line.AddSingleBlock "      " & numPass & " " & passPercent & " passed", , CONSOLE_COLOR_GREEN
    line.AddSingleBlock ","
    line.AddSingleBlock " " & numFail & " " & failPercent & " failed", , CONSOLE_COLOR_RED
    line.AddSingleBlock ", of " & numTests & " total."
    model.AddLine line

End Sub
