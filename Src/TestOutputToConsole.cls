VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestOutputToConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestingFramework")
'@PredeclaredID
Option Explicit

Public Sub OutputTestResults(ByVal tRunner As TestRunner)
    Dim msg As String
    msg = MessageToBeOutput(tRunner)
    Debug.Print msg
End Sub
Public Function MessageToBeOutput(ByVal tRunner As TestRunner) As String
    Dim outputMsg As String
    outputMsg = "RUNNING TESTS ---------------------------------------------" & vbCrLf
    Dim tCollector As TestCollector
    For Each tCollector In tRunner.listOfTestCollectors
        outputMsg = outputMsg & vbCrLf & OutputTestMessageFromTestCollector(tCollector)
    Next tCollector
    outputMsg = outputMsg & vbCrLf
    outputMsg = outputMsg & vbCrLf & "END TESTS --------------------------------------------------------"
    MessageToBeOutput = outputMsg
End Function
Private Function OutputTestMessageFromTestCollector(ByVal tCollector As TestCollector) As String
    Dim headerLine As String
    Dim footerLine As String
    headerLine = " Found " & tCollector.listOfTestResults.Count & " test(s) in " & tCollector.nameOfCaller
    OutputTestMessageFromTestCollector = headerLine & outputTestsResult(tCollector) & vbCrLf & footerLine
End Function

Private Function outputTestsResult(ByVal tCollector As TestCollector) As String
    Dim outputMessage As String
    Dim numPass As Byte
    Dim numFail As Byte
    Dim numInconclusive As Byte
    
    Dim currentTestResult As testResult
    For Each currentTestResult In tCollector.listOfTestResults
        outputMessage = outputMessage & vbCrLf & TestResultMessage(currentTestResult)
        If currentTestResult.status = Pass Then numPass = numPass + 1
        If currentTestResult.status = Fail Then numFail = numFail + 1
        If currentTestResult.status = Inconclusive Then numInconclusive = numInconclusive + 1
    Next currentTestResult
    outputMessage = outputMessage & vbCrLf & SubLevelTestAgregationMessage(numPass, numFail, numInconclusive)
    outputTestsResult = outputMessage
End Function
Private Function StatusMessage(ByVal status As TestStatus)
    If status = Fail Then StatusMessage = "   |  FAIL        "
    If status = Pass Then StatusMessage = "   |  PASS        "
    If status = Inconclusive Then StatusMessage = "   |  INCONCLUSIVE"
End Function

Private Function TestResultMessage(ByVal tResult As testResult) As String
    Dim outPut As String
    outPut = StatusMessage(tResult.status) & "  -> " & tResult.testDescription
    If Not tResult.resultDescription = vbNullString Then outPut = outPut & ": " & tResult.resultDescription
    TestResultMessage = outPut
End Function

Private Function SubLevelTestAgregationMessage(ByVal numPass As Byte, numFail As Byte, numInc As Byte) As String
    Dim message As String
    message = "   +-------------------------------------------------+"
    message = message & vbCrLf & "   |   PASS: " & Left(numPass & " ", 2) & "  |   " & "FAIL: " & Left(numFail & " ", 2) & "  |   " & "INCONCLUSIVE: " & Left(numInc & " ", 2) & "  |"
    message = message & vbCrLf & "   +-------------------------------------------------+"
    SubLevelTestAgregationMessage = message
End Function
