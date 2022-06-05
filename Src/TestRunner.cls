VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestingFramework")
Option Explicit

Private listOfTestSuites As Collection
Public listOfTestCollectors As Collection

Private outputMsg As String

Private Sub Class_Initialize()
    Set listOfTestSuites = New Collection
    Set listOfTestCollectors = New Collection
End Sub
Public Sub AddTestSuite(ByVal nameOfTestSuite As String)
    listOfTestSuites.Add nameOfTestSuite
End Sub

Public Sub Run()
    Dim calledTestSuite As Variant
    Dim cTimer As New CustomTimer
    
    cTimer.Start
    For Each calledTestSuite In listOfTestSuites
        Set pTestCollector = New TestCollector
        pTestCollector.nameOfCaller = CStr(calledTestSuite)
        Application.Run calledTestSuite
        listOfTestCollectors.Add pTestCollector
        Set pTestCollector = Nothing
        cTimer.LogTime
    Next calledTestSuite
    
    OutputTestResults cTimer
    
End Sub

Private Sub OutputTestResults(ByVal cTimer As CustomTimer)
    'TestOutputToConsole.OutputTestResults Me
    'TestOutputToWorksheet.OutputTestResults Me, Feuil1
    TestOutputToConsoleUserForm.OutputTestResults Me, cTimer
End Sub

Public Function TotalNumberOfTestWithStatus(ByVal status As TestStatus) As Integer
    Dim numberResult As Integer
    Dim tCollector As TestCollector
    For Each tCollector In Me.listOfTestCollectors
        numberResult = numberResult + tCollector.numOfTestsWithStatus(status)
    Next tCollector
    TotalNumberOfTestWithStatus = numberResult
End Function
