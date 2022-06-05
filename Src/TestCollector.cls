VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestingFramework")
'@PredeclaredId
Option Explicit

Public listOfTestResults As Collection
Public nameOfCaller As String

Private Sub Class_Initialize()
    Set listOfTestResults = New Collection
End Sub

Public Sub AddTest(ByVal tResult As testResult)
    listOfTestResults.Add tResult
End Sub

Public Function numOfTestsWithStatus(ByVal status As TestStatus) As Integer
    Dim numResult As Integer
    Dim tResult As testResult
    For Each tResult In Me.listOfTestResults
        If tResult.status = status Then numResult = numResult + 1
    Next tResult
    numOfTestsWithStatus = numResult
End Function
