VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestOutputToWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TestingFramework")
Option Explicit

Public Sub OutputTestResults(ByVal tRunner As TestRunner, wks As Worksheet)
    Dim msg As String
    msg = TestOutputToConsole.MessageToBeOutput(tRunner)
    wks.Range("A:A").ClearContents
    wks.Range("A1").Value = msg
End Sub
