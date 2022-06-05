VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CustomTimer")
Option Explicit


Public startTime As Double
Public loggedTimes As Collection
Public elapsedTimes As Collection

Public Sub Start()
    startTime = CDbl(timer)
    Set loggedTimes = New Collection
    Set elapsedTimes = New Collection
    loggedTimes.Add startTime
End Sub

Public Sub LogTime()
    loggedTimes.Add CDbl(timer)
    elapsedTimes.Add (loggedTimes(loggedTimes.Count) - loggedTimes(loggedTimes.Count - 1))
End Sub
Public Function totalElapsedTime() As Double
    totalElapsedTime = (loggedTimes(loggedTimes.Count) - loggedTimes(1))
End Function

