Attribute VB_Name = "TestCustomTimer"
'@Folder("CustomTimer")
Option Explicit

Public Sub RunTestTimer()
With New TestRunner
    .AddTestSuite "TimerObjectExists"
    .Run
End With
End Sub

Public Sub TimerObjectExists()
    Dim cTimer As New CustomTimer
    Assert.test("Timer can be initialized").Expect(cTimer).ToBeSomething
    Assert.test("Timer has a Start method").Expect(cTimer).ToHaveMethod ("Start")
    cTimer.Start
    Assert.test("Timer can be started and log a starting time").Expect(cTimer.startTime).ToBeLargerThan (0)
    Assert.test("Timer has a LogTime method").Expect(cTimer).ToHaveMethod ("LogTime")
    cTimer.LogTime
    Application.Wait CDbl(Now) + 1 / 24 / 36000
    cTimer.LogTime
    Application.Wait CDbl(Now) + 2 / 24 / 3600
    cTimer.LogTime
    Assert.test("Timer can log 3 times after starting").Expect(cTimer.loggedTimes.Count).ToEqual (4)
    Assert.test("Timer can log elapsed times").Expect(cTimer.elapsedTimes.Count).ToEqual (cTimer.loggedTimes.Count - 1)
    
    Dim totalElapsedTime As Double
    Dim totalLoggedTime As Double
    totalElapsedTime = cTimer.elapsedTimes(1) + cTimer.elapsedTimes(2) + cTimer.elapsedTimes(3)
    totalLoggedTime = (cTimer.loggedTimes(cTimer.loggedTimes.Count) - cTimer.startTime)
    Assert.test("Timer total elapsed times is correct").Expect(totalElapsedTime).ToEqual (totalLoggedTime)
    Assert.test("Timer has property totalElapsedTime").Expect(cTimer).ToHaveMethod ("totalElapsedTime")
    Assert.test("Timer totalElapsedTime is the sum of all elapsed times").Expect(cTimer.totalElapsedTime).ToEqual (totalElapsedTime)
End Sub
