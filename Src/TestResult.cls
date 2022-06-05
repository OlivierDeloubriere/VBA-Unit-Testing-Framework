VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestingFramework")
'@PredeclaredID
Option Explicit

Implements ITestResult

Public Enum TestStatus
    Fail
    Pass
    Inconclusive
End Enum

Private Type TTestResult
    testDescription As String
    expectedValue As Variant
    computedValue As Variant
    status As Byte
    resultDescription As String
End Type

Private this As TTestResult

Public Property Get testDescription() As String
    testDescription = this.testDescription
End Property

Public Property Let testDescription(ByVal RHS As String)
    this.testDescription = RHS
End Property

Public Property Get expectedValue() As Variant
    If IsObject(this.expectedValue) Then
        Set expectedValue = this.expectedValue
    Else
        expectedValue = this.expectedValue
    End If
End Property

Public Property Let expectedValue(ByVal RHS As Variant)
    this.expectedValue = RHS
End Property

Public Property Set expectedValue(ByVal RHS As Variant)
    Set this.expectedValue = RHS
End Property

Public Property Get computedValue() As Variant
    If IsObject(this.computedValue) Then
        Set computedValue = this.computedValue
    Else
        computedValue = this.computedValue
    End If
End Property

Public Property Let computedValue(ByVal RHS As Variant)
    this.computedValue = RHS
End Property

Public Property Set computedValue(ByVal RHS As Variant)
    Set this.computedValue = RHS
End Property

Public Property Get status() As Byte
    status = this.status
End Property

Public Property Let status(ByVal RHS As Byte)
    this.status = RHS
End Property

Public Property Get resultDescription() As String
    resultDescription = this.resultDescription
End Property

Public Property Let resultDescription(ByVal RHS As String)
    this.resultDescription = RHS
End Property

Private Property Get ITestResult_TestDescription() As String
    ITestResult_TestDescription = testDescription
End Property

Private Property Get ITestResult_ExpectedValue() As Variant
    ITestResult_ExpectedValue = expectedValue
End Property

Private Property Get ITestResult_ComputedValue() As Variant
    ITestResult_ComputedValue = computedValue
End Property

Private Property Get ITestResult_status() As Byte
    ITestResult_status = status
End Property

Private Property Get ITestResult_ResultDescription() As String
    ITestResult_ResultDescription = resultDescription
End Property

