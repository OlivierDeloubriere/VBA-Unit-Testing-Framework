VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestingFramework")
'@PredeclaredID
Option Explicit

Private Type TAssert
    result As testResult
End Type

Private this As TAssert

Private Property Get result() As testResult
    Set result = this.result
End Property

Private Property Set result(ByVal RHS As testResult)
    Set this.result = RHS
End Property

Private Sub Class_Initialize()
    Set this.result = New testResult
End Sub

Public Function test(ByVal testDescription As String) As Assert
    If this.result Is Nothing Then Set this.result = New testResult
    this.result.testDescription = testDescription
    Set test = Me
End Function

Public Function Expect(ByVal computedValue As Variant) As ExpectedAssert
    If this.result Is Nothing Then Set this.result = New testResult
    
    Dim resultingExpectedAssert As New ExpectedAssert
    If IsObject(computedValue) Then
        Set this.result.computedValue = computedValue
    Else
        this.result.computedValue = computedValue
    End If
    Set resultingExpectedAssert.result = this.result
    Set Expect = resultingExpectedAssert
    Set this.result = Nothing
End Function

