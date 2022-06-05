VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpectedAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestingFramework")

Option Explicit

Private Type TExpectedAssert
    result As testResult
End Type

Public Event OnTestHasCompleted(ByVal result As testResult)

Private this As TExpectedAssert
Public Property Set result(ByVal RHS As testResult)
    Set this.result = RHS
End Property
Private Sub Class_Initialize()
    Set this.result = New testResult
End Sub

Public Sub ToEqual(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If Not this.result.computedValue = this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should equal " & this.result.expectedValue
    End If
    If this.result.computedValue = this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeDifferentFrom(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If this.result.computedValue = this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should be different from " & this.result.expectedValue
    End If
    If Not this.result.computedValue = this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeLargerThan(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If Not this.result.computedValue > this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should be larger than " & this.result.expectedValue
    End If
    If this.result.computedValue > this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeLargerOrEqualTo(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If Not this.result.computedValue >= this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should be larger or equal to " & this.result.expectedValue
    End If
    If this.result.computedValue >= this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeSmallerThan(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If Not this.result.computedValue > this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should be smaller than " & this.result.expectedValue
    End If
    If this.result.computedValue < this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeSmallerOrEqualTo(ByVal expectedValue As Variant)
    this.result.expectedValue = expectedValue
    If Not this.result.computedValue <= this.result.expectedValue Then
        this.result.status = Fail
        this.result.resultDescription = this.result.computedValue & " should be smaller or equal to " & this.result.expectedValue
    End If
    If this.result.computedValue <= this.result.expectedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeNothing()
    
    If Not IsObject(this.result.computedValue) Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not an object"
        OutPutTest this.result
        Exit Sub
    End If
    
    If Not (this.result.computedValue Is Nothing) Then
        this.result.status = Fail
        this.result.resultDescription = "Computed value is Something but should be Nothing"
    End If
    
    If this.result.computedValue Is Nothing Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeSomething()
    
    If Not IsObject(this.result.computedValue) Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not an object"
        OutPutTest this.result
        Exit Sub
    End If
    
    If this.result.computedValue Is Nothing Then
        this.result.status = Fail
        this.result.resultDescription = "Computed value is Nothing but should be Something"
    End If
    
    If Not (this.result.computedValue Is Nothing) Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToBeFalse()
    If Not VarType(this.result.computedValue) = vbBoolean Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not a boolean expression"
        OutPutTest this.result
        Exit Sub
    End If
    
    If this.result.computedValue Then
        this.result.status = Fail
        this.result.resultDescription = "Computed value is True but should be False"
    End If
    
    If Not this.result.computedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
    
End Sub
Public Sub ToBeTrue()
    If Not VarType(this.result.computedValue) = vbBoolean Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not a boolean expression"
        OutPutTest this.result
        Exit Sub
    End If
    
    If Not this.result.computedValue Then
        this.result.status = Fail
        this.result.resultDescription = "Computed value is False but should be True"
    End If
    
    If this.result.computedValue Then
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
    
End Sub
Public Sub ToHaveMethod(ByVal methodName As String)
    
    If Not IsObject(this.result.computedValue) Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not an object"
        OutPutTest this.result
        Exit Sub
    End If
    
    If Not ObjectHasMethod(this.result.computedValue, methodName) Then
        this.result.status = Fail
        this.result.resultDescription = "The object does not have a " & methodName & " method"
    Else
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub
Public Sub ToHaveProperty(ByVal propertyName As String)
    
    If Not IsObject(this.result.computedValue) Then
        this.result.status = Inconclusive
        this.result.resultDescription = "Computed value is not an object"
        OutPutTest this.result
        Exit Sub
    End If
    
    If Not ObjectHasProperty(this.result.computedValue, propertyName) Then
        this.result.status = Fail
        this.result.resultDescription = "The object does not have a " & propertyName & " property"
    Else
        this.result.status = Pass
        this.result.resultDescription = ""
    End If
    
    OutPutTest this.result
End Sub

Private Sub OutPutTest(ByVal result As testResult)
    If pTestCollector Is Nothing Then
        MsgBox "Cannot Assert, no object TestRunner found!"
        Exit Sub
    End If
    pTestCollector.AddTest result
End Sub
