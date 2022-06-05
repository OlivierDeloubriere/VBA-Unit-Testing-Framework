VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TestingFramework")
'@Interface

Option Explicit


Public Property Get testDescription() As String
End Property

Public Property Get expectedValue() As Variant
End Property

Public Property Get computedValue() As Variant
End Property

Public Property Get status() As Byte
End Property

Public Property Get resultDescription() As String
End Property
