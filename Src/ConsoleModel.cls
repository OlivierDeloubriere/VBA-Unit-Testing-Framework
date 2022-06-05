VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CustomConsole")
Option Explicit

Public listOfLines As Collection

Public Sub Class_Initialize()
    Set listOfLines = New Collection
End Sub
Public Sub AddLine(ByVal line As ConsoleLine)
    Me.listOfLines.Add line
End Sub

Public Function numberOfLines() As Integer
    numberOfLines = Me.listOfLines.Count
End Function

Public Function numberOfBlocks() As Integer
    Dim result As Integer
    Dim line As ConsoleLine
    For Each line In Me.listOfLines
        result = result + line.listOfBlocks.Count
    Next line
    numberOfBlocks = result
End Function

Public Sub AddSingleBlockLine(ByVal blockText As String, Optional backColor As Long = -1, Optional fontColor As Long = -1, Optional fontSize As Long = -1)
    Dim line As ConsoleLine
    Set line = New ConsoleLine
        
    line.AddBlock ConsoleBlock.Create(blockText, fontColor, backColor, fontSize)
    Me.listOfLines.Add line
End Sub
