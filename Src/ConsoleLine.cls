VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CustomConsole")
Option Explicit

Public listOfBlocks As Collection

Public Sub Class_Initialize()
    Set listOfBlocks = New Collection
End Sub
Public Sub AddBlock(ByVal block As ConsoleBlock)
    Me.listOfBlocks.Add block
End Sub
Public Function SingleBlockLine(ByVal blockText As String, Optional backColor As Long = -1, Optional fontColor As Long = -1, Optional fontSize As Long = -1) As ConsoleLine
    Dim line As ConsoleLine
    Set line = New ConsoleLine
    line.AddBlock ConsoleBlock.Create(blockText, backColor, fontColor, fontSize)
    Set SingleBlockLine = line
End Function
Public Sub AddSingleBlock(ByVal blockText As String, Optional backColor As Long = -1, Optional fontColor As Long = -1, Optional fontSize As Long = -1, Optional fontName As String = vbNullString)
    Me.listOfBlocks.Add ConsoleBlock.Create(blockText, fontColor, backColor, fontSize, fontName)
End Sub
