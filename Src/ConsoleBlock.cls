VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CustomConsole")
'@PredeclaredId
Option Explicit

Public text As String
Public fontColor As Long
Public backColor As Long
Public fontSize As Integer
Public fontName As String

Private Sub Class_Initialize()
    Me.fontColor = CONSOLE_COLOR_WHITE
    Me.backColor = CONSOLE_COLOR_BLACK
    Me.fontSize = DEFAULT_CONSOLE_FONTSIZE
    Me.fontName = DEFAULT_CONSOLE_FONT
End Sub

Public Function Create(ByVal text As String, Optional specifiedFontColor As Long = -1, Optional specifiedBackColor As Long = -1, Optional specifiedFontSize As Long = -1, Optional specifiedFontName As String = vbNullString) As ConsoleBlock
    Dim block As New ConsoleBlock
    block.text = text
    If Not specifiedFontName = vbNullString Then block.fontName = specifiedFontName
    If Not specifiedFontColor = -1 Then block.fontColor = specifiedFontColor
    If Not specifiedBackColor = -1 Then block.backColor = specifiedBackColor
    If Not specifiedFontSize = -1 Then block.fontSize = specifiedFontSize
    Set Create = block
End Function
