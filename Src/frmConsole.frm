VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsole 
   Caption         =   "Custom Console"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15960
   OleObjectBlob   =   "frmConsole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "CustomConsole"
Option Explicit
Public consoleListOfLines As Collection
Public consolelistOfBlocks As Collection
Public model As ConsoleModel


Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
       
    Me.Height = DEFAULT_CONSOLE_HEIGHT
    Me.Width = DEFAULT_CONSOLE_WIDTH

End Sub

Private Function TopOfNextLine()
    TopOfNextLine = DEFAULT_MARGIN_TOP + consoleListOfLines.Count * (DEFAULT_CONSOLE_LINE_HEIGHT)
End Function
Public Function DisplayAllLines()
    Dim line As ConsoleLine
    For Each line In Me.model.listOfLines
        DisplaySingleLine line
        consoleListOfLines.Add line
    Next line
End Function
Private Function CurrentNumberOfLabels()
    Dim result As Integer
    Dim line As ConsoleLine
    If Me.consoleListOfLines.Count = 0 Then
        CurrentNumberOfLabels = 0
    Else
        For Each line In Me.consoleListOfLines
            result = result + line.listOfBlocks.Count
        Next line
        CurrentNumberOfLabels = result
    End If
End Function
Private Function DisplaySingleLine(ByVal line As ConsoleLine)
    Dim block As ConsoleBlock
    Dim label As Control
    Dim nextLeftPosition As Integer
    For Each block In line.listOfBlocks
        AddLabelFromBlock block, nextLeftPosition
        consolelistOfBlocks.Add block
    Next block
End Function

Public Sub Populate(ByVal model As ConsoleModel)
    Set Me.consoleListOfLines = New Collection
    Set Me.consolelistOfBlocks = New Collection
    Set Me.model = model
    DisplayAllLines
    Me.Repaint
End Sub

Public Sub AddLabelFromBlock(ByVal block As ConsoleBlock, ByRef leftPosition As Integer)
    Dim nextLabelIndex As Integer
    Dim label As Control
    nextLabelIndex = CurrentNumberOfLabels + 1
    Set label = Me.Controls.Add("Forms.Label.1", "block" & nextLabelIndex)
    label.Top = TopOfNextLine
    label.Left = DEFAULT_MARGIN_LEFT + leftPosition
    label = block.text
    label.backColor = block.backColor
    label.ForeColor = block.fontColor
    label.Font.Size = block.fontSize
    label.Font.Name = block.fontName
    
    If block.fontName = "Wingdings" Then
        label.Font.Charset = 2
        label.Width = 7
    Else
        label.Font.Charset = 1
        label.AutoSize = True
        label.WordWrap = False
    End If
    'label.Font.Weight = 350
    'label.Height = DEFAULT_CONSOLE_LINE_HEIGHT
    leftPosition = leftPosition + label.Width - 1
End Sub
