Attribute VB_Name = "modTestSuites"
'@Folder("TestingFramework")
Option Explicit
Public Sub testConsoleForm()
    With New TestRunner
        .AddTestSuite "testingConsoleModel"
        .AddTestSuite "testingConsole"
        .Run
    End With
End Sub

Private Sub testingAssertObject()
    Assert.test("1 = 2 is false").Expect(1 = 2).ToBeTrue
    Assert.test("1 = 1 is true").Expect(1 = 1).ToBeTrue
    Assert.test("2+1 = 3").Expect(1 + 2).ToEqual (3)
    Assert.test("Testing that an object is not a boolean").Expect(ThisWorkbook).ToBeTrue
    Assert.test("1 is equal to 1").Expect(1).ToEqual (1)
    Assert.test("Name of Worksheet 1 is equal to Name of ThisWorkbook.Worksheets(1)").Expect(ThisWorkbook.Worksheets(1).Name).ToEqual (Worksheets("Feuil1").Name)
End Sub
Private Sub testingSomethingElse()
    Assert.test("Trying if 1 equals 1").Expect(1).ToEqual (1)
    Assert.test("Testing the truth").Expect(True).ToBeTrue
End Sub
Private Sub testingAgain()
    Assert.test("Now testing for something else").Expect(1345 * 12).ToEqual (5)
    Assert.test("Now testing for something else").Expect(5).ToEqual (5)
End Sub
Private Sub testingToSuccess()
    Assert.test("Now we are sussessful!!!").Expect(True).ToBeTrue
    Assert.test("True is different from True should fail").Expect(True).ToBeDifferentFrom (True)
    Assert.test("True is False should fail").Expect(True).ToBeFalse
    Assert.test("True is True should pass").Expect(False).ToBeFalse
    Assert.test("1 is larger or equal to 1 should pass").Expect(1).ToBeLargerOrEqualTo (1)
    Assert.test("1 is larger than 1 should fail").Expect(1).ToBeLargerThan (1)
    Assert.test("1 is larger or equal to 0 should pass").Expect(1).ToBeLargerOrEqualTo (0)
    Assert.test("1 is larger than 0 should pass").Expect(1).ToBeLargerThan (0)
    Assert.test("1 is smaller or equal to 1 should pass").Expect(1).ToBeSmallerOrEqualTo (1)
    Assert.test("1 is smaller than 1 should fail").Expect(1).ToBeSmallerThan (1)
    Assert.test("1 is smaller or equal to 2 should pass").Expect(1).ToBeSmallerOrEqualTo (2)
    Assert.test("1 is smaller than 2 should pass").Expect(1).ToBeSmallerThan (2)
End Sub
Public Sub testGlobal()
    With New TestRunner
        .AddTestSuite "testingAssertObject"
        .AddTestSuite "testingSomethingElse"
        .AddTestSuite "testingAgain"
        .AddTestSuite "testingToSuccess"
        .Run
    End With
End Sub
Private Sub testingConsoleModel()
    With New ConsoleModel
        Assert.test("Empty console model has 0 lines").Expect(.numberOfLines).ToEqual (0)
        Dim line As ConsoleLine
        Dim block As ConsoleBlock
        Set line = New ConsoleLine
        Set block = New ConsoleBlock
        block.text = "Ma première ligne de console"
        block.fontColor = CONSOLE_COLOR_RED
        line.AddBlock block
        .AddLine line
        Assert.test("Console model has one line").Expect(.numberOfLines).ToEqual (1)
    End With
End Sub

Private Sub testingConsole()
    With New frmConsole
        Dim model As ConsoleModel
        Set model = New ConsoleModel
        Dim line As ConsoleLine
        Dim block As ConsoleBlock
        Set line = New ConsoleLine
        Set block = New ConsoleBlock
        line.AddBlock block.Create("Ma première ligne de console", CONSOLE_COLOR_RED, CONSOLE_COLOR_BLACK)
        model.AddLine line
        Assert.test("Console model has one line").Expect(model.numberOfLines).ToEqual (1)
        Set line = New ConsoleLine
        Set block = New ConsoleBlock
        block.fontColor = CONSOLE_COLOR_WHITE
        block.backColor = CONSOLE_COLOR_GREEN
        block.text = "   PASS   "
        line.AddBlock block
        Set block = New ConsoleBlock
        block.fontColor = CONSOLE_COLOR_WHITE
        block.backColor = CONSOLE_COLOR_BLACK
        block.text = "Test has not passed"
        line.AddBlock block
        model.AddLine line
        With New frmConsole
            .Populate model
            .Show
        End With
    End With
End Sub


Public Sub monTest()
    Dim parole As String
    parole = "RonnnPshii"
    Assert.test("Marianne va s'endormir").Expect(parole).ToEqual "RonnnPshiii"
End Sub
Public Sub test()
    With New TestRunner
        .AddTestSuite "monTest"
        .Run
    End With
End Sub

Public Sub testForm()
    Dim model As ConsoleModel
    Set model = New ConsoleModel
    
    model.AddSingleBlockLine "Bonjour, Marianne !", CONSOLE_COLOR_RED, CONSOLE_COLOR_BLACK
    
    Dim line As ConsoleLine
    Set line = New ConsoleLine
    
    line.AddSingleBlock "Bonjour, Marianne !", CONSOLE_COLOR_RED, CONSOLE_COLOR_BLACK
    line.AddSingleBlock "   -", , CONSOLE_COLOR_BLACK
    line.AddSingleBlock "Comment ça va ?", , , , "Comic Sans MS"
    model.AddLine line
    
    With New frmConsole
        .Populate model
        .Show
    End With
End Sub


