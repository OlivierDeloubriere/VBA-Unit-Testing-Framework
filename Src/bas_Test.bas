Attribute VB_Name = "bas_Test"
'@Folder("VBA_Objects_API")

Option Explicit



'Example1:
'List all Methods and Properties of the excel application Object.
Sub Test_1()

    Dim oObj As Object
    Dim vFuncArray()

    Set oObj = Application ''<=== Choose here a target object as required.

    vFuncArray = GetObjectFunctions(TheObject:=oObj, FuncType:=VbGet + VbLet + VbSet + VbMethod)

    If UBound(vFuncArray) Then
        With ThisWorkbook.Sheets(1)
            .Range("a2") = "Object Browsed:" & Space(2) & "(" & oObj.Name & ")"
            .Range("b2") = "Total Functions Found:" & Space(2) & "(" & UBound(vFuncArray, 1) & ")"
            .Range("a4").Resize(Rows.Count - 4, 6).ClearContents
            .Range("a4").Resize(UBound(vFuncArray, 1) + 1, 6) = vFuncArray
            .Range("a4").Select
        End With
    End If

End Sub



'Example2:
'List all Methods and Properties of Class1
'Sub Test_2()
'
'    Dim oClass As New Class1
'    Dim vFuncArray() As Variant
'
'    vFuncArray = GetObjectFunctions(TheObject:=oClass, FuncType:=VbGet + VbSet + VbLet + VbMethod)
'
'    If UBound(vFuncArray) Then
'        With Sheet1
'            .Range("a2") = "Object Browsed:" & Space(2) & "(Class1)"
'            .Range("b2") = "Total Functions Found:" & Space(2) & "(" & UBound(vFuncArray, 1) & ")"
'            .Range("a4").Resize(.Rows.Count - 4, 6).ClearContents
'            .Range("a4").Resize(UBound(vFuncArray, 1) + 1, 6) = vFuncArray
'            .Range("a4").Select
'        End With
'    End If
'
'End Sub

'
'Sub ClearTable()
'    With Sheet1
'        .Range("a2") = "Object Browsed:"
'        .Range("b2") = "Total Functions Found:"
'        .Range(Range("a4"), .Range("a4").End(xlDown).Offset(, 5)).ClearContents
'    End With
'End Sub
