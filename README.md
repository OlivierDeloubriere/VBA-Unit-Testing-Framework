# VBA Unit Testing Framework

This project is aimed at building a framework for automatic unit testing in Excel VBA projects using best OOP practices and Clean Code (design patterns, SOLID principles...).

The project is for the moment a WIP.

## Unit testing principles

Writing (or reading) unit tests should be easy and obvious. This framework provides the user with fluent keywords, reminiscent of what can be found in modern JavaScript or Java testing frameworks. Here is an example:

```vb
Assert.Test("The variable x should take the value 1").Expect(x).ToBeEqualTo(1)
```

The test will return one of the three outcomes:
- ```PASS``` if the test passes,
- ```FAIL``` if the test fails,
- and ```INCONCLUSIVE``` if the test cannot be evaluated.

In the above example if the variable  ```x``` is not a number, then the result of the test will be ```INCONCLUSIVE```.

## Test result output

The results of the tests can be read via two different output methods:
- In the *ugly* VBE immediate window (here with customized dark color theme)

<img src="./capture_VBE_Immediate_Window.jpg"/>

- In a *much prettier* userform, that plays the role of a custom console (with text formatting options that are lacking in the VBE debug immediate window)

<img src="./capture_Custom_Console.jpg"/>

## Running the tests
The tests are run by executing a ````Sub```` written in a standard module. Inside this ````Sub```` should be defined a ```TestRunner``` object. This object has two important methods:
- ```.AddTestSuite(ByVal macroName as string)``` specifies the name of the ````Sub```` containing the ````Assert```` commands
- ````.Run```` which starts the whole testing process

### Example

The test here will be run by executing the following macro ``RunTests()``:

```vb
'This code can be placed in a Standard Module
Public Sub RunTests()
    With New TestRunner
        .AddTestSuite "myTestingModule"
        .Run
    End With
End Sub

Public Sub myTestingModule()
    Dim x As Integer
    Assert.Test("The variable x should be equal to 1").Expect(x).ToEqual(1)
End Sub
```
Many different test suites can be added one after the other to the ```TestRunner``` object. Every test suite ```Sub``` can have as many ```Assert``` commands as necessary.

Test suites can be seen as tests subcategories each containing a list of individual tests.

## API

A unit test can be described (or labeled) by using the ```.Test``` method of the ```Assert``` object:

```vb 
.Test(ByVal descriptionText as String)
```

(this method is optional).

Then one specifies the variable or object, whose value or property or method must be tested, with the method ```.Expect``` of the ```Assert``` object:

```vb
Assert.Expect(ByVal computedVariant as Variant)
```

Finally, one specifies the actual test to be run with the following tests comparison methods:

```vb
.ToEqual(ByVal expectedValue as Variant)
```
```vb
.ToBeDifferentFrom(ByVal expectedValue as Variant)
```
```vb
.ToBeLargerThan(ByVal expectedValue as Variant)
```
```vb
.ToBeLargerOrEqualTo(ByVal expectedValue as Variant)
```
```vb
.ToBeSmallerThan(ByVal expectedValue as Variant)
```
```vb
.ToBeSmallerOrEqualTo(ByVal expectedValue as Variant)
```
```vb
.ToHaveMethod(ByVal expectedValue as String)
```
```vb
.ToBeTrue()
```
```vb
.ToBeFalse()
```
```vb
.ToBeNothing()
```
```vb
.ToBeSomething()
```
