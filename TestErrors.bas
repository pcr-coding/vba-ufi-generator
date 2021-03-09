Attribute VB_Name = "TestErrors"
' VBA UFI Generator
' Copyright (C) 2021  Philipp C. Ruedinger
' https://github.com/pcr-coding/vba-ufi-generator
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.


'@IgnoreModule UseMeaningfulName, LineLabelNotUsed, EmptyMethod, VariableNotUsed
Option Explicit
Option Private Module

'@TestModule
'@Folder("Unique Formula Identifier.Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("Return Errors")
Private Sub ErrorCountryCodeDoesNotExist()
    Const ExpectedError As Long = vbObjectError + 513
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As String
    TestResult = sut.Generate("XX", "aa1828639338661", "156920229")

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("Return Errors")
Private Sub ErrorInvalidVAT()
    Const ExpectedError As Long = vbObjectError + 515
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As String
    TestResult = sut.Generate("DE", "aa1828639338661", "156920229")

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL001()
    Const ExpectedError As Long = vbObjectError + 551
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("GMTT-2SQN-6FD-6TV1") 'only 15 characters

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL002()
    Const ExpectedError As Long = vbObjectError + 552
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("GMTT-2SQN-6FDD-6TVI") 'invalid character I

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL003()
    Const ExpectedError As Long = vbObjectError + 553
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("FMTT-2SQN-6FDD-6TV1") 'invalid first character (checksum)

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub



'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL004A()
    Const ExpectedError As Long = vbObjectError + 554
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("733S-4QNF-4MHA-DTUU") 'invalid country group code

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL004B()
    Const ExpectedError As Long = vbObjectError + 555
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("123S-0Q2K-NMHH-WCJQ") 'inconsistency between country group code and number of bits / country code

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub



'@TestMethod("Return Errors")
Private Sub ErrorInvalidUFI_VAL005()
    Const ExpectedError As Long = vbObjectError + 556
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator

    'Act:
    Dim TestResult As DecodedUFI
    Set TestResult = sut.Decode("U23S-PQ2V-AMH9-VVRG") 'version bit = 1

Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

