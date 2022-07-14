Attribute VB_Name = "TestCountries"
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

'@TestMethod("TestCountry")
Private Sub TestCountryAT()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "AT"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("U12345678", "178956970", "C23S-PQ2V-AMH9-VVRF")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryBE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "BE"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("0987654321", "89478485", "U1JV-SUMH-N988-U751")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TestCountry")
Private Sub TestCountryBG()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 2
    Const CountryCode As String = "BG"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("987654321", "89478485", "A1JV-EUD3-498W-U23R")
    Test(2) = Array("9987654321", "252644556", "80SW-N2UD-FGRY-F6DA")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryCY()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 2
    Const CountryCode As String = "CY"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("12345678C", "87654321", "7EHW-3KC7-748V-S1RH")
    Test(2) = Array("12345678Y", "87654321", "FEHW-3KH5-X487-SNRD")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryCZ()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 3
    Const CountryCode As String = "CZ"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("81726354", "156920229", "HKC1-Q7N0-DKF6-7N3E")
    Test(2) = Array("978563421", "15790899", "24VQ-VGV0-WF16-7WC9")
    Test(3) = Array("9785634210", "268435455", "W3NN-SKC4-JXSS-V4WG")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryDE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "DE"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("112358132", "134217728", "KMTT-DSP3-7FD7-6RWY")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryDK()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "DK"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("31415926", "524544", "3FQU-5GP0-Y105-J64N")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryEE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "EE"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("271828182", "230087533", "QY3Q-327C-QDPR-EE11")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryES()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "ES"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("A1234567X", "178956970", "T23S-6QD4-3MHA-VT69")
              'Official sample in table B-1 was missing (this is a own online generator verified sample)
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryFR()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 4
    Const CountryCode As String = "FR"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("RF987654321", "134217728", "6KTT-PSK1-AFDM-KEFU")
    Test(2) = Array("ZY999999999", "230087533", "NX3Q-2263-2DP5-UQ4T")
    Test(3) = Array("01012345678", "268435455", "F3NN-3K1J-EXSK-7PHY")
    Test(4) = Array("10012345678", "268435455", "23NN-5KKE-GXSF-7MD3")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryGB()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 5
    Const CountryCode As String = "GB"
    Const MultiCountryCode As String = "GB|XN"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("987654321", "156920229", "GJC1-S7TH-AKFK-TR8P")
    Test(2) = Array("999987654321", "156920229", "CJC1-47ES-AKFS-WTFW")
    Test(3) = Array("999999999999", "156920229", "3JC1-47ET-6KFH-WMV8")
    Test(4) = Array("ZY123", "268435455", "53NN-7KTT-1XS1-DDPH")
    Test(5) = Array("AB987", "268435455", "M3NN-7KTS-YXSK-D8UW")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = MultiCountryCode, "ResultDecode.CountryCode " & i & " should be " & MultiCountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryXN()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 5
    Const CountryCode As String = "XN"
    Const MultiCountryCode As String = "GB|XN"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("987654321", "156920229", "GJC1-S7TH-AKFK-TR8P")
    Test(2) = Array("999987654321", "156920229", "CJC1-47ES-AKFS-WTFW")
    Test(3) = Array("999999999999", "156920229", "3JC1-47ET-6KFH-WMV8")
    Test(4) = Array("ZY123", "268435455", "53NN-7KTT-1XS1-DDPH")
    Test(5) = Array("AB987", "268435455", "M3NN-7KTS-YXSK-D8UW")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = MultiCountryCode, "ResultDecode.CountryCode " & i & " should be " & MultiCountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryGR()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "GR"
    Const MultiCountryCode As String = "GR|EL"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("567438921", "66260700", "QNWM-9X6E-E46N-G4GJ")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = MultiCountryCode, "ResultDecode.CountryCode " & i & " should be " & MultiCountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryEL()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "EL"
    Const MultiCountryCode As String = "GR|EL"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("567438921", "66260700", "QNWM-9X6E-E46N-G4GJ")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = MultiCountryCode, "ResultDecode.CountryCode " & i & " should be " & MultiCountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryFI()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "FI"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("18273645", "29979245", "VWF9-CDT4-2S2N-PDTV")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryHR()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "HR"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("16021765654", "268435455", "53NN-KKPX-SXSD-QJY7")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryHU()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "HU"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("22334455", "238219293", "AU06-7HHD-64QN-8RHF")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryIE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 6
    Const CountryCode As String = "IE"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("9Z54321Y", "134217728", "GMTT-2SQN-6FDD-6TV1")
    Test(2) = Array("9+54321Y", "134217728", "KMTT-2SQQ-0FDQ-6A5D")
    Test(3) = Array("9*54321Y", "134217728", "GMTT-2SQR-UFD0-6TFR")
    Test(4) = Array("9876543Z", "230087533", "JY3Q-R2M8-GDP2-DQRS")
    Test(5) = Array("9876543ZW", "230087533", "XY3Q-S215-2DPF-DA4U")
    Test(6) = Array("9876543AB", "182319099", "TUG4-PE6C-4XHP-RSAM")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryIS()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 2
    Const CountryCode As String = "IS"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("AB3D5F", "182319099", "XUG4-WE32-RXHD-RHWU")
    Test(2) = Array("1ZY2BA", "268435455", "53NN-1KDC-JXSH-W6WV")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryIT()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "IT"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("14286244833", "214783315", "WK3F-PYSX-TXM0-K9AV")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryLI()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "LI"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("99999", "182319099", "CUG4-FEHP-HXHW-SC4E")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryLT()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 2
    Const CountryCode As String = "LT"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("987654321", "15790899", "W3VQ-HGW8-UF12-W7QP")
    Test(2) = Array("987654321098", "156920229", "SJC1-P7FR-DKF3-Y2YC")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryLU()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "LU"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("16726218", "214783315", "FK3F-8YC6-1XMK-SAQ5")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryLV()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "LV"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("39903127176", "182319099", "WUG4-5E2S-UXHN-M5C7")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryMT()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "MT"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("99887766", "83144621", "7KW0-SMM5-2Q7N-4K8D")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryNL()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "NL"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("999999999B77", "96485337", "QJ0V-J1JU-9Y8W-37TG")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryNO()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "NO"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("958473621", "268435455", "63NN-7KPT-WXS8-WGYA")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryPL()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "PL"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("4835978701", "19621109", "XRMW-9HU2-PT1U-7JNN")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryPT()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "PT"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("998776554", "30051977", "K9WS-JKK3-WS2E-1WSC")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountryRO()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 2
    Const CountryCode As String = "RO"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("98", "252644556", "E0SW-U2CF-KGR6-FRKW")
    Test(2) = Array("9081726354", "214783315", "2K3F-QYHK-YXMU-RTN5")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountrySE()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "SE"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("987654321098", "156920229", "1KC1-87DH-0KFR-2WGE")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountrySI()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "SI"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("12345678", "178956970", "U23S-WQK5-AMH7-V03N")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("TestCountry")
Private Sub TestCountrySK()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = "SK"
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("9987654321", "252644556", "N0SW-W2AP-FGRV-F9RH")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("TestCountry")
Private Sub TestCompanyKey()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut As UFIgenerator
    Set sut = New UFIgenerator
    
    Const n As Long = 1
    Const CountryCode As String = vbNullString
    
    Dim Test() As Variant
    ReDim Test(1 To n) As Variant
    
    Test(1) = Array("1828639338661", "156920229", "NJC1-671A-UKF3-J0M8")
    
    'Act:
    Dim ResultGenerate() As String
    ReDim ResultGenerate(1 To n) As String
    Dim ResultDecode() As DecodedUFI
    ReDim ResultDecode(1 To n) As DecodedUFI
    
    Dim i As Long
    For i = 1 To n
        ResultGenerate(i) = sut.Generate(CountryCode, Test(i)(0), Test(i)(1))
        Set ResultDecode(i) = sut.Decode(Test(i)(2))
    Next i
    
    'Assert:
    For i = 1 To n
        Assert.IsTrue ResultGenerate(i) = Test(i)(2), "ResultGenerate " & i & " should be " & Test(i)(2) & " but was " & ResultGenerate(i)
        
        If Not ResultDecode(i) Is Nothing Then
            Assert.IsTrue ResultDecode(i).CountryCode = CountryCode, "ResultDecode.CountryCode " & i & " should be " & CountryCode & " but was " & ResultDecode(i).CountryCode
            Assert.IsTrue ResultDecode(i).VAT = Test(i)(0), "ResultDecode.VAT " & i & " should be " & Test(i)(0) & " but was " & ResultDecode(i).VAT
            Assert.IsTrue ResultDecode(i).FormulationNumber = Test(i)(1), "ResultDecode.FormulationNumber " & i & " should be " & Test(i)(1) & " but was " & ResultDecode(i).FormulationNumber
        Else
            Assert.Fail "Decode of UFI " & Test(i)(2) & " failed with no result."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


