VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UFIgenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


'@IgnoreModule UseMeaningfulName
Option Explicit
'@Folder("Unique Formula Identifier")

Const Base31CharSet As String = "0123456789ACDEFGHJKMNPQRSTUVWXY"  'Table 2-3: Base-31 character set
Const Version As Long = 0



Public Function Generate(ByVal CountryCode As String, ByVal VAT As String, ByVal FormulationNumber As Long) As String
    '@Ignore UseMeaningfulName
    Dim g As String
    g = GetCountryGroupCodeBinary(CountryCode)
    
    '@Ignore UseMeaningfulName
    Dim C As String
    C = GetCountryCodeBinary(CountryCode)
    
    '@Ignore UseMeaningfulName
    Dim b As Long
    b = Len(C)
    
    '@Ignore UseMeaningfulName
    Dim F As String
    F = Dec2Bin(FormulationNumber, 28)
    
    '@Ignore UseMeaningfulName
    Dim V As Variant
    V = GetBinaryVAT(CountryCode, VAT)

    '@Ignore UseMeaningfulName
    Dim n As String
    n = Dec2Bin(V, 41 - b)
    
    Dim PayloadBinary As String
    PayloadBinary = F & g & C & n & Version
    
    Dim Payload As Variant
    Payload = Bin2Dec(PayloadBinary)
    
    
    'STEP 2 - UFI payload in base-31
    Dim Base31() As Long '1 … 15
    Base31 = GetBase31(Payload)
    
    Dim Base31Reorg() As Variant '0 … 15
    Base31Reorg = ReorganizeBase31(Base31)
    

    'STEP 4 - Checksum
    Dim WeightedSum As Long
    '@Ignore UseMeaningfulName
    Dim p As Long
    For p = 1 To 15
        WeightedSum = WeightedSum + (p + 1) * Base31Reorg(p)
    Next p
    
    'Checksum (first character)
    Base31Reorg(0) = (31 - (WeightedSum Mod 31)) Mod 31
    
    
    'STEP 3 - Turn into alpha-numeric UFI
    Dim UFI As String

    Dim i As Long
    For i = 0 To 15
        If i <> 0 And i Mod 4 = 0 Then UFI = UFI & "-"
        UFI = UFI & Mid$(Base31CharSet, Base31Reorg(i) + 1, 1)
    Next i

    Generate = UFI
End Function



Public Function Decode(ByVal UFI As String) As DecodedUFI
'decoding includes all validating steps
    
    Dim localUFI As String
    localUFI = UFI
    
    'remove spaces
    localUFI = Replace$(localUFI, " ", vbNullString, Compare:=vbBinaryCompare)
    
    
    'VAL001 - UFI has 19 characters in the correct format?
    If Not IsMatchingPattern("[^-]{4}-[^-]{4}-[^-]{4}-[^-]{4}", localUFI) Then
        'ERR: 551 VAL001 - Invalid UFI.  UFI must have 19 characters (including hypens, excluding spaces)
        Err.Raise vbObjectError + 551, "VAL001 - Invalid UFI.", "UFI must have 19 characters (including hypens, excluding spaces)."
        Exit Function
    End If
    
    
    'remove dashes
    localUFI = Replace$(localUFI, "-", vbNullString, Compare:=vbBinaryCompare)
    
    
    'VAL002 - UFI has only allowed characters?
    Dim WeightedSum As Long
    Dim Base31Reorg(0 To 15) As Long
    
    Dim i As Long
    For i = 1 To Len(localUFI)
        Dim Character As String
        Character = Mid$(localUFI, i, 1)
        
        Base31Reorg(i - 1) = InStr(1, Base31CharSet, Character, vbBinaryCompare) - 1
        If Base31Reorg(i - 1) = -1 Then
            'ERR: 552 VAL002 - Invalid UFI.  UFI may only contain the following characters '0123456789ACDEFGHJKMNPQRSTUVWXY'. Additional dashes or spaces for better readability are allowed.
            Err.Raise vbObjectError + 552, "VAL002 - Invalid UFI.", "UFI may only contain the following characters '" & Base31CharSet & "'. Additional dashes or spaces for better readability are allowed."
            Exit Function
        End If
        
        'VAL003 (part 1)
        WeightedSum = WeightedSum + Base31Reorg(i - 1) * i
    Next i
    
    'VAL003 - Checksum consistency (part 2)
    If WeightedSum Mod 31 <> 0 Then
        'ERR: 553 VAL003 - Invalid UFI.  UFI checksum inconsistency found.
        Err.Raise vbObjectError + 553, "VAL003 - Invalid UFI.", "UFI checksum inconsistency found."
        Exit Function
    End If
    
    
    'VAL004 - Consistency of group code and county encoding
    Dim Base31() As Variant 'undo character reorganization
    Base31 = UndoReorganizeBase31(Base31Reorg)
    
    Dim Payload As Variant
    Payload = UndoBase31(Base31)
       
    Dim PayloadBinary As String
    PayloadBinary = Dec2Bin(Payload, 74)
    
    Dim FormulationNumber As Long
    FormulationNumber = Bin2Dec(Left$(PayloadBinary, 28))
    
    Dim g As Long
    g = Bin2Dec(Mid$(PayloadBinary, 29, 4))
    
    Dim b As Long
    Dim C As Long
    Dim CountryCode As String
    
    Select Case g
        Case 0 'company code
            CountryCode = vbNullString
            
        Case 1
            CountryCode = "FR"
            
        Case 2
            CountryCode = "GB|XN"

        Case 3
            b = 1
            C = Bin2Dec(Mid$(PayloadBinary, 33, b))
            If C < 0 Or C > 1 Then
                'ERR: 555 VAL004B - Invalid UFI.  Inconsistency in country group code / number of bits / country code found.
                Err.Raise vbObjectError + 555, "VAL004 - Invalid UFI.", "Inconsistency in country group code / number of bits / country code found."
                Exit Function
            End If
            CountryCode = Array("LT", "SE")(C)
            
        Case 4
            b = 4
            C = Bin2Dec(Mid$(PayloadBinary, 33, b))
            If C < 0 Or C > 3 Then
                'ERR: 555 VAL004B - Invalid UFI.  Inconsistency in country group code / number of bits / country code found.
                Err.Raise vbObjectError + 555, "VAL004 - Invalid UFI.", "Inconsistency in country group code / number of bits / country code found."
                Exit Function
            End If
            CountryCode = Array("HR", "IT", "LV", "NL")(C)
            
        Case 5
            b = 7
            C = Bin2Dec(Mid$(PayloadBinary, 33, b))
            If C < 0 Or C > 22 Then
                'ERR: 555 VAL004B - Invalid UFI.  Inconsistency in country group code / number of bits / country code found.
                Err.Raise vbObjectError + 555, "VAL004 - Invalid UFI.", "Inconsistency in country group code / number of bits / country code found."
                Exit Function
            End If
            CountryCode = Array("BG", "CZ", "IE", "ES", "PL", "RO", "SK", "CY", "IS", "BE", "DE", "EE", "GR|EL", "NO", "PT", "AT", "DK", "FI", "HU", "LU", "MT", "SI", "LI")(C)
        
        Case Else
            'ERR: 554 VAL004A - Invalid UFI.  Invalid country group code found.
            Err.Raise vbObjectError + 554, "VAL004 - Invalid UFI.", "Invalid country group code found."
            Exit Function
    End Select


    'VAL005 - Version bit is zero
    If Right$(PayloadBinary, 1) <> "0" Then
        'ERR: 556 VAL005 - Invalid UFI.  Version bit has to be 0.
        Err.Raise vbObjectError + 556, "VAL004 - Invalid UFI.", "Version bit has to be 0."
        Exit Function
    End If
    
    
    'UFI is technically valid, try to decode VAT …
    Dim BinaryVAT As String
    BinaryVAT = Mid$(PayloadBinary, 33 + b, 41 - b)
    
    Dim VAT As String
    VAT = UndoBinaryVAT(CountryCode, BinaryVAT)
    
    
    'Return decoded data
    Dim decUFI As DecodedUFI
    Set decUFI = New DecodedUFI
    
    decUFI.CountryCode = CountryCode
    decUFI.VAT = VAT
    decUFI.FormulationNumber = FormulationNumber
    
    Set Decode = decUFI
End Function



Public Function IsValid(ByVal UFI As String) As Boolean
    On Error Resume Next
    '@Ignore FunctionReturnValueDiscarded
    Decode UFI 'throws error if not valid
    IsValid = (Err.Number = 0)
    On Error GoTo 0
End Function



Private Function ReorganizeBase31(ByVal Base31 As Variant) As Variant
    'Table 2-4: UFI characters reorganization tables
    'In   1  2  3  4  5  6  7  8  9 10 11 12 13 14 15
    'Out  6  5  4  8  3  9 10 11  2  1 12  7 13 14 15
    ReorganizeBase31 = Array(CLng(0), Base31(6), Base31(5), Base31(4), Base31(8), Base31(3), Base31(9), Base31(10), Base31(11), Base31(2), Base31(1), Base31(12), Base31(7), Base31(13), Base31(14), Base31(15))
End Function

Private Function UndoReorganizeBase31(ByVal Base31 As Variant) As Variant
    'Undo of Table 2-4: UFI characters reorganization tables
    'In   1  2  3  4  5  6  7  8  9 10 11 12 13 14 15
    'Out 10  9  5  3  2  1 12  4  6  7  8 11 13 14 15
    UndoReorganizeBase31 = Array(CLng(0), Base31(10), Base31(9), Base31(5), Base31(3), Base31(2), Base31(1), Base31(12), Base31(4), Base31(6), Base31(7), Base31(8), Base31(11), Base31(13), Base31(14), Base31(15))
End Function



Private Function GetBinaryVAT(ByVal CountryCode As String, ByVal VAT As String) As Variant
    '@Ignore UseMeaningfulName
    Dim d As Variant
    '@Ignore MultipleDeclarations, HungarianNotation, UseMeaningfulName
    Dim c1 As Long, Character1 As String
    '@Ignore MultipleDeclarations, HungarianNotation, UseMeaningfulName
    Dim c2 As Long, Character2 As String
    '@Ignore UseMeaningfulName
    Dim V As Variant
    Dim i As Long
    
    V = CDec(-1)
    
    Select Case UCase$(CountryCode)
        Case "AT"
            If IsMatchingPattern("^U[0-9]{8}$", VAT) Then
                V = Right$(VAT, 8)
            End If
        
        Case "BE"
            If IsMatchingPattern("^0[0-9]{9}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "BG"
            If IsMatchingPattern("^[0-9]{9,10}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "CY"
            If IsMatchingPattern("^[0-9]{8}[A-Z]$", VAT) Then
                d = CLng(Left$(VAT, 8))
                c1 = Asc(Right$(VAT, 1)) - 65
                V = CDec(c1 * 10 ^ 8) + d
            End If
       
        Case "CZ"
            If IsMatchingPattern("^[0-9]{8,10}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "DE", "EE", "GR", "EL", "NO", "PT"
            If IsMatchingPattern("^[0-9]{9}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "DK", "FI", "HU", "LU", "MT", "SI"
            If IsMatchingPattern("^[0-9]{8}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "ES"
            If IsMatchingPattern("^[0-9A-Z][0-9]{7}[0-9A-Z]$", VAT) Then
                d = CLng(Mid$(VAT, 2, 7))
                
                Character1 = Left$(VAT, 1)
                If IsNumeric(Character1) Then
                    c1 = CLng(Character1)
                Else
                    c1 = Asc(Character1) - 65 + 10
                End If
                
                Character2 = Right$(VAT, 1)
                If IsNumeric(Character2) Then
                    c2 = CLng(Character2)
                Else
                    c2 = Asc(Character2) - 65 + 10
                End If
                
                V = CDec((36 * c1 + c2) * 10 ^ 7 + d)
            End If
            
        Case "FR"
            If IsMatchingPattern("^[0-9A-Z]{2}[0-9]{9}$", VAT) Then
                d = CLng(Right$(VAT, 9))
                
                Character1 = Left$(VAT, 1)
                If IsNumeric(Character1) Then
                    c1 = CLng(Character1)
                Else
                    c1 = Asc(Character1) - 65 + 10
                End If
                
                Character2 = Mid$(VAT, 2, 1)
                If IsNumeric(Character2) Then
                    c2 = CLng(Character2)
                Else
                    c2 = Asc(Character2) - 65 + 10
                End If
                
                V = CDec((36 * c1 + c2) * 10 ^ 9 + d)
            End If
            
        Case "GB", "XN"
            If IsMatchingPattern("^[0-9]{9}([0-9]{3})?$", VAT) Then
                d = CDec(VAT)
                V = CDec(2 ^ 40 + d)
            ElseIf IsMatchingPattern("^[A-Z]{2}[0-9]{3}$", VAT) Then
                d = CLng(Right$(VAT, 3))
                c1 = Asc(Left$(VAT, 1)) - 65
                c2 = Asc(Mid$(VAT, 2, 1)) - 65
                V = CDec((26 * c1 + c2) * 10 ^ 3 + d)
            End If
            
        Case "IE"
            If IsMatchingPattern("^[0-9][A-Z*+][0-9]{5}[A-Z]$", VAT) Then
                d = CLng(Left$(VAT, 1) & Mid$(VAT, 3, 5))
                
                Character1 = Mid$(VAT, 2, 1)
                If Character1 = "+" Then
                    c1 = 26
                ElseIf Character1 = "*" Then
                    c1 = 27
                Else
                    c1 = Asc(Character1) - 65
                End If
                
                Character2 = Right$(VAT, 1)
                c2 = Asc(Character2) - 65
                
                V = CDec((26 * c1 + c2) * 10 ^ 6 + d)
            ElseIf IsMatchingPattern("^[0-9]{7}[A-Z]W?$", VAT) Or _
                   IsMatchingPattern("^[0-9]{7}[A-Z]{2}$", VAT) Then
                d = CLng(Left$(VAT, 7))
                
                Character1 = Mid$(VAT, 8, 1)
                c1 = Asc(Character1) - 65
                
                If Len(VAT) = 9 Then
                    Character2 = Right$(VAT, 1)
                    c2 = Asc(Character2) - 65
                Else
                    c2 = 0
                End If
                
                V = CDec(2 ^ 33 + ((26 * c2 + c1) * (10 ^ 7) + d))
            End If
            
        Case "IS"
            If IsMatchingPattern("^[A-Z0-9]{6}$", VAT) Then
                V = CDec(0)
                For i = 1 To Len(VAT)
                    Character1 = Mid$(VAT, i, 1)
                    If IsNumeric(Character1) Then
                        c1 = CLng(Character1)
                    Else
                        c1 = Asc(Character1) - 65 + 10
                    End If

                    V = V + CDec(36 ^ (6 - i) * c1)
                Next i
            End If
            
        Case "IT", "LV", "HR"
            If IsMatchingPattern("^[0-9]{11}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "LI"
            If IsMatchingPattern("^[0-9]{5}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "LT"
            If IsMatchingPattern("^[0-9]{9}$|^[0-9]{12}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "NL"
            If IsMatchingPattern("^[0-9]{9}B[0-9]{2}$", VAT) Then
                V = CDec(Left$(VAT, 9) & Right$(VAT, 2))
            End If
            
        Case "PL", "SK"
            If IsMatchingPattern("^[0-9]{10}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "RO"
            If IsMatchingPattern("^[0-9]{2,10}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case "SE"
            If IsMatchingPattern("^[0-9]{12}$", VAT) Then
                V = CDec(VAT)
            End If
            
        Case vbNullString 'VAT is company key
            V = CDec(VAT)
            
        Case Else
            'ERR: 513 Country Code does not exist.
            Err.Raise vbObjectError + 513, "UFIgenerator", "The country code '" & CountryCode & "' does not exist."
    End Select
    
    If V = -1 Then
        'ERR: 515 VAT is not valid.
        Err.Raise vbObjectError + 515, "UFIgenerator", "VAT '" & CountryCode & " " & VAT & "' is not valid."
    End If
    
    GetBinaryVAT = V
End Function

Private Function UndoBinaryVAT(ByVal CountryCode As String, ByVal BinaryVAT As String) As String
    '@Ignore UseMeaningfulName
    Dim d As Variant
    '@Ignore MultipleDeclarations, HungarianNotation, UseMeaningfulName
    Dim c1 As Long, Character1 As String
    '@Ignore MultipleDeclarations, HungarianNotation, UseMeaningfulName
    Dim c2 As Long, Character2 As String
    '@Ignore UseMeaningfulName
    Dim V As Variant
    '@Ignore MultipleDeclarations, UseMeaningfulName
    Dim i As Long, j As Long
    Dim FirstPart As Long
    Dim VAT As String
    
    V = Bin2Dec(BinaryVAT)
    
    Select Case UCase$(CountryCode)
        Case "AT"
            'For matching U[0-9]{8}
            VAT = "U" & Right$(String(8, "0") & CStr(V), 8)
            If Not IsMatchingPattern("^U[0-9]{8}$", VAT) Then VAT = vbNullString
        Case "BE"
            'For matching 0[0-9]{9}
            VAT = "0" & Right$(String(9, "0") & CStr(V), 9)
            
        Case "BG"
            'For matching [0-9]{9,10}
            VAT = CStr(V)
            If Len(VAT) < 9 Then
                VAT = Right$(String(9, "0") & VAT, 9)
            End If
            If Not IsMatchingPattern("^[0-9]{9,10}$", VAT) Then VAT = vbNullString
        
        Case "CY"
            'For matching [0-9]{8}[A-Z]
            d = CLng(Right$(V, 8))
            c1 = CLng(CStr(Left$(V, Len(V) - 8)))
            Character1 = Chr$(c1 + 65)
            VAT = Right$(String(8, "0") & CStr(d), 8) & Character1
            If Not IsMatchingPattern("^[0-9]{8}[A-Z]$", VAT) Then VAT = vbNullString
        
        Case "CZ"
            'For matching [0-9]{8,10}
            VAT = CStr(V)
            If Len(VAT) < 8 Then
                VAT = Right$(String(8, "0") & VAT, 8)
            End If
            If Not IsMatchingPattern("^[0-9]{8,10}$", VAT) Then VAT = vbNullString
        
        Case "DE", "EE", "GR", "EL", "GR|EL", "NO", "PT"
            'For matching [0-9]{9}
            VAT = Right$(String(9, "0") & CStr(V), 9)
            If Not IsMatchingPattern("^[0-9]{9}$", VAT) Then VAT = vbNullString
        
        Case "DK", "FI", "HU", "LU", "MT", "SI"
            'For matching [0-9]{8}
            VAT = Right$(String(8, "0") & CStr(V), 8)
            If Not IsMatchingPattern("^[0-9]{8}$", VAT) Then VAT = vbNullString
        
        Case "ES"
            'For matching [0-9A-Z][0-9]{7}[0-9A-Z]
            d = CLng(Right$(V, 7))
            FirstPart = CLng(CStr(Left$(V, Len(V) - 7)))
            
            'Brut force the possibility to find c1 and c2
            c1 = -1: c2 = -1
            i = 0
            Do While c1 = -1 And i < 36
                j = 0
                Do While c2 = -1 And j < 36
                    If 36 * i + j = FirstPart Then
                        c1 = i
                        c2 = j
                    End If
                    j = j + 1
                Loop
                i = i + 1
            Loop
            
            Character1 = IIf(c1 < 10, CStr(c1), Chr$(c1 + 65 - 10))
            Character2 = IIf(c2 < 10, CStr(c2), Chr$(c2 + 65 - 10))
            
            VAT = Character1 & Right$(String(7, "0") & CStr(d), 7) & Character2
            If Not IsMatchingPattern("^[0-9A-Z][0-9]{7}[0-9A-Z]$", VAT) Then VAT = vbNullString
        
        
        Case "FR"
            'For matching [0-9A-Z]{2}[0-9]{9}
            d = CLng(Right$(V, 9))
            FirstPart = CLng(CStr(Left$(V, Len(V) - 9)))
            
            'Brut force the possibility to find c1 and c2
            c1 = -1: c2 = -1
            i = 0
            Do While c1 = -1 And i < 36
                j = 0
                Do While c2 = -1 And j < 36
                    If 36 * i + j = FirstPart Then
                        c1 = i
                        c2 = j
                    End If
                    j = j + 1
                Loop
                i = i + 1
            Loop
            
            Character1 = IIf(c1 < 10, CStr(c1), Chr$(c1 + 65 - 10))
            Character2 = IIf(c2 < 10, CStr(c2), Chr$(c2 + 65 - 10))
            
            VAT = Character1 & Character2 & Right$(String(9, "0") & CStr(d), 9)
            If Not IsMatchingPattern("^[0-9A-Z]{2}[0-9]{9}$", VAT) Then VAT = vbNullString
        
        Case "GB", "XN", "GB|XN"
            d = CLngLng(V - 2 ^ 40)
            If d > 0 Then
                'For matching ([0-9]{9}([0-9]{3})?
                VAT = CStr(d)
                If Len(VAT) < 9 Then
                    VAT = Right$(String(9, "0") & VAT, 9)
                End If
                If Not IsMatchingPattern("^[0-9]{9}([0-9]{3})?$", VAT) Then VAT = vbNullString
            Else
                'For matching [A-Z]{2}[0-9]{3}
                d = CLng(Right$(V, 3))
                FirstPart = CLng(CStr(Left$(V, Len(V) - 3)))
                
                'Brut force the possibility to find c1 and c2
                c1 = -1: c2 = -1
                i = 0
                Do While c1 = -1 And i < 26
                    j = 0
                    Do While c2 = -1 And j < 26
                        If 26 * i + j = FirstPart Then
                            c1 = i
                            c2 = j
                        End If
                        j = j + 1
                    Loop
                    i = i + 1
                Loop
                
                Character1 = Chr$(c1 + 65)
                Character2 = Chr$(c2 + 65)
                
                VAT = Character1 & Character2 + Right$(String(3, "0") & CStr(d), 3)
                If Not IsMatchingPattern("^[A-Z]{2}[0-9]{3}$", VAT) Then VAT = vbNullString
            End If
            
        Case "HR", "IT", "LV"
            'For matching [0-9]{11}
            VAT = Right$(String(11, "0") & CStr(V), 11)
            If Not IsMatchingPattern("^[0-9]{11}$", VAT) Then VAT = vbNullString
        
        Case "IE"
            If Len(V) <= 9 Then
                'For matching [0-9][A-Z*+][0-9]{5}[A-Z]
                d = CLng(Right$(V, 6))
                FirstPart = CLng(CStr(Left$(V, Len(V) - 6)))
                
                'Brut force the possibility to find c1 and c2
                c1 = -1: c2 = -1
                i = 0
                Do While c1 = -1 And i < 28
                    j = 0
                    Do While c2 = -1 And j < 28
                        If 26 * i + j = FirstPart Then
                            c1 = i
                            c2 = j
                        End If
                        j = j + 1
                    Loop
                    i = i + 1
                Loop
                
                If c1 = 26 Then
                    Character1 = "+"
                ElseIf c1 = 27 Then
                    Character1 = "*"
                Else
                    Character1 = Chr$(c1 + 65)
                End If
                
                If c2 = 26 Then
                    Character2 = "+"
                ElseIf c2 = 27 Then
                    Character2 = "*"
                Else
                    Character2 = Chr$(c2 + 65)
                End If
                
                VAT = Left$(CStr(d), 1) & Character1 & Right$(CStr(d), 5) & Character2
                If Not IsMatchingPattern("^[0-9][A-Z*+][0-9]{5}[A-Z]$", VAT) Then VAT = vbNullString
            Else
                'For matching [0-9]{7}[A-Z]W? or [0-9]{7}[A-Z]{2}
                
                'V = CDec(2 ^ 33 + ((26 * c2 + c1) * (10 ^ 7) + d))
                V = V - 2 ^ 33
                d = CLng(Right$(V, 7))
                FirstPart = CLng(CStr(Left$(V, Len(V) - 7)))

                'Brut force the possibility to find c1 and c2
                c1 = -1: c2 = -1
                i = 0
                Do While c1 = -1 And i < 28
                    j = 0
                    Do While c2 = -1 And j < 28
                        If 26 * j + i = FirstPart Then
                            c1 = i
                            c2 = j
                        End If
                        j = j + 1
                    Loop
                    i = i + 1
                Loop
                
                Character1 = Chr$(c1 + 65)
                Character2 = IIf(c2 = 0, vbNullString, Chr$(c2 + 65))
                
                VAT = CStr(d) & Character1 & Character2
                If Not IsMatchingPattern("^[0-9]{7}[A-Z]W?$|^[0-9]{7}[A-Z]{2}$", VAT) Then VAT = vbNullString
            End If
            
        Case "IS"
            'For matching [A-Z0-9]{6}
            VAT = vbNullString
            Do While V > 0 'reverse sequence of factorial pow
                c1 = V Mod 36
                V = V \ 36
                Character1 = IIf(c1 < 10, CStr(c1), Chr$(c1 + 65 - 10))
                VAT = Character1 & VAT
            Loop
            
            If Not IsMatchingPattern("^[A-Z0-9]{6}$", VAT) Then VAT = vbNullString

        Case "LI"
            'For matching [0-9]{5}
            VAT = Right$(String(5, "0") & CStr(V), 5)
            If Not IsMatchingPattern("^[0-9]{5}$", VAT) Then VAT = vbNullString
            
        Case "LT"
            'For matching ([0-9]{9}|[0-9]{12})
            VAT = CStr(V)
            If Len(VAT) < 9 Then
                VAT = Right$(String(9, "0") & VAT, 9)
            ElseIf Len(VAT) > 9 And Len(VAT) < 12 Then
                VAT = Right$(String(12, "0") & VAT, 12)
            End If
            
            If Not IsMatchingPattern("(^[0-9]{9}$|^[0-9]{12}$)", VAT) Then VAT = vbNullString

        Case "NL"
            'For matching [0-9]{9}B[0-9]{2}
            VAT = Right$(String(11, "0") & CStr(V), 11)
            VAT = Left$(VAT, 9) & "B" & Right$(VAT, 2)
            If Not IsMatchingPattern("^[0-9]{9}B[0-9]{2}$", VAT) Then VAT = vbNullString
            
        Case "PL", "SK"
            'For matching [0-9]{10}
            VAT = Right$(String(10, "0") & CStr(V), 10)
            If Not IsMatchingPattern("^[0-9]{10}$", VAT) Then VAT = vbNullString
            
        Case "RO"
            'For matching [0-9]{2,10}
            VAT = CStr(V)
            If Len(VAT) < 2 Then
                VAT = Right$(String(2, "0") & VAT, 2)
            End If
            If Not IsMatchingPattern("^[0-9]{2,10}$", VAT) Then VAT = vbNullString
            
        Case "SE"
            'For matching [0-9]{12}
            VAT = Right$(String(12, "0") & CStr(V), 12)
            If Not IsMatchingPattern("^[0-9]{12}$", VAT) Then VAT = vbNullString
            
        Case vbNullString 'VAT is company key
            VAT = CStr(V)
            
        Case Else
            'ERR: 513 Country Code does not exist.
            Err.Raise vbObjectError + 513, "Invalid VAT", "The country code '" & CountryCode & "' does not exist."
    End Select
    
    If VAT = vbNullString Then
        'ERR: 516 Decoded VAT is invalid for the country code.
        Err.Raise vbObjectError + 516, "Invalid VAT", "Decoded VAT is invalid for the country code '" & CountryCode & "'."
    End If
    
    UndoBinaryVAT = VAT
End Function



'@Ignore UseMeaningfulName
Private Function GetBase31(ByVal Payload As Variant) As Variant
    Dim Base31(1 To 15) As Long

    Dim localPayload As Variant
    localPayload = Payload
    Dim i As Long

    Do While localPayload > 0
        Base31(15 - i) = ModLarge(localPayload, 31)
        localPayload = (localPayload - Base31(15 - i)) / 31
        i = i + 1
    Loop
    
    GetBase31 = Base31
End Function

Private Function UndoBase31(ByVal Base31 As Variant) As Variant
    Dim Payload As Variant
    Payload = CDec(0)
    
    Dim i As Long
    For i = 1 To 15
        Payload = CDec(Payload * 31 + Base31(i))
    Next i
    
    UndoBase31 = Payload
End Function



Private Function ModLarge(ByVal Value As String, ByVal Divisor As Long) As Long
    Dim i As Long
    For i = 1 To Len(Value)
        ModLarge = (ModLarge * 10 + CInt(Mid$(Value, i, 1))) Mod Divisor
    Next i
End Function



Private Function Bin2Dec(ByVal BinaryString As String) As Variant
    '@Ignore UseMeaningfulName
    Const TwoToThe48 As Double = 281474976710656#
    
    Dim i As Long
    For i = 0 To Len(BinaryString) - 1
        If i > 48 Then
            Bin2Dec = CDec(Bin2Dec) + Val(Mid$(BinaryString, Len(BinaryString) - i, 1)) * TwoToThe48 * CDec(2 ^ (i - 48))
        Else
            Bin2Dec = CDec(Bin2Dec) + Val(Mid$(BinaryString, Len(BinaryString) - i, 1)) * CDec(2 ^ i)
        End If
    Next i
    
    'If Len(Bin2Dec) > 10 Then Bin2Dec = CStr(Bin2Dec)
End Function



Private Function Dec2Bin(ByVal DecimalValue As Variant, Optional ByVal NumberOfBits As Variant) As String
    Dim localDecimalValue As Variant
    localDecimalValue = DecimalValue
    localDecimalValue = CDec(localDecimalValue)
    Do While localDecimalValue <> 0
        Dec2Bin = Trim$(Str$(localDecimalValue - 2 * Int(localDecimalValue / 2))) & Dec2Bin
        localDecimalValue = Int(localDecimalValue / 2)
    Loop
    
    If Not IsMissing(NumberOfBits) Then
        If Len(Dec2Bin) > NumberOfBits Then
            'ERR: 514 The number X does not fit in Y bits.
            Err.Raise vbObjectError + 514, "UFIgenerator", "The number " & localDecimalValue & " does not fit in " & NumberOfBits & " bits."
        Else
            Dec2Bin = Right$(String$(NumberOfBits, "0") & Dec2Bin, NumberOfBits)
        End If
    End If
End Function


Private Function IsMatchingPattern(ByVal Pattern As String, ByVal Value As String) As Boolean
    Dim RegExp As Object
    Set RegExp = CreateObject("vbscript.regexp")
    
    With RegExp
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = Pattern
    End With
    
    IsMatchingPattern = RegExp.Test(Value)
    Set RegExp = Nothing
End Function



Private Function GetCountryGroupCodeBinary(ByVal CountryCode As String) As String
    Dim ReturnValue As String
    
    Select Case UCase$(CountryCode)
        Case vbNullString 'company keys
            ReturnValue = "0000"   '0
            
        Case "FR"
            ReturnValue = "0001"   '1
            
        Case "GB", "XN"
            ReturnValue = "0010"   '2
            
        Case "LT", "SE"
            ReturnValue = "0011"   '3
            
        Case "HR", "IT", "LV", "NL"
            ReturnValue = "0100"   '4
            
        Case "BG", "CZ", "IE", "ES", "PL", "RO", "SK", "CY", "IS", "BE", "DE", "EE", "GR", "EL", "NO", "PT", "AT", "DK", "FI", "HU", "LU", "MT", "SI", "LI"
            ReturnValue = "0101"   '5
            
        Case Else
            'ERR: 513 Country code does not exist.
            Err.Raise vbObjectError + 513, "UFIgenerator", "The country code '" & CountryCode & "' does not exist."
    End Select
    
    GetCountryGroupCodeBinary = ReturnValue
End Function



Private Function GetCountryCodeBinary(ByVal CountryCode As String) As String
    Dim ReturnValue As String
    
    Select Case UCase$(CountryCode)
        Case vbNullString 'company keys
                          ReturnValue = vbNullString
            
        Case "FR", _
             "GB", "XN": ReturnValue = vbNullString
        
        Case "LT":        ReturnValue = "0"       '0
        Case "SE":        ReturnValue = "1"       '1
        
        Case "HR":        ReturnValue = "0000"    '0
        Case "IT":        ReturnValue = "0001"    '1
        Case "LV":        ReturnValue = "0010"    '2
        Case "NL":        ReturnValue = "0011"    '3
        
        Case "BG":        ReturnValue = "0000000" '0
        Case "CZ":        ReturnValue = "0000001" '1
        Case "IE":        ReturnValue = "0000010" '2
        Case "ES":        ReturnValue = "0000011" '3
        Case "PL":        ReturnValue = "0000100" '4
        Case "RO":        ReturnValue = "0000101" '5
        Case "SK":        ReturnValue = "0000110" '6
        Case "CY":        ReturnValue = "0000111" '7
        Case "IS":        ReturnValue = "0001000" '8
        Case "BE":        ReturnValue = "0001001" '9
        Case "DE":        ReturnValue = "0001010" '10
        Case "EE":        ReturnValue = "0001011" '11
        Case "GR", "EL":  ReturnValue = "0001100" '12
        Case "NO":        ReturnValue = "0001101" '13
        Case "PT":        ReturnValue = "0001110" '14
        Case "AT":        ReturnValue = "0001111" '15
        Case "DK":        ReturnValue = "0010000" '16
        Case "FI":        ReturnValue = "0010001" '17
        Case "HU":        ReturnValue = "0010010" '18
        Case "LU":        ReturnValue = "0010011" '19
        Case "MT":        ReturnValue = "0010100" '20
        Case "SI":        ReturnValue = "0010101" '21
        Case "LI":        ReturnValue = "0010110" '22
        
        Case Else
            'ERR: 513 Country code does not exist.
            Err.Raise vbObjectError + 513, "UFIgenerator", "The country code '" & CountryCode & "' does not exist."
    End Select
    
    GetCountryCodeBinary = ReturnValue
End Function


