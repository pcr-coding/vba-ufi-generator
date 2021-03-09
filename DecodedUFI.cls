VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecodedUFI"
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


Option Explicit
'@Folder("Unique Formula Identifier")

Private m_CountryCode As String
Private m_VAT As String
Private m_FormulationNumber As Long


Public Property Get CountryCode() As String
    CountryCode = m_CountryCode
End Property

Public Property Let CountryCode(ByVal Value As String)
     m_CountryCode = Value
End Property


Public Property Get VAT() As String
    VAT = m_VAT
End Property

Public Property Let VAT(ByVal Value As String)
    m_VAT = Value
End Property


Public Property Get FormulationNumber() As Long
    FormulationNumber = m_FormulationNumber
End Property

Public Property Let FormulationNumber(ByVal Value As Long)
    m_FormulationNumber = Value
End Property

