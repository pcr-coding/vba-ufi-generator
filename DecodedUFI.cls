VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecodedUFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

