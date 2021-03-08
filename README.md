# vba-ufi-generator
[![License: MIT](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://opensource.org/licenses/gpl-3.0)

VBA Generator/decoder for [ECHA](https://ufi.echa.europa.eu/#/create) UFIs (Unique Formula Identifier)  
:books: [UFI Developer manual](https://poisoncentres.echa.europa.eu/documents/22284544/22295820/ufi_developers_manual_en.pdf)

## Syntax
### Generate function
Returns a **String** representing the encoded UFI including dashes.

    Generate(CountryCode, VAT, FormulationNumber)
    
| Argument          | Type   | Description                                       |
| :---------------- | :----- | :------------------------------------------------ |
| CountryCode       | String | Country Code of the VAT number eg. `"AT"`         |
| VAT               | String | VAT number without country code eg. `"U12345678"` |
| FormulationNumber | Long   | Numeric formulation number eg. `178956970`        |

### Decode function
Returns a **DecodedUFI** object of the decoded UFI.

    Decode(UFI)
    
| Argument | Type   | Description                                                |
| :------- | :----- | :--------------------------------------------------------- |
| UFI      | String | UFI (with or without dashes) eg. `"C23S-PQ2V-AMH9-VVRF"`   |

    
### IsValid function
Returns a **True** if the UFI is valid.

    IsValid(UFI)
    
| Argument | Type   | Description                                                |
| :------- | :----- | :--------------------------------------------------------- |
| UFI      | String | UFI (with or without dashes) eg. `"C23S-PQ2V-AMH9-VVRF"`   |


## Examples

### Generate UFI

```vba
Public Sub ExampleGenerate()
    Dim UFIgen As New UFIgenerator
    
    Debug.Print UFIgen.Generate("AT", "U12345678", 178956970)  'C23S-PQ2V-AMH9-VVRF
End Sub
```

### Decode UFI

```vba
Public Sub ExampleDecode()
    Dim UFIgen As New UFIgenerator
    
    Dim UFI As DecodedUFI
    Set UFI = UFIgen.Decode("C23S-PQ2V-AMH9-VVRF")

    Debug.Print UFI.CountryCode, UFI.VAT, UFI.FormulationNumber  'AT U12345678      178956970
End Sub
```
Decoder includes validation of the UFI and will throw errors if decoding fails or UFI is not valid.

### Validate UFI

```vba
Public Sub ExampleValidate()
    Dim UFIgen As New UFIgenerator
    
    If UFIgen.IsValid("C23S-PQ2V-AMH9-VVRF") Then
        Debug.Print "Valid"
    Else
        Debug.Print "Invalid"
    End If
End Sub
```
Validator will only return `True`/`False` and does not throw errors.
