<div align="center">

## HSL\<\-\>RGB and Color Manipulation Routines


</div>

### Description

Routines to convert between Hue-Saturation-Luminescence values an Red-Green-Blue color values (Converted from C++). Also several unique routines using these functions to manipulate color such as Brighten, Invert, PhotoNegative, Blend, Tint, etc.

***NOTE: I have reposted this with the original sorce (now complete). I am working on and will soon submit this in class form, once comments are added to the source.***
 
### More Info
 
Each routine is different, but in general they pass RGB colors as Longs (standard VB color values) and HSL values in a user-defined type called HSLCol, which contains three integer members Hue, Sat & Lum.

The "Windows API/Global Declarations" section contains the code for modHSL.bas; which is all you really need. The code listed in "Source Code" is a demonstration routine. Create a form with a picture control called Picture1 and a command button called Command1. Load a bitmap into the picture control and set Autosize to True. Have fun expirimenting!

See Inputs


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Redding \- Blue Knot Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-redding-blue-knot-software.md)
**Level**          |Advanced
**User Rating**    |4.3 (77 globes from 18 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-redding-blue-knot-software-hsl-rgb-and-color-manipulation-routines__1-2144/archive/master.zip)

### API Declarations

```
Option Explicit
'
' modHSL.bas
' HSL/RGB + Color Manipulation routines
'
'Portions of this code marked with *** are converted from
'C/C++ routines for RGB/HSL conversion found in the
'Microsoft Knowledge Base (PD sample code):
'http://support.microsoft.com/support/kb/articles/Q29/2/40.asp
'In addition to the language conversion, some internal
'calculations have been modified and converted to FP math to
'reduce rounding errors.
'Conversion to VB and original code by
'Dan Redding (bwsoft@revealed.net)
'http://home.revealed.net/bwsoft
'Free to use, please give proper credit
Const HSLMAX As Integer = 240 '***
 'H, S and L values can be 0 - HSLMAX
 '240 matches what is used by MS Win;
 'any number less than 1 byte is OK;
 'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
 'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
 'Hue is undefined if Saturation = 0 (greyscale)
Public Type HSLCol 'Datatype used to pass HSL Color values
 Hue As Integer
 Sat As Integer
 Lum As Integer
End Type
Public Function RGBRed(RGBCol As Long) As Integer
'Return the Red component from an RGB Color
 RGBRed = RGBCol And &HFF
End Function
Public Function RGBGreen(RGBCol As Long) As Integer
'Return the Green component from an RGB Color
 RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function
Public Function RGBBlue(RGBCol As Long) As Integer
'Return the Blue component from an RGB Color
 RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Private Function iMax(a As Integer, B As Integer) As Integer
'Return the Larger of two values
 iMax = IIf(a > B, a, B)
End Function
Private Function iMin(a As Integer, B As Integer) As Integer
'Return the smaller of two values
 iMin = IIf(a < B, a, B)
End Function
Public Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value
Dim R As Integer, G As Integer, B As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
 BDelta As Double
Dim H As Double, S As Double, L As Double
Dim cMinus As Long, cPlus As Long
 R = RGBRed(RGBCol)
 G = RGBGreen(RGBCol)
 B = RGBBlue(RGBCol)
 cMax = iMax(iMax(R, G), B) 'Highest and lowest
 cMin = iMin(iMin(R, G), B) 'color values
 cMinus = cMax - cMin 'Used to simplify the
 cPlus = cMax + cMin 'calculations somewhat.
 'Calculate luminescence (lightness)
 L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
 If cMax = cMin Then 'achromatic (r=g=b, greyscale)
 S = 0 'Saturation 0 for greyscale
 H = UNDEFINED 'Hue undefined for greyscale
 Else
 'Calculate color saturation
 If L <= (HSLMAX / 2) Then
 S = ((cMinus * HSLMAX) + 0.5) / cPlus
 Else
 S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
 End If
 'Calculate hue
 RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
 GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
 BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
 Select Case cMax
 Case CLng(R)
 H = BDelta - GDelta
 Case CLng(G)
 H = (HSLMAX / 3) + RDelta - BDelta
 Case CLng(B)
 H = ((2 * HSLMAX) / 3) + GDelta - RDelta
 End Select
 If H < 0 Then H = H + HSLMAX
 End If
 RGBtoHSL.Hue = CInt(H)
 RGBtoHSL.Lum = CInt(L)
 RGBtoHSL.Sat = CInt(S)
End Function
Public Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
  Dim R As Long, G As Long, B As Long
  Dim H As Long, L As Long, S As Long
  Dim Magic1 As Integer, Magic2 As Integer
  H = HueLumSat.Hue
  L = HueLumSat.Lum
  S = HueLumSat.Sat
  If S = 0 Then 'Greyscale
    R = (L * RGBMAX) / HSLMAX 'luminescence,
        'converted to the proper range
    G = R 'All RGB values same in greyscale
    B = R
    If H <> UNDEFINED Then
      'This is technically an error.
      'The RGBtoHSL routine will always return
      'Hue = UNDEFINED (in this case 160)
      'when Sat = 0.
      'if you are writing a color mixer and
      'letting the user input color values,
      'you may want to set Hue = UNDEFINED
      'in this case.
    End If
  Else
    'Get the "Magic Numbers"
    If L <= HSLMAX / 2 Then
      Magic2 = (L * (HSLMAX + S) + _
        (HSLMAX / 2)) / HSLMAX
    Else
      Magic2 = L + S - ((L * S) + _
        (HSLMAX / 2)) / HSLMAX
    End If
    Magic1 = 2 * L - Magic2
    'get R, G, B; change units from HSLMAX range
    'to RGBMAX range
    R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) _
      * RGBMAX + (HSLMAX / 2)) / HSLMAX
    G = (HuetoRGB(Magic1, Magic2, H) _
      * RGBMAX + (HSLMAX / 2)) / HSLMAX
    B = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) _
      * RGBMAX + (HSLMAX / 2)) / HSLMAX
  End If
  HSLtoRGB = RGB(CInt(R), CInt(G), CInt(B))
End Function
Private Function HuetoRGB(mag1 As Integer, mag2 As Integer, _
  Hue As Long) As Long '***
'Utility function for HSLtoRGB
'Range check
  If Hue < 0 Then
    Hue = Hue + HSLMAX
  ElseIf Hue > HSLMAX Then
    Hue = Hue - HSLMAX
  End If
  'Return r, g, or b value from parameters
  Select Case Hue 'Values get progressively larger.
        'Only the first true condition will execute
    Case Is < (HSLMAX / 6)
      HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
        (HSLMAX / 12)) / (HSLMAX / 6)))
    Case Is < (HSLMAX / 2)
      HuetoRGB = mag2
    Case Is < (HSLMAX * 2 / 3)
      HuetoRGB = (mag1 + (((mag2 - mag1) * _
        ((HSLMAX * 2 / 3) - Hue) + _
        (HSLMAX / 12)) / (HSLMAX / 6)))
    Case Else
      HuetoRGB = mag1
  End Select
End Function
'
' The following are individual functions
' that use the HSL/RGB routines
' This is not intended to be a comprehensive library,
' just a sampling to demonstrate how to use the routines
' and what kind of things are possible
'
Public Function ContrastingColor(RGBCol As Long) As Long
'Returns Black or White, whichever will show up better
'on the specified color
'Useful for setting label forecolors with transparent
'backgrounds (send it the form backcolor - RGB value, not
'system value!)
Dim HSL As HSLCol
  HSL = RGBtoHSL(RGBCol)
  If HSL.Lum > HSLMAX / 2 Then ContrastingColor = 0 _
    Else: ContrastingColor = &HFFFFFF
End Function
'Color adjustment routines
'These accept a color and return a modified color.
'Perhaps the most common use might be to apply a process
'to an image pixel by pixel
Public Function Greyscale(RGBColor As Long) As Long
'Returns the achromatic version of a color
Dim HSL As HSLCol
  HSL = RGBtoHSL(RGBColor)
  HSL.Sat = 0
  HSL.Hue = UNDEFINED
  Greyscale = HSLtoRGB(HSL)
End Function
Public Function Tint(RGBColor As Long, Hue As Integer)
'Changes the Hue of a color to a specified Hue
'For example, changing all the pixels in a picture to
'a hue of 80 would tint the picture green
Dim HSL As HSLCol
  If Hue < 0 Then
    Hue = 0
  ElseIf Hue > HSLMAX Then
    Hue = HSLMAX
  End If
  HSL = RGBtoHSL(RGBColor)
  HSL.Hue = Hue
  Tint = HSLtoRGB(HSL)
End Function
Public Function Brighten(RGBColor As Long, Percent As Single)
'Lightens the color by a specifie percent, given as a Single
'(10% = .10)
Dim HSL As HSLCol, L As Long
  If Percent <= 0 Then
    Brighten = RGBColor
    Exit Function
  End If
  HSL = RGBtoHSL(RGBColor)
  L = HSL.Lum + (HSL.Lum * Percent)
  If L > HSLMAX Then L = HSLMAX
  HSL.Lum = L
  Brighten = HSLtoRGB(HSL)
End Function
Public Function Darken(RGBColor As Long, Percent As Single)
'Darkens the color by a specifie percent, given as a Single
Dim HSL As HSLCol, L As Long
  If Percent <= 0 Then
    Darken = RGBColor
    Exit Function
  End If
  HSL = RGBtoHSL(RGBColor)
  L = HSL.Lum - (HSL.Lum * Percent)
  If L < 0 Then L = 0
  HSL.Lum = L
  Darken = HSLtoRGB(HSL)
End Function
Public Function ReverseLight(RGBColor As Long) As Long
'Make dark colors light and vice versa without changing Hue
'or saturation
Dim HSL As HSLCol
  HSL = RGBtoHSL(RGBColor)
  HSL.Lum = HSLMAX - HSL.Lum
  ReverseLight = HSLtoRGB(HSL)
End Function
Public Function ReverseColor(RGBColor As Long) As Long
'Swap colors without changing saturation or luminescence
Dim HSL As HSLCol
  HSL = RGBtoHSL(RGBColor)
  HSL.Hue = HSLMAX - HSL.Hue
  ReverseColor = HSLtoRGB(HSL)
End Function
Public Function CycleColor(RGBColor As Long) As Long
'Cycle colors thru a 12 stage pattern without changing
'saturation or luminescence
Dim HSL As HSLCol, H As Long
  HSL = RGBtoHSL(RGBColor)
  H = HSL.Hue + (HSLMAX / 12)
  If H > HSLMAX Then H = H - HSLMAX
  HSL.Hue = H
  CycleColor = HSLtoRGB(HSL)
End Function
Public Function Blend(RGB1 As Long, RGB2 As Long, _
  Percent As Single) As Long
'This one doesn't really use the HSL routines, just the
'RGB Component routines. I threw it in as a bonus ;)
'Takes two colors and blends them according to a
'percentage given as a Single
'For example, .3 will return a color 30% of the way
'between the first color and the second.
'.5, or 50%, will be an even blend (halfway)
Dim R As Integer, R1 As Integer, R2 As Integer, _
  G As Integer, G1 As Integer, G2 As Integer, _
  B As Integer, B1 As Integer, B2 As Integer
  If Percent >= 1 Then
    Blend = RGB2
    Exit Function
  ElseIf Percent <= 0 Then
    Blend = RGB1
    Exit Function
  End If
  R1 = RGBRed(RGB1)
  R2 = RGBRed(RGB2)
  G1 = RGBGreen(RGB1)
  G2 = RGBGreen(RGB2)
  B1 = RGBBlue(RGB1)
  B2 = RGBBlue(RGB2)
  R = ((R2 * Percent) + (R1 * (1 - Percent)))
  G = ((G2 * Percent) + (G1 * (1 - Percent)))
  B = ((B2 * Percent) + (B1 * (1 - Percent)))
  Blend = RGB(R, G, B)
End Function
```


### Source Code

```
Private Sub Command1_Click()
Dim i%, j%, R&, c&
'Simple routine to demonstrate color manipulation
'in a picture. Not fast but it works.
'Picture1 must contain an image and be Autosized to it.
'(Point will return -1 for pixels outside an image, and
'this is invalid)
For i = 0 To (Picture1.ScaleWidth - Screen.TwipsPerPixelX) _
 Step Screen.TwipsPerPixelX
 For j = 0 To (Picture1.ScaleHeight - Screen.TwipsPerPixelY) _
 Step Screen.TwipsPerPixelY
 c = Picture1.Point(i, j)
 If c >= 0 Then
 'Point will return -1 for pixels outside an image
 c = PhotoNegative(c) 'Substitute any color routine here
 'c = Tint(c,80)
 'c = Brighten(c,0.1)
 'c = Greyscale(c)
 'etc.
 Picture1.PSet (i, j), c
 End If
 Next j
Next i
End Sub
```

