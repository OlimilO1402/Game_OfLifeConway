Attribute VB_Name = "MField"
Option Explicit '2008_08_06 Zeilen:  75
Public Enum ELifeForm
    LifeFormRectangle
    LifeFormCircle
End Enum
Public Type Point
    x As Long 'könnte auch genausogut Integer sein
    y As Long
End Type
Public Type Rectangle
    P1 As Point
    P2 As Point
End Type
Public Type TField
    FieldColor As Long    ' Farbe des Feldes
    LifeColor  As Long    ' Farbe des lebenden Individuums
    GittColor  As Long    ' Farbe des Gitters (= Farbe des Feldes)
    bFixPtSize As Boolean ' sind die Punkte alle gleich groß
                          ' -> hat Vorteil beim Klicken der Punkte,
                          '    das Raster wird gleichmäßig gezeichnet
    PointSize  As Double  ' Größe eines einzelnen Feldpunktes
    LifeForm   As ELifeForm
    Arr()      As Rectangle
End Type
'hat ausgedient:
'Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As Any, ByVal X As Long, ByVal Y As Long) As Long

Public Sub New_Field(this As TField, ByVal lSizeX As Long, ByVal lSizeY As Long)
    With this
        ReDim .Arr(0 To lSizeX, 0 To lSizeY)
    End With
End Sub
Public Sub CalcField(this As TField, aPB As PictureBox)
    Dim i As Long, j As Long
    Dim W As Double, H As Double
    Dim d As Double
    Dim s As Long
    Dim div As Double
    Dim ubX As Long, ubY As Long
    div = IIf(aPB.ScaleMode = vbTwips, Screen.TwipsPerPixelX, 1)
    With this
        ubX = UBound(.Arr, 1)
        ubY = UBound(.Arr, 2)
        W = (aPB.ScaleWidth / div) / (ubX + 1)
        H = (aPB.ScaleHeight / div) / (ubY + 1)
        .PointSize = Min(W, H)
        d = PointSize(this)
        For j = 0 To ubY
            For i = 0 To ubX
                With .Arr(i, j)
                    .P1.x = CLng(i * d)
                    .P1.y = CLng(j * d)
                    .P2.x = .P1.x + CLng(d) + 1
                    .P2.y = .P1.y + CLng(d) + 1
                End With
            Next
        Next
    End With
End Sub
Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Private Property Get PointSize(this As TField) As Double
    With this
        PointSize = IIf(.bFixPtSize, CDbl(CLng(.PointSize)), .PointSize)
    End With
End Property
Public Function GetIndexPoint(this As TField, ByVal InX As Single, ByVal InY As Single) As Point
    'gibt die Indices des Feldes zurück das geklickt wurde
    Dim i As Long, j As Long
    Dim d As Double: d = PointSize(this)
    With this
        GetIndexPoint.x = CLng((CDbl(InX) + d / 2) / d) - 1
        GetIndexPoint.y = CLng((CDbl(InY) + d / 2) / d) - 1
    End With
End Function
