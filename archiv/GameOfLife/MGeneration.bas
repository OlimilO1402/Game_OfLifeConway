Attribute VB_Name = "MGeneration"
Option Explicit '2008_07_31 Zeilen:  74
Public Type Point
    X As Long
    Y As Long
End Type
Public Type Rectangle
    P1 As Point
    P2 As Point
End Type
Public Type TField
    Arr() As Rectangle
End Type
Public Type TIndividual
    Life  As Boolean 'Byte
End Type

Public Type TGeneration
'beinhaltet das Array das VB verwalten darf, d.h.
'das Array und sein Datenbereich wird von VB angelegt
'und wieder gelöscht
    Arr() As TIndividual
End Type
Public Type TGenerationPtr
'stellt einen Zeiger auf die Generation dar,
'mit diesem Array wird gearbeitet, Es darf nicht von
'VB gelöscht werden es muß von Hand gelöscht werden
    pUDT       As TUDTPtr2D
    Generation As TGeneration
End Type
'Private Declare Function DrawRect Lib "gdi32" Alias "Rectangle" (ByVal hhdc As Long, ByVal Rect As Rectangle) As Long
Public Declare Function DrawRect Lib "gdi32" Alias "Rectangle" ( _
    ByVal hhdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
Private EmptyInd As TIndividual
'


Public Sub New_GenerationPtr(this As TGenerationPtr, ByVal szX As Long, ByVal szY As Long)
    With this
        Call New_UDTPtr2D(.pUDT, FADF_EMBEDDED Or FADF_RECORD, _
                          LenB(EmptyInd), szX, 0, szY, 0)
        Call GetMem4(ByVal VarPtr(.pUDT.pSA), ByVal ArrPtr(.Generation.Arr))
    End With
End Sub
Public Sub DeletePtr(this As TGenerationPtr)
    With this.Generation
        Call GetMem4(0&, ByVal ArrPtr(.Arr))
    End With
End Sub
Public Sub DrawGeneration(this As TGeneration, Field As TField, aPB As PictureBox)
    Dim i  As Long, j  As Long
    Dim c  As Long
    Dim hr As Long
    Dim hhdc As Long: hhdc = aPB.hDC
    With this
        For j = 0 To UBound(.Arr, 2)
            For i = 0 To UBound(.Arr, 1)
                If .Arr(i, j).Life Then
                    c = C_LifeColor
                Else
                    c = C_FieldColor
                End If
                aPB.FillColor = c
                With Field.Arr(i, j)
                    hr = DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                End With
            Next
        Next
    End With
End Sub
