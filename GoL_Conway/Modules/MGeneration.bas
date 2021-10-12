Attribute VB_Name = "MGeneration"
Option Explicit '2008_08_05 Zeilen:  80
Public Type TIndividual
    Life  As Boolean 'Byte
End Type
Public Type TGeneration
'beinhaltet das Array das VB verwalten darf, d.h.
'das Array und sein Datenbereich wird von VB angelegt
'und wieder gel�scht
    Arr() As TIndividual
End Type
Public Type TGenerationPtr
'stellt einen Zeiger auf die Generation dar,
'mit diesem Array wird gearbeitet, Es darf nicht von
'VB gel�scht werden es mu� von Hand gel�scht werden
    pUDT       As TUDTPtr2D
    Generation As TGeneration
End Type
'Private Declare Function DrawRect Lib "gdi32" Alias "Rectangle" (ByVal hhdc As Long, ByVal Rect As Rectangle) As Long
Public Declare Function DrawRect Lib "gdi32" Alias "Rectangle" ( _
    ByVal hhdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long
    
Public Declare Function DrawPixel Lib "gdi32.dll" Alias "SetPixel" ( _
    ByVal hhdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal crColor As Long _
) As Long
Public Declare Function DrawCircle Lib "gdi32.dll" Alias "Ellipse" ( _
    ByVal hhdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long

Private Declare Function DrawRC Lib "user32.dll" Alias "CallWindowProcA" ( _
    ByVal pFnc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" ( _
    ByVal lpModuleName As String _
) As Long
Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Public Declare Sub RtlZeroMemory Lib "kernel32" ( _
    ByRef pDst As Any, _
    ByVal bytlength As Long)

Private EmptyInd As TIndividual
Private LenBIndividual As Long
'
Public Sub New_Generation(this As TGeneration, ByVal szX As Long, ByVal szY As Long)
    With this
        ReDim .Arr(0 To szX, 0 To szY)
    End With
End Sub
Public Sub New_GenerationPtr(this As TGenerationPtr, ByVal szX As Long, ByVal szY As Long)
    LenBIndividual = LenB(EmptyInd)
    With this
        Call New_UDTPtr2D(.pUDT, FADF_EMBEDDED Or FADF_RECORD, _
                          LenBIndividual, szX + 1, 0, szY + 1, 0)
        Call GetMem4(ByVal VarPtr(.pUDT.pSA), ByVal ArrPtr(.Generation.Arr))
    End With
End Sub
Public Sub AssignPtr(this As TGenerationPtr, FromG As TGeneration)
    With this.pUDT
        .pvData = VarPtr(FromG.Arr(0, 0))
    End With
End Sub
Public Sub DeletePtr(this As TGenerationPtr)
    With this.Generation
        Call GetMem4(0&, ByVal ArrPtr(.Arr))
    End With
End Sub
Public Sub Clear(this As TGeneration)
    'Cleard einen Bereich der Generation
    With this
        Call RtlZeroMemory(.Arr(0, 0), LenBIndividual * (UBound(.Arr, 1) + 1) * (UBound(.Arr, 2) + 1))
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
                    c = Field.LifeColor
                Else
                    c = Field.FieldColor
                End If
                aPB.FillColor = c
                With Field.Arr(i, j)
                    If Field.LifeForm = LifeFormRectangle Then
                        If Not Field.bFixPtSize Then
                            hr = DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                        Else
                            If Field.PointSize <= 1 Then
                                hr = DrawPixel(hhdc, .P1.X, .P1.Y, c)
                            Else
                                hr = DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                            End If
                        End If
                    ElseIf Field.LifeForm = LifeFormCircle Then
                        hr = DrawCircle(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                    End If
                End With
            Next
        Next
    End With
End Sub
Public Sub DrawIndividual(this As TGeneration, Field As TField, aPB As PictureBox, pt As Point)
    Dim c  As Long
    Dim hr As Long
    Dim hhdc As Long: hhdc = aPB.hDC
    With this
    'halt, hier index outofbounds
    'erst �berpr�fen ob der geklickte
    'Punkt innerhalb des Spielfelds liegt
        If .Arr(pt.X, pt.Y).Life Then
            c = Field.LifeColor
        Else
            c = Field.FieldColor
        End If
        aPB.FillColor = c
        With Field.Arr(pt.X, pt.Y)
            If Field.LifeForm = LifeFormRectangle Then
                If Not Field.bFixPtSize Then
                    hr = DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                Else
                    If Field.PointSize <= 1 Then
                        hr = DrawPixel(hhdc, .P1.X, .P1.Y, c)
                    Else
                        hr = DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                    End If
                End If
            ElseIf Field.LifeForm = LifeFormCircle Then
                hr = DrawCircle(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
            End If
        End With
    End With
End Sub
