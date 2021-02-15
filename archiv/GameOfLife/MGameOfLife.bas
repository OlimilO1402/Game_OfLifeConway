Attribute VB_Name = "MGameOfLife"
Option Explicit '2008_07_31 Zeilen: 192
Public Type TGameOfLife
    SizeX      As Long ' im Feld die Anzahl an Individuen in X-Ri
    SizeY      As Long ' im Feld die Anzahl an Individuen in Y-Ri
    Counter    As Long ' zählt die Generationen
    LifeTime   As Long ' die Verzögerung, wie lange eine Generation
                       ' bestehen bleiben soll in ms (für Sleep)
    bIsRunning As Boolean 'läuft der Algo?
    'wie stark soll das Feld am Anfang belegt sein in Prozent
    Density    As Integer '1-100
    pThis      As TGenerationPtr
    ThisG      As TGeneration
    pNext      As TGenerationPtr
    NextG      As TGeneration
    Field      As TField 'vorberechnete Koordinaten für's Zeichnen der Kästchen
    'im Algo werden die Zeiger auf die Arrays nach jedem Durchlauf vertauscht
End Type
Public Const C_LifeColor  As Long = vbRed '&H00
Public Const C_FieldColor As Long = &H4000&     'vbGreen '&H00
Private Declare Sub Sleep Lib "kernel32" (ByVal dwms As Long)

Public Sub New_GameOfLife(this As TGameOfLife, _
                          ByVal lSizeX As Long, _
                          ByVal lSizeY As Long, _
                          ByVal ldelaytime As Long)
    
    With this
        .LifeTime = ldelaytime
        .SizeX = lSizeX
        .SizeY = lSizeY
        Dim sx As Long: sx = .SizeX + 1
        Dim sy As Long: sy = .SizeY + 1
        
        ReDim .ThisG.Arr(0 To sx, 0 To sy)
        Call MGeneration.New_GenerationPtr(.pThis, sx + 1, sy + 1)
        .pThis.pUDT.pvData = VarPtr(this.ThisG.Arr(0, 0))
        
        ReDim .NextG.Arr(0 To sx, 0 To sy)
        Call MGeneration.New_GenerationPtr(.pNext, sx + 1, sy + 1)
        .pNext.pUDT.pvData = VarPtr(this.NextG.Arr(0, 0))
        
        ReDim .Field.Arr(0 To sx, 0 To sy)
    End With
End Sub
Public Sub Delete(this As TGameOfLife)
    With this
        Call MGeneration.DeletePtr(.pNext)
        Call MGeneration.DeletePtr(.pThis)
    End With
End Sub
Public Sub CalcField(this As TGameOfLife, aPB As PictureBox)
    Dim i As Long, j As Long
    Dim W As Long, H As Long, DW As Long, DH As Long
    Dim d As Double
    Dim s As Long
    With this
        W = (aPB.ScaleWidth / Screen.TwipsPerPixelX) / (.SizeX + 2)
        H = (aPB.ScaleHeight / Screen.TwipsPerPixelY) / (.SizeY + 2)
        d = Min(W, H)
        For j = 0 To .SizeY + 1
            For i = 0 To .SizeX + 1
                With .Field.Arr(i, j)
                    .P1.X = CLng(i * d)
                    .P1.Y = CLng(j * d)
                    .P2.X = .P1.X + CLng(d)
                    .P2.Y = .P1.Y + CLng(d)
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

Public Sub InitRandom(this As TGameOfLife, Optional ByVal iDensity As Integer = -1)
    Randomize
    Dim i As Long, j As Long
    Dim r As Double
    Dim b As Boolean
    With this
        If iDensity > 0 Then
            'idensity wurde angegeben
            .Density = iDensity
        Else
            'idensity wurde nicht angegeben
            If .Density = 0 Then .Density = 50
            iDensity = .Density
        End If
        r = 2 * iDensity / 100
        For j = 1 To UBound(.ThisG.Arr, 2) - 1
            For i = 1 To UBound(.ThisG.Arr, 1) - 1
                b = CBool(CInt(Rnd * 2 ^ r))
                .ThisG.Arr(i, j).Life = b
                .NextG.Arr(i, j).Life = b
            Next
        Next
    End With
End Sub
Public Sub DrawAll(this As TGameOfLife, aPB As PictureBox)
    aPB.FillStyle = vbFSSolid
    With this
        Call MGeneration.DrawGeneration(.pThis.Generation, .Field, aPB)
    End With
End Sub
Public Sub Run(this As TGameOfLife, aPB As PictureBox)
    Dim i As Long, j As Long
    Dim hhdc As Long: hhdc = aPB.hDC
    Dim n As Long, neighbours As Long
    Dim bDraw As Boolean
    Dim pTemp As Long
    Dim C As Long ' die Farbe
    With this
        Do While .bIsRunning
            For j = 1 To .SizeY
                For i = 1 To .SizeX
                    With .pThis.Generation
                        'die Anzahl der Nachbarn zählen
                        'den links oben prüfen
                        If .Arr(i - 1, j - 1).Life Then neighbours = neighbours + 1
                        'den darüber prüfen
                        If .Arr(i, j - 1).Life Then neighbours = neighbours + 1
                        'den rechts oben prüfen
                        If .Arr(i + 1, j - 1).Life Then neighbours = neighbours + 1
                        'den links prüfen
                        If .Arr(i - 1, j).Life Then neighbours = neighbours + 1
                        'den rechts prüfen
                        If .Arr(i + 1, j).Life Then neighbours = neighbours + 1
                        'den links unten prüfen
                        If .Arr(i - 1, j + 1).Life Then neighbours = neighbours + 1
                        'den darunter prüfen
                        If .Arr(i, j + 1).Life Then neighbours = neighbours + 1
                        'den rechts unten prüfen
                        If .Arr(i + 1, j + 1).Life Then neighbours = neighbours + 1
                        If .Arr(i, j).Life Then
                            If (neighbours < 2) Or (3 < neighbours) Then
                                'das Individuum stirbt
                                ' an Einsamkeit:       neighbours < 2
                                ' an Überbevölkerung:  neighbours > 3
                                bDraw = True
                                this.pNext.Generation.Arr(i, j).Life = False
                                C = MGameOfLife.C_FieldColor
                            Else
                                this.pNext.Generation.Arr(i, j).Life = True
                            End If
                        Else
                            If neighbours = 3 Then
                                'geboren wird wenn
                                ' genau drei Nachbarn
                                bDraw = True
                                this.pNext.Generation.Arr(i, j).Life = True
                                C = MGameOfLife.C_LifeColor
                            Else
                                this.pNext.Generation.Arr(i, j).Life = False
                            End If
                        End If
                        neighbours = 0
                    End With
                    If bDraw Then
                        bDraw = False
                        aPB.FillColor = C
                        With .Field.Arr(i, j)
                            Call DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                        End With
                    End If
                Next
            Next
            'die Generation zählen
            .Counter = .Counter + 1
            ' hier ein Trick damit man endlich weiß wofür der Zeiger überhaupt
            ' gut sein soll: einfach die Zeiger auf die Arraydaten vertauschen,
            ' hat den Vorteil dass nichts umkopiert werden muß, man spart sich
            ' haufenweise Copy-Arbeit
            pTemp = .pNext.pUDT.pvData
            .pNext.pUDT.pvData = .pThis.pUDT.pvData
            .pThis.pUDT.pvData = pTemp
            
            If (.Counter Mod 50) = 0 Then
                'damit man im Programm noch irgendwas machen kann
                'gibt es alle 50 Generationen ein Doevents
                'Form1.Caption = CStr(.Counter)
                Form1.SetCounter (.Counter)
                DoEvents
            End If
            Call Sleep(.LifeTime)
        Loop 'Do While .bIsRunning
    End With 'this
End Sub
