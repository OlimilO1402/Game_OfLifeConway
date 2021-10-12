Attribute VB_Name = "MGameOfLife"
Option Explicit '2008_08_05 Zeilen: 205
Public Type TGameOfLife
    SizeX      As Long ' im Feld die Anzahl an Individuen in X-Ri
    SizeY      As Long ' im Feld die Anzahl an Individuen in Y-Ri
    
    Counter    As Long  ' zählt die Generationen
    GenTilDoEv As Long  ' Generationen zwischen Doevents
    LifeTime   As Long  ' die Verzögerung, wie lange eine Generation
                        ' bestehen bleiben soll in ms (für Sleep)
    bIsRunning As Boolean 'läuft der Run-Algo?
    'wie stark soll das Feld am Anfang belegt sein in Prozent
    Density    As Integer '1-100 %
    LifeRule   As LifeRule
    Field      As TField 'vorberechnete Koordinaten für's Zeichnen der Kästchen
    pThis      As TGenerationPtr
    ThisG      As TGeneration
    pNext      As TGenerationPtr
    NextG      As TGeneration
    'im Algo werden die Zeiger auf die Arrays nach jedem Durchlauf vertauscht
End Type
Private Declare Sub Sleep Lib "kernel32" (ByVal dwms As Long)

Public Sub New_GameOfLife(this As TGameOfLife, _
                          ByVal lSizeX As Long, _
                          ByVal lSizeY As Long, _
                          ByVal sRule As String)
    
    With this
        Dim szX As Long: szX = lSizeX + 1
        Dim szY As Long: szY = lSizeY + 1
        .SizeX = lSizeX
        .SizeY = lSizeY
        .LifeRule = MLifeRule.New_LifeRule(sRule)
        Call MField.New_Field(.Field, szX, szY)
        Call MGeneration.New_Generation(.ThisG, szX, szY)
        Call MGeneration.New_GenerationPtr(.pThis, szX, szY)
        Call MGeneration.AssignPtr(.pThis, .ThisG)
        Call MGeneration.New_Generation(.NextG, szX, szY)
        Call MGeneration.New_GenerationPtr(.pNext, szX, szY)
        Call MGeneration.AssignPtr(.pNext, .NextG)
    End With
End Sub
Public Sub Clear(this As TGameOfLife)
    With this
        Call MGeneration.Clear(.ThisG)
        Call MGeneration.Clear(.NextG)
    End With
End Sub
Public Sub SwitchLife(this As TGameOfLife, IndexPt As Point, ByVal kill As Boolean)
    'wird aufgerufen beim Klicken von Punkten
    With this
        If IsPointInside(this, IndexPt) Then
            With .pThis.Generation.Arr(IndexPt.X, IndexPt.Y)
                .Life = Not .Life And kill
            End With
        End If
    End With
End Sub
Public Sub Delete(this As TGameOfLife)
    'wird bei Programmende aufgerufen
    With this
        Call MGeneration.DeletePtr(.pNext)
        Call MGeneration.DeletePtr(.pThis)
    End With
End Sub
Public Function IsPointInside(this As TGameOfLife, pt As Point) As Boolean
    With this
        IsPointInside = (0 < pt.X And pt.X <= .SizeX) And _
                        (0 < pt.Y And pt.Y <= .SizeY)
    End With
End Function
Public Sub InitRandom(this As TGameOfLife, _
                      Optional ByVal iDensity As Integer = -1, _
                      Optional ByVal bEdgelife As Boolean)
    'erzeugt eine zufällig verteilte Generation mit der angegebenen Bevölkerungsdichte
    Randomize
    Dim i As Long, j As Long, k As Long
    Dim r As Double
    Dim b As Boolean
    Dim ubX As Long, ubY As Long
    With this
        .Counter = 0
        If iDensity > 0 Then
            'iDensity wurde angegeben
            .Density = iDensity
        Else
            'iDensity wurde nicht angegeben
            If .Density = 0 Then .Density = 50
            iDensity = .Density
        End If
        r = iDensity / 100
        ubX = UBound(.ThisG.Arr, 1)
        ubY = UBound(.ThisG.Arr, 2)
        For j = 0 To ubY ' - 1
            For i = 0 To ubX ' - 1
                If i = 0 Or i = ubX Or _
                   j = 0 Or j = ubY Then
                    b = bEdgelife
                Else
                    b = (Rnd < r) 'Danke Henrik Ilgen das ist einfach genial
                End If
                .ThisG.Arr(i, j).Life = b
                .NextG.Arr(i, j).Life = b
            Next
        Next
    End With
End Sub

Public Sub DeleteRandom(this As TGameOfLife, _
                        Optional ByVal iDensity As Integer = -1, _
                        Optional ByVal bEdgelife As Boolean)
    'löscht in zufälliger Verteilung mit der angegebenen Dichte
    Randomize
    Dim i As Long, j As Long, k As Long
    Dim r As Double
    Dim b As Boolean
    Dim ubX As Long, ubY As Long
    With this
        .Counter = 0
        If iDensity > 0 Then
            'iDensity wurde angegeben
            .Density = iDensity
        Else
            'iDensity wurde nicht angegeben
            If .Density = 0 Then .Density = 50
            iDensity = .Density
        End If
        r = iDensity / 100
        ubX = UBound(.ThisG.Arr, 1)
        ubY = UBound(.ThisG.Arr, 2)
        For j = 0 To ubY ' - 1
            For i = 0 To ubX ' - 1
                If i = 0 Or i = ubX Or _
                   j = 0 Or j = ubY Then
                    b = bEdgelife
                Else
                    b = (Rnd < r) 'Danke Henrik Ilgen das ist einfach genial
                End If
                If b Then
                    .ThisG.Arr(i, j).Life = False
                End If
                '.NextG.Arr(i, j).Life = b
            Next
        Next
    End With
End Sub

'Private Function RandomBetween(ByVal dblmin As Double, ByVal dblmax As Double) As Double
'    RandomBetween = dblmin + (dblmax - dblmin) * Rnd
'End Function
Public Sub DrawAll(this As TGameOfLife, aPB As PictureBox)
    aPB.FillStyle = vbFSSolid
    With this
        Call MGeneration.DrawGeneration(.pThis.Generation, .Field, aPB)
    End With
End Sub
Public Sub Run(this As TGameOfLife, aPB As PictureBox)
    ' diese Funktion schneller machen?
    ' siehe Trick 1 - 5
    Dim i As Long, j As Long
    Dim hhdc As Long: hhdc = aPB.hDC
    Dim hr   As Long
    Dim n    As Long, neighbours As Long, inb As ERule
    Dim bDraw As Boolean
    Dim pTemp As Long
    Dim c     As Long ' die Farbe
    Dim bRect As Boolean 'True: zeichne Rectangles, False: zeichne Kreise
    Dim rSurvive As ERule
    Dim rNewBorn As ERule
    Dim bb(0 To 8) As ERule
    For i = 0 To 8
        bb(i) = 2 ^ i
    Next
    Call MProcess.SetProcessPriority(THREAD_PRIORITY_ABOVE_NORMAL, HIGH_PRIORITY_CLASS)
    
    With this
        If .Counter = 2147483647 Then .Counter = 0
        rSurvive = .LifeRule.RuleSurvive
        rNewBorn = .LifeRule.RuleNewBorn
        bRect = (.Field.LifeForm = LifeFormRectangle)
        Do While .bIsRunning
            For j = 1 To .SizeY
                For i = 1 To .SizeX
                    With .pThis.Generation
                        ' Trick 1:
                        ' wir spielen Compiler und inlinen die Funktion CheckNB.
                        ' die Anzahl der Nachbarn zählen, hier innerhalb von Run
                        '
                        ' den links oben prüfen
                        If .Arr(i - 1, j - 1).Life Then neighbours = neighbours + 1
                        ' den darüber prüfen
                        If .Arr(i, j - 1).Life Then neighbours = neighbours + 1
                        ' den rechts oben prüfen
                        If .Arr(i + 1, j - 1).Life Then neighbours = neighbours + 1
                        ' den links prüfen
                        If .Arr(i - 1, j).Life Then neighbours = neighbours + 1
                        ' den rechts prüfen
                        If .Arr(i + 1, j).Life Then neighbours = neighbours + 1
                        ' den links unten prüfen
                        If .Arr(i - 1, j + 1).Life Then neighbours = neighbours + 1
                        ' den darunter prüfen
                        If .Arr(i, j + 1).Life Then neighbours = neighbours + 1
                        ' den rechts unten prüfen
                        If .Arr(i + 1, j + 1).Life Then neighbours = neighbours + 1
                        'Select Case neighbours
                        'Case 0: inb = ERule.RNB0
                        'Case 1: inb = ERule.RNB1
                        'Case 2: inb = ERule.RNB2
                        'Case 3: inb = ERule.RNB3
                        'Case 4: inb = ERule.RNB4
                        'Case 5: inb = ERule.RNB5
                        'Case 6: inb = ERule.RNB6
                        'Case 7: inb = ERule.RNB7
                        'Case 8: inb = ERule.RNB8
                        'End Select
                        inb = bb(neighbours) 'lokales Array
                        If .Arr(i, j).Life Then
                            If rSurvive And inb Then
                                this.pNext.Generation.Arr(i, j).Life = True
                            Else
                                bDraw = True
                                this.pNext.Generation.Arr(i, j).Life = False
                                c = this.Field.FieldColor
                            End If
                        Else
                            If rNewBorn And inb Then
                                bDraw = True
                                this.pNext.Generation.Arr(i, j).Life = True
                                c = this.Field.LifeColor
                            Else
                                this.pNext.Generation.Arr(i, j).Life = False
                            End If
                        End If
                        'für den Nächsten wieder zu null setzen
                        neighbours = 0
                    End With
                    If bDraw Then
                        bDraw = False
                        'jetzt zeichnen
                        aPB.FillColor = c
                        'Trick 2:
                        '  nur die Änderung und direkt hier in der Prozedur zeichnen
                        'Trick 3:
                        '  die Koordinaten sind bereits berechnet
                        With .Field.Arr(i, j)
                            If bRect Then
                                Call DrawRect(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                            Else
                                Call DrawCircle(hhdc, .P1.X, .P1.Y, .P2.X, .P2.Y)
                            End If
                        End With
                    End If
                Next
            Next
            'die Generationen hochzählen
            .Counter = .Counter + 1
            'Trick 4:
            '  hier ein Trick damit man endlich weiß wofür die Zeigergeschichte
            '  überhaupt gut ist:
            '  einfach die Zeiger auf die Arraydaten vertauschen.
            '  Das hat den Vorteil dass nichts umkopiert werden muß.
            '  Man spart sich dadurch die Copy-Arbeit
            pTemp = .pNext.pUDT.pvData
            .pNext.pUDT.pvData = .pThis.pUDT.pvData
            .pThis.pUDT.pvData = pTemp
            'Trick 5:
            ' kein VB-Timer sondern nur ein Sleep und in bestimmten Intervallen
            ' ein Doevents
            If (.Counter Mod .GenTilDoEv) = 0 Then
                'damit man im Programm noch irgendwas machen kann
                'gibt es alle x Generationen ein Doevents x = GenTilDoEv ist einstellbar
                Call FrmGameOfLife.SetCounter(.Counter)
                DoEvents
            End If
            If .LifeTime > 0 Then Call Sleep(.LifeTime)
        Loop 'Do While .bIsRunning
    End With 'this
    Call MProcess.SetProcessPriority(THREAD_PRIORITY_NORMAL, NORMAL_PRIORITY_CLASS)
End Sub
Public Sub SaveFile(this As TGameOfLife, aPFN As String)
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Output As FNr
    With this
        '.ThisG.Arr(i, j).Life
        Dim i As Long, m As Long: m = .SizeX
        Dim j As Long, n As Long: n = .SizeY
        Print #FNr, "0" '; vbCrLf
        Print #FNr, "0" '; vbCrLf
        Print #FNr, "0" '; vbCrLf
        Print #FNr, "0" '; vbCrLf
        'Print FNr, "0"
        
        Print #FNr, True '; vbCrLf
        Print #FNr, False '; vbCrLf
        Print #FNr, "1" '; vbCrLf
        Print #FNr, m '; vbCrLf
        Print #FNr, n '; vbCrLf
        With .ThisG '.Arr(i, j).Life
            For j = 1 To n
                For i = 1 To m
                    Print #FNr, CStr(Abs(.Arr(i, j).Life));
                Next
                Print #FNr, "" 'vbCrLf
                'sline = ""
            Next
        End With
    End With
Finally:
    Close FNr
End Sub

