VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function Rectangle Lib "gdi32.dll" ( _
                                                ByVal hdc As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long) As Long
                                                
Private Declare Function RoundRect Lib "gdi32.dll" ( _
                                                ByVal hdc As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long, _
                                                ByVal X3 As Long, _
                                                ByVal Y3 As Long) As Long
                                                

Private Declare Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" _
                    (pArray() As Any, lngPtr As Long)
                    


Public Event Progress(ByVal Generation As Long, ByVal FPS As Single)

Public Event Started(ByVal AutoStop As Boolean)
Public Event Stopped(ByVal AutoStop As Boolean)

Private m_byte_World() As Byte              'zweidimensionales Array mit Abbild der Aktuellen Population
Private m_byte_World2() As Byte             'zweidimensionales Array zur temporären Ablage der Population
                           
Private m_byte_BorderType As Byte           'Verhalten an den Life Welt -Grenzen 0 = begrenzte Welt / 1 = unbegrenzte Welt (Kugelähnliche Oberfläche)
Private m_lng_Generation As Long
Private m_bln_DrawGrid As Boolean
Private m_ctlDestPic As PictureBox
Private m_byteCellSize As Byte              'X/Y Ausdehnung einer Zelle in Pixel
Private m_intWorldSize As Integer           'Weltgröße (X/Y Ausdehnung der Life-Welt in Pixel)
Private m_bln_Running As Boolean
Private m_int_Interval As Integer           'Generationsinterwall der Lifewelt in Millisekunden
Private m_int_StepsPassed As Integer
Private m_int_Steps2Go As Integer
Private m_bln_AutoStop As Boolean
Private m_bln_CellsRounded As Boolean

Private m_BirthControlRule As String        'Geburts-Regel
Private m_RuleOfLife As String              'Überlebens-Regel


Public Property Get RulesDefinition() As String
    RulesDefinition = Trim(m_RuleOfLife) & "/" & Trim(m_BirthControlRule)
End Property
Public Property Let RulesDefinition(ByVal New_Value As String)
    If (InStr(1, New_Value, "/") > 0) Then
        m_RuleOfLife = Trim(Split(New_Value, "/")(0))
        m_BirthControlRule = Trim(Split(New_Value, "/")(1))
    Else
        MsgBox "Die Regeldefinition enthält einen Fehler!" & vbCrLf & _
                "Bitte überprüfen Sie ob die Regeldefinition im Format:" & vbCrLf & _
                "Anzahlen lebender Nachbarzellen für Überleben / Anzahlen lebender Nachbarzellen für Geburt" & vbCrLf & _
                "Beispiel: ""23/3"" oder ""2,3/3"" vorliegt.", vbCritical
    End If
End Property


Public Property Get DestPic() As PictureBox
    Set DestPic = m_ctlDestPic
End Property

Public Property Get World() As Byte()
    World = m_byte_World()
End Property
Public Property Let World(New_World() As Byte)
    m_byte_World = New_World
    m_intWorldSize = UBound(m_byte_World, 1) + 1
End Property

Public Property Set DestPic(ctlDestPic As PictureBox)
    Set m_ctlDestPic = ctlDestPic
End Property

Public Property Get DrawGrid() As Boolean
    DrawGrid = m_bln_DrawGrid
End Property
Public Property Let DrawGrid(ByVal New_DrawGrid As Boolean)
    m_bln_DrawGrid = New_DrawGrid
End Property

Public Property Get CellSize() As Byte
    'X/Y Ausdehnung einer Zelle in Pixel
    CellSize = m_byteCellSize
End Property
Public Property Let CellSize(ByVal NewCellSize As Byte)
    'X/Y Ausdehnung einer Zelle in Pixel
    m_byteCellSize = NewCellSize
End Property

Public Property Get WorldXSize() As Integer
    WorldXSize = UBound(m_byte_World(), 1)
End Property

Public Property Get WorldSize() As Integer
    'Weltgröße (X/Y Ausdehnung der Life-Welt in Pixel)
    WorldSize = m_intWorldSize
End Property
Public Property Let WorldSize(ByVal New_WorldSize As Integer)
    'Weltgröße (X/Y Ausdehnung der Life-Welt in Pixel)
    m_intWorldSize = New_WorldSize
End Property

Public Property Get isRunning() As Boolean
    isRunning = m_bln_Running
End Property

Public Property Get Generation() As Long
    Generation = m_lng_Generation
End Property
Public Property Let Generation(ByVal New_Generation As Long)
    m_lng_Generation = New_Generation
End Property

Public Property Get CellsRounded() As Boolean
    CellsRounded = m_bln_CellsRounded
End Property
Public Property Let CellsRounded(ByVal New_CellsRounded As Boolean)
    m_bln_CellsRounded = New_CellsRounded
    DrawAll
End Property

Public Property Get BorderType() As Byte
    BorderType = m_byte_BorderType
End Property
Public Property Let BorderType(ByVal New_BorderType As Byte)
    m_byte_BorderType = New_BorderType
End Property

Public Property Get Interval() As Integer
    Interval = m_int_Interval
End Property
Public Property Let Interval(ByVal New_Interval As Integer)
    m_int_Interval = New_Interval
End Property

Public Property Get Steps2Go() As Integer
    Steps2Go = m_int_Steps2Go
End Property
Public Property Let Steps2Go(ByVal New_Steps2Go As Integer)
    m_int_Steps2Go = New_Steps2Go
End Property

Public Property Let SetCell(ByVal X As Long, ByVal Y As Long, ByVal New_Value As Byte)
    On Error GoTo ErrHandler
    m_byte_World(X, Y) = New_Value
    DrawBox X, Y
    Exit Property
ErrHandler:
    Err.Clear
End Property
Public Property Get GetCell(ByVal X As Long, ByVal Y As Long) As Byte
    On Error GoTo ErrHandler
    GetCell = m_byte_World(X, Y)
    Exit Property
ErrHandler:
    Err.Clear
End Property

Public Sub GoSteps(Optional ByVal Steps As Integer)
    If Steps > 0 Then
        m_int_Steps2Go = Steps
    End If
    m_bln_AutoStop = True
    m_int_StepsPassed = 0
    m_bln_Running = True
    Call GoL_Loop
End Sub

Public Sub WorldStart()
    m_bln_Running = True
    Call GoL_Loop
End Sub

Public Sub WorldStop()
    m_bln_Running = False
End Sub

Public Sub Clear_World()
    Dim X As Long
    Dim Y As Long
    For X = 0 To m_intWorldSize
        For Y = 0 To m_intWorldSize
            m_byte_World(X, Y) = 0
            m_byte_World2(X, Y) = 0
        Next
    Next
    m_lng_Generation = 0
    DrawAll
End Sub

Public Sub Randomize_World()
    'Zufällige Start-Populationen erstellen
    Dim X As Long
    Dim Y As Long
    Randomize
    For X = 0 To m_intWorldSize
        For Y = 0 To m_intWorldSize
            m_byte_World(X, Y) = IIf(Int(Rnd * 2) = 1, 1, 0)
        Next Y
    Next X
    m_lng_Generation = 0
    DrawAll
End Sub

Public Sub Resize_World(ByVal New_ZellSize As Integer, ByVal New_WorldSize As Integer)
    Dim X As Long
    Dim Y As Long
    Dim lngPtr As Long
    
    m_byteCellSize = New_ZellSize    '5       'Größe der einzelnen Felder
    m_intWorldSize = New_WorldSize   '200     'Feldgröße (also g_intNew_WorldSize * g_intNew_WorldSize groß)
   
    m_ctlDestPic.Height = m_byteCellSize * m_intWorldSize
    m_ctlDestPic.Width = m_byteCellSize * m_intWorldSize
    m_byte_World2 = m_byte_World
    ReDim m_byte_World(m_intWorldSize, m_intWorldSize)
    GetSafeArrayPointer m_byte_World2(), lngPtr
    If lngPtr > 0 Then
        For X = 0 To UBound(m_byte_World2, 1)
            If X < UBound(m_byte_World, 1) Then
                For Y = 0 To UBound(m_byte_World2, 1)
                    If Y < UBound(m_byte_World, 1) Then
                        m_byte_World(X, Y) = m_byte_World2(X, Y)
                    End If
                Next Y
            End If
        Next X
    End If
    ReDim m_byte_World2(m_intWorldSize, m_intWorldSize)
    Call DrawAll
End Sub

Friend Function Copy_Rect(ByRef rSrc As RECT) As Byte()
On Error GoTo ErrOut
    Dim rFigur As RECT
    Dim X As Long
    Dim Y As Long
    Dim tmpFigur() As Byte
    With rSrc
        ReDim tmpFigur(0 To .Right - .Left, 0 To .Bottom - .Top)
        For X = 0 To UBound(tmpFigur, 1)
            For Y = 0 To UBound(tmpFigur, 2)
                tmpFigur(X, Y) = m_byte_World(X + .Left, Y + .Top)
            Next Y
        Next X
    End With
    Copy_Rect = tmpFigur
    Exit Function
ErrOut:
    Err.Clear
End Function

Friend Sub Clear_Rect(ByRef DestRect As RECT)
    Dim X As Long
    Dim Y As Long
    With DestRect
        For X = .Left To .Right - 1
            For Y = .Top To .Bottom - 1
                m_byte_World(X, Y) = 0
                m_byte_World2(X, Y) = 0
                DrawBox X, Y
            Next
        Next
    End With
    m_lng_Generation = 0
    DrawAll
End Sub

Friend Sub Fill_Rect(ByRef DestRect As RECT)
    Dim X As Long
    Dim Y As Long
    With DestRect
        For X = .Left To .Right - 1
            For Y = .Top To .Bottom - 1
                m_byte_World(X, Y) = 1
                m_byte_World2(X, Y) = 0
                DrawBox X, Y
            Next
        Next
    End With
    m_lng_Generation = 0
    DrawAll
End Sub

Friend Function Rotate_Figur(ByRef Figur() As Byte) As RECT
On Error GoTo ErrOut
    Dim rFigur As RECT
    Dim X As Long
    Dim Y As Long
    Dim tmpFigur() As Byte
    With rFigur
        .Left = 0
        .Top = 0
        .Right = UBound(Figur, 2) 'Dimension 2 war zuvor X
        .Bottom = UBound(Figur, 1) 'Dimension 1 war zuvor X

        tmpFigur = Figur
        ReDim Figur(.Left To .Right, .Top To .Bottom)
        For Y = 0 To .Bottom
            For X = 0 To .Right
                Figur(X, Y) = tmpFigur(Y, Abs(X - .Right))
            Next X
        Next Y
    End With
    Rotate_Figur = rFigur
    Exit Function
ErrOut:
    Err.Clear
End Function

Friend Function Rotate_Rect(ByRef rSrc As RECT) As RECT
On Error GoTo ErrOut
Dim rDst As RECT
Dim X As Long
Dim Y As Long
Dim lXLen As Long
Dim lYLen As Long
Dim tmpArea() As Byte
    rDst = rSrc

    With rDst
        .Right = .Left + rSrc.Bottom - rSrc.Top
        .Bottom = .Top + rSrc.Right - rSrc.Left
    End With
    tmpArea = Copy_Rect(rSrc)
    Clear_Rect rSrc
    Rotate_Figur tmpArea
    InsertFigur2Rect tmpArea, rDst

    Rotate_Rect = rDst
    Exit Function
ErrOut:
    Err.Clear
End Function

Public Sub Insert_Figur(ByRef Figur() As Byte, ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    Dim xx As Long
    Dim yy As Long
    Dim Xw As Long 'X Koordinate Welt
    Dim Yw As Long 'Y Koordinate Welt
    
    X = X - ((UBound(Figur, 1) + 1) \ 2)
    Y = Y - ((UBound(Figur, 2) + 1) \ 2)
    For xx = 0 To UBound(Figur, 1)
        For yy = 0 To UBound(Figur, 2)
            Xw = xx + 1 + X
            Yw = yy + 1 + Y
            If (Xw + 1) < 1 Then Xw = m_intWorldSize + Xw
            If (Yw + 1) < 1 Then Yw = m_intWorldSize + Yw
            If (Xw + 1) > m_intWorldSize Then Xw = Abs((m_intWorldSize) - Xw)
            If (Yw + 1) > m_intWorldSize Then Yw = Abs((m_intWorldSize) - Yw)
            m_byte_World(Xw, Yw) = Figur(xx, yy)
            DrawBox Xw, Yw
        Next yy
    Next xx

    m_ctlDestPic.Refresh
    
End Sub

Friend Sub InsertFigur2Rect(ByRef Figur() As Byte, ByRef DestRect As RECT, Optional ByVal Center As Boolean = False)
    On Error Resume Next
    Dim src As RECT
    Dim dst As RECT
    Dim X As Long
    Dim Y As Long
    Dim Xd As Long
    Dim Yd As Long
    Dim Xs As Long
    Dim Ys As Long
    
    dst = DestRect

    With src
        .Left = 0
        .Top = 0
        .Right = UBound(Figur, 1) '+ 1
        .Bottom = UBound(Figur, 2) '+ 1
    End With

    Xd = (src.Right - src.Left) - (dst.Right - dst.Left)
    Yd = (src.Bottom - src.Top) - (dst.Bottom - dst.Top)

    Select Case Xd
        Case Is < 0     ' Zielbreite ist größer als Quellbreite
            Xd = Abs(Xd)
            If Center Then
                dst.Left = dst.Left + Xd \ 2
                dst.Right = dst.Right - (Xd - Xd \ 2)
            Else
                dst.Right = dst.Right + Xd
            End If
            Xs = (src.Right - src.Left)

        Case Is > 0     ' Quellbreite ist größer als Zielbreite
            If Center Then
                src.Left = src.Left + Xd \ 2
                src.Right = src.Right - (Xd - Xd \ 2)
            Else
                src.Right = src.Right + Xd
            End If
            Xs = (dst.Right - dst.Left)
        Case Else       ' Quellbreite ist gleich Zielbreite
            Xs = (src.Right - src.Left)
    End Select

    Select Case Yd
        Case Is < 0     ' Zielhöhe ist größer als Quellhöhe
            Yd = Abs(Yd)
            If Center Then
                dst.Top = dst.Top + Yd \ 2
                dst.Bottom = dst.Bottom - (Yd - Yd \ 2)
            Else
                dst.Bottom = dst.Bottom + Yd
            End If
            Ys = (src.Bottom - src.Top)
        Case Is > 0     ' Quellhöhe ist größer als Zielhöhe
            If Center Then
                src.Top = src.Top + Yd \ 2
                src.Bottom = src.Bottom - (Yd - Yd \ 2)
            Else
                src.Bottom = src.Bottom + Yd
            End If
            Ys = (dst.Bottom - dst.Top)
            
        Case Else       ' Quellhöhe ist gleich Zielhöhe
            Ys = (src.Bottom - src.Top)
    End Select

    With DestRect
        For X = 0 To Xs '- 1
            For Y = 0 To Ys '- 1
                m_byte_World(dst.Left + X, dst.Top + Y) = Figur(src.Left + X, src.Top + Y)
                DrawBox dst.Left + X, dst.Top + Y
            Next Y
        Next X
    End With
    m_ctlDestPic.Refresh
    
End Sub

Friend Function GetFigurRect(SrcArea As RECT) As RECT
On Error GoTo ErrOut
Dim X As Integer
Dim Y As Integer
Dim rDest As RECT
    With SrcArea
        rDest.Top = .Bottom
        rDest.Left = .Right
        rDest.Bottom = .Top
        rDest.Right = .Left
        For Y = .Top To .Bottom
            For X = .Left To .Right
                If m_byte_World(X, Y) = 1 Then
                    If X < rDest.Left Then rDest.Left = X
                    Exit For
                End If
            Next X
            For X = .Right To .Left Step -1
                If m_byte_World(X, Y) = 1 Then
                    If X > rDest.Right Then rDest.Right = X
                    Exit For
                End If
            Next X
        Next Y
        For X = .Left To .Right
            For Y = .Top To .Bottom
                If m_byte_World(X, Y) = 1 Then
                    If Y < rDest.Top Then rDest.Top = Y
                    Exit For
                End If
            Next Y
            For Y = .Bottom To .Top Step -1
                If m_byte_World(X, Y) = 1 Then
                    If Y > rDest.Bottom Then rDest.Bottom = Y
                    Exit For
                End If
            Next Y
        Next X
    End With 'SrcArea
    With rDest
        If .Left > .Right Then .Left = 0
        If .Top > .Bottom Then .Top = 0
    End With
    GetFigurRect = rDest
    Exit Function
ErrOut:
    GetFigurRect = SrcArea
    Err.Clear
End Function

Private Sub GoL_Loop()  'Loop: Game of Life
'On Error Resume Next
    If m_bln_Running = False Then Exit Sub
    Dim iNB As Long
    Dim X As Long, Y As Long
    Dim sng_FPS As Single
    Dim lng_FramesCount  As Long
    Dim lastTime As Long
    Dim GoTime As Long
    Dim DeathCandidate As Boolean
    
    sng_FPS = 0
    lng_FramesCount = 0
    lastTime = 0
    GoTime = 0
    
    RaiseEvent Started(m_bln_AutoStop)
    Do
        For X = 0 To m_intWorldSize
            For Y = 0 To m_intWorldSize
                Select Case m_byte_BorderType
                    Case 0
                        iNB = CheckNB_A(m_byte_World, X, Y)
                    Case 1
                        iNB = CheckNB_B(m_byte_World, X, Y)
                End Select
                
'                Select Case iNB
'                    Case 3                                      'Geburts-Regel:
'                        'Eine tote Zelle mit genau drei lebenden Nachbarn
'                        'wird in der Folgegeneration neu geboren.
'                        'Eine lebende Zelle mit genau drei lebenden Nachbarn
'                        'bleibt in der Folgegeneration lebend.
'                        m_byte_World2(X, Y) = 1
'                    Case 0, 1, 4 To 8                           'Todes-Regel
'                        'Lebende Zellen mit weniger als zwei lebenden Nachbarn
'                        'sterben in der Folgegeneration an Einsamkeit.
'                        'Lebende Zellen mit mehr als drei lebenden Nachbarn
'                        'sterben in der Folgegeneration an Überbevölkerung.
'                        m_byte_World2(X, Y) = 0
'                    Case Else                   'Case 2
'                        'Bei zwei lebenden Nachbarzellen behält
'                        'die Zelle in der Folgegeneration ihren Status.
'                        m_byte_World2(X, Y) = m_byte_World(X, Y)
'                End Select
                
                    DeathCandidate = True
                    
                    Select Case True
                        Case (InStr(1, m_BirthControlRule, CStr(iNB)) > 0) And (m_byte_World(X, Y) = 0)         'Geburts-Regel:
                            'Eine tote Zelle mit m_BirthControlRule lebenden Nachbarn
                            'wird in der Folgegeneration neu geboren.
                            m_byte_World2(X, Y) = 1
                            DeathCandidate = False
                        Case (InStr(1, m_RuleOfLife, CStr(iNB)) > 0)                                            'Überlebens-Regel:
                            'Bei m_RuleOfLife lebenden Nachbarzellen behält
                            'die Zelle in der Folgegeneration ihren Status.
                            m_byte_World2(X, Y) = m_byte_World(X, Y)
                            DeathCandidate = False
                    End Select
                    Select Case DeathCandidate
                        Case True  '0, 1, 4 To 8                                                                'Todes-Regel
                            'Alle oben nicht berücksichtigten Zellen müssen sterben
                            m_byte_World2(X, Y) = 0
                    End Select
                Select Case DeathCandidate

                End Select
            Next Y
        Next X
        
        For X = 0 To m_intWorldSize
            For Y = 0 To m_intWorldSize
                ' nur die geänderten Felder neu zeichen
                Select Case True
                    Case m_byte_World(X, Y) <> m_byte_World2(X, Y)
                    m_byte_World(X, Y) = m_byte_World2(X, Y)
                    DrawBox X, Y
                End Select
            Next Y
        Next X
        
        If m_lng_Generation >= &H7FFFFFFF Then
            m_bln_Running = False
            'mnuStartStop.Caption = "&Start"
            Stop
        Else
            m_lng_Generation = m_lng_Generation + 1
        End If
        m_ctlDestPic.Refresh

        Do
            DoEvents: Sleep 0
        Loop While (Abs(timeGetTime - GoTime) < m_int_Interval)
        GoTime = timeGetTime
        
        'FPS Berechnen
        lng_FramesCount = lng_FramesCount + 1
        Select Case True
            Case lng_FramesCount Mod 10 = 0
                lng_FramesCount = 0
                If lastTime <> 0 Then sng_FPS = (sng_FPS + (1000 * 10) / Abs(timeGetTime - lastTime)) / 2
                lastTime = timeGetTime
        End Select
        
        RaiseEvent Progress(m_lng_Generation, sng_FPS)
        
        If m_bln_AutoStop Then
            m_int_StepsPassed = m_int_StepsPassed + 1
            If m_int_StepsPassed >= m_int_Steps2Go Then
                m_bln_AutoStop = False
                m_bln_Running = False
                RaiseEvent Stopped(True)
                m_int_StepsPassed = 0
                Exit Sub
            End If
        End If
    Loop While m_bln_Running
    RaiseEvent Stopped(False)
End Sub

Public Sub DrawAll()
    ' zeichnet das ganze Spielfeld neu
    Dim X As Long
    Dim X1 As Long
    Dim Y As Long
    Dim Y1 As Long
    Dim ext  As Long
    Dim lngPtr As Long
    On Error GoTo ErrOut
    ext = m_byteCellSize
    m_ctlDestPic.Cls
    GetSafeArrayPointer m_byte_World(), lngPtr
    If lngPtr <= 0 Then Exit Sub
    
    For X = 0 To m_intWorldSize
        For Y = 0 To m_intWorldSize
            X1 = X * m_byteCellSize
            Y1 = Y * m_byteCellSize
            m_ctlDestPic.FillColor = GetColor(m_byte_World(X, Y))
            If m_bln_DrawGrid Then
                If ext <= 2 Then
                    m_ctlDestPic.ForeColor = m_ctlDestPic.FillColor
                    Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + ext + 1, Y1 + ext + 1
                Else
                    m_ctlDestPic.ForeColor = RGB(200, 200, 250)
                    If m_bln_CellsRounded Then
                        RoundRect m_ctlDestPic.hdc, X1, Y1, X1 + ext + 1, Y1 + ext + 1, 25, 25
                    Else
                        Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + ext + 1, Y1 + ext + 1
                    End If
                End If
            Else
                m_ctlDestPic.ForeColor = m_ctlDestPic.FillColor
                If ext <= 2 Then
                    Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + 2, Y1 + 2
                Else
                    If m_bln_CellsRounded Then
                        RoundRect m_ctlDestPic.hdc, X1, Y1, X1 + ext, Y1 + ext, 25, 25
                    Else
                        Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + ext, Y1 + ext
                    End If
                End If
            End If
        Next
    Next
    m_ctlDestPic.Refresh
ErrOut:
    Err.Clear
End Sub

Public Sub DrawBox(ByVal X As Long, ByVal Y As Long)
    ' neu Funktion, um nur ein Feld zu zeichnen
    Dim ext&
    Dim X1 As Long
    Dim Y1 As Long
    X1 = X * m_byteCellSize
    Y1 = Y * m_byteCellSize
    ext = m_byteCellSize
    
    m_ctlDestPic.FillColor = GetColor(m_byte_World(X, Y))
    If m_bln_DrawGrid Then
        If ext <= 2 Then
            m_ctlDestPic.ForeColor = m_ctlDestPic.FillColor
            Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + 2, Y1 + 2
        Else
            m_ctlDestPic.ForeColor = RGB(200, 200, 250)
            If m_bln_CellsRounded Then
                RoundRect m_ctlDestPic.hdc, X1, Y1, X1 + ext + 1, Y1 + ext + 1, 25, 25
            Else
                Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + ext + 1, Y1 + ext + 1
            End If
        End If
    Else
        m_ctlDestPic.ForeColor = m_ctlDestPic.FillColor
        If ext <= 2 Then
            Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + 2, Y1 + 2
        Else
            If m_bln_CellsRounded Then
                RoundRect m_ctlDestPic.hdc, X1, Y1, X1 + ext, Y1 + ext, 25, 25
            Else
                Rectangle m_ctlDestPic.hdc, X1, Y1, X1 + ext, Y1 + ext
            End If
        End If
    End If
End Sub

Public Sub DrawIt()
    Dim X As Long
    Dim Y As Long
    For X = 0 To m_intWorldSize
        For Y = 0 To m_intWorldSize
            m_ctlDestPic.Line (X * m_byteCellSize, Y * m_byteCellSize)- _
                    (X * m_byteCellSize + m_byteCellSize - IIf(m_bln_DrawGrid, 30, 0), Y * m_byteCellSize + m_byteCellSize - IIf(m_bln_DrawGrid, 30, 0)), _
                    GetColor(m_byte_World(X, Y)), BF
        Next Y
    Next X
End Sub

Public Function GetColor(ByVal ii As Integer) As Long
    If ii = 0 Then GetColor = &HFFFFFF Else GetColor = &HFF0000
End Function

Public Function CheckNB_B(ByRef World() As Byte, ByVal xx As Long, ByVal yy As Long) As Long   'Check NeighBours (= Anzahl der besetzten Felder um ein bestimmtes Feld herum)
    Dim iNB As Integer
    Dim X As Long
    Dim Y As Long
    Dim rDest As RECT
    Dim Xr As Long
    Dim Yr As Long
    iNB = 0
    
    With rDest
        .Left = xx - 1
        .Right = xx + 1
        .Top = yy - 1
        .Bottom = yy + 1
        
        For X = .Left To .Right
            For Y = .Top To .Bottom
                Xr = X
                Yr = Y
                If Xr < 0 Then Xr = X + m_intWorldSize
                If Yr < 0 Then Yr = Y + m_intWorldSize
                If Xr > m_intWorldSize - 1 Then Xr = X - m_intWorldSize
                If Yr > m_intWorldSize - 1 Then Yr = Y - m_intWorldSize
                If World(Xr, Yr) >= 1 Then iNB = iNB + 1
            Next Y
        Next X
    End With 'rDest
    If World(xx, yy) >= 1 Then iNB = iNB - 1 ' Korrektur, da zu prüfendes Feld als Nachbar mitgezählt wurde
    CheckNB_B = iNB
      
End Function

Public Function CheckNB_A(ByRef World() As Byte, ByVal xx As Long, ByVal yy As Long) As Long  'Check NeighBours (= Anzahl der besetzten Felder um ein bestimmtes Feld herum)
    Dim iNB As Integer
    Dim X As Long
    Dim Y As Long
    Dim rDest As RECT
    iNB = 0
    With rDest
        .Left = xx - 1
        If .Left < 0 Then .Left = 0
        .Right = xx + 1
        If .Right > m_intWorldSize Then .Right = m_intWorldSize
        
        .Top = yy - 1
        If .Top < 0 Then .Top = 0
        .Bottom = yy + 1
        If .Bottom > m_intWorldSize Then .Bottom = m_intWorldSize
        
        For X = .Left To .Right
            For Y = .Top To .Bottom
                If X < 0 Or X > m_intWorldSize Or Y < 0 Or Y > m_intWorldSize Then
                    iNB = iNB + 2
                Else
                    If World(X, Y) >= 1 Then iNB = iNB + 1
                End If
            Next Y
        Next X
    End With 'rDest
    If World(xx, yy) >= 1 Then iNB = iNB - 1 ' Korrektur, da zu prüfendes Feld als Nachbar mitgezählt wurde
    CheckNB_A = iNB
  
End Function

Private Sub Class_Initialize()
    m_byte_BorderType = 0               'Verhalten an den Life Welt -Grenzen 0 = begrenzte Welt / 1 = unbegrenzte Welt (Kugelähnliche Oberfläche)
    m_bln_DrawGrid = False
    m_byteCellSize = 5                  'X/Y Ausdehnung einer Zelle in Pixel
    m_intWorldSize = 100                'Weltgröße (X/Y Ausdehnung der Life-Welt in Pixel)
    m_int_Interval = 50                 'Generationsinterwall der Lifewelt in Millisekunden
    m_int_Steps2Go = 1
    m_bln_AutoStop = False
    m_bln_CellsRounded = True
    m_BirthControlRule = 3          'Geburts-Regel
    m_RuleOfLife = 23               'Überlebens-Regel
End Sub

Private Sub Class_Terminate()
    Set m_ctlDestPic = Nothing
    Erase m_byte_World
    Erase m_byte_World2
End Sub
