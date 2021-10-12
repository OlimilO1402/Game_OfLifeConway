VERSION 5.00
Begin VB.Form FrmGameOfLife 
   Caption         =   "Conway's Game Of Life "
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton BtnDelRandom 
      Caption         =   "Clear Random Density"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton BtnSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   8160
      Width           =   1815
   End
   Begin VB.CommandButton BtnCreateLabyrinth 
      Caption         =   "Create Labyrinth"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton BtnRandom 
      Caption         =   "Fill Random w. Density"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton BtnCreateNewField 
      Caption         =   "Create New Field"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   1815
      Begin VB.CommandButton BtnColor2 
         Caption         =   "2"
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton BtnColor1 
         Caption         =   "1"
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox TxtLColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtFColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtGColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Life:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Field:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Frame:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.CommandButton BtnSetLifeRule 
      Caption         =   "SetLifeRule"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   1815
   End
   Begin VB.OptionButton OptPtFormCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.OptionButton OptPtFormRect 
      Caption         =   "Rect"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox ChkFixPointSize 
      Caption         =   "FixPointSize"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox TxtLifetime 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox TxtDoevents 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox PBGOL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   6  'Innen ausgef�llt
      FillColor       =   &H000000FF&
      ForeColor       =   &H00004000&
      Height          =   8535
      Left            =   2040
      ScaleHeight     =   8475
      ScaleWidth      =   8475
      TabIndex        =   16
      Top             =   0
      Width           =   8535
   End
   Begin VB.CommandButton BtnStartStop 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton BtnDrawNew 
      Caption         =   "Draw"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox TxtX 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox TxtY 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label LblLifeRule 
      Alignment       =   2  'Zentriert
      Caption         =   "/"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "lifetime [ms]:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Gener. until events:"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "SizeX:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "SizeY:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   615
   End
End
Attribute VB_Name = "FrmGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '2008_08_06 Zeilen:  124
Private m_GOL   As TGameOfLife
Private m_Timer As Double
Private m_r     As Long 'Prozent Populationsdichte

Private Sub BtnColor1_Click()
    Me.TxtFColr.Text = "&H" & Hex$(&H0) 'Hex(&H4000&) '& "&"
    Me.TxtLColr.Text = "&H" & Hex$(&HFF007F)   '& "&"
    Me.TxtGColr.Text = "&H" & Hex$(&H0)   '& "&"
End Sub
Private Sub BtnColor2_Click()
    Me.TxtFColr.Text = "&H" & Hex$(&HFFFFFF) 'Hex(&H4000&) '& "&"
    Me.TxtLColr.Text = "&H" & Hex$(&H0)   '& "&"
    Me.TxtGColr.Text = "&H" & Hex$(&HCCCCCC)   '& "&"
End Sub

Private Sub BtnCreateLabyrinth_Click()
    Call BtnCreateNewField_Click
'Fill Random
    m_r = 6
    Call BtnRandom_Click
    'setze Regeln 234/3 oder 2346/3
    With m_GOL.LifeRule
        .RuleSurvive = RNB2 Or RNB3 Or RNB4 Or RNB6
        .RuleNewBorn = RNB3
    End With
    LblLifeRule.Caption = MLifeRule.LifeRuleToString(m_GOL.LifeRule)
    BtnStartStop_Click
'delete some lifes
    Dim sr As String
    m_r = 20 '25
    sr = InputBox("Please give the delete density in %", "Delete Density", CStr(m_r))
    If StrPtr(sr) = 0 Then Exit Sub
    If Len(sr) = 0 Then Exit Sub
    If Not IsNumeric(sr) Then
        MsgBox "Bitte Zahl eingeben!"
        Exit Sub
    End If
    m_r = CLng(sr)
    Call MGameOfLife.DeleteRandom(m_GOL, m_r)
    Call BtnDrawNew_Click

End Sub

Private Sub BtnSave_Click()
    Dim aPFN As String
    aPFN = App.Path & "\Lab.pff"
    'aPFN = ""
    aPFN = InputBox("Pfad", "Pfad?", aPFN)
    Call MGameOfLife.SaveFile(m_GOL, aPFN)
End Sub

Private Sub Form_Load()
    'einige Voreinstellungen vornehmen
    Me.PBGOL.AutoRedraw = True
    Me.PBGOL.ScaleMode = vbPixels  '3 ' Pixel
    Me.TxtDoevents.Text = CStr(20) 'alle x Generationen ein Doevents
    Me.TxtLifetime.Text = CStr(0)  'in Millisekunden
    Me.TxtX.Text = CStr(180)
    Me.TxtY.Text = CStr(180)
    m_r = 5 '%
    Me.TxtFColr.Text = "&H" & Hex$(&H0) 'Hex(&H4000&) '& "&"
    Me.TxtLColr.Text = "&H" & Hex$(&HFF007F)   '& "&"
    Me.TxtGColr.Text = "&H" & Hex$(&H0)   '& "&"
    Me.Label6.Caption = "Click into the field " & vbCrLf & _
                        "to switch on/off the " & vbCrLf & _
                        "individual life " & vbCrLf
    Me.OptPtFormCircle.Value = True
    Me.ChkFixPointSize.Value = vbChecked
    Me.LblLifeRule.Caption = "236/3"
    Call BtnCreateNewField_Click
    Call MGameOfLife.InitRandom(m_GOL, m_r)
    Call BtnDrawNew_Click
End Sub
Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    With Me.PBGOL
        L = .Left:             T = .Top
        W = Me.ScaleWidth - L: H = Me.ScaleHeight - T
        If W > 0 And H > 0 Then Call .Move(L, T, W, H)
    End With
    Call BtnDrawNew_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    m_GOL.bIsRunning = False
    Call MGameOfLife.Delete(m_GOL)
End Sub

Public Sub SetCounter(ByVal CountGenerations As Long)
On Error GoTo catch
    Dim d As Double: d = Timer - m_Timer
    m_Timer = Timer
    Me.Caption = "Generations: " & CStr(CountGenerations) & "   " & CStr(Int(m_GOL.GenTilDoEv / d)) & " pro sec"
    Exit Sub
catch:
    'MsgBox Err.Description
End Sub
Private Sub BtnStartStop_Click()
    m_GOL.bIsRunning = Not m_GOL.bIsRunning
    If IsNumeric(TxtDoevents.Text) Then
        m_GOL.GenTilDoEv = Abs(CLng(TxtDoevents.Text))
    End If
    If IsNumeric(TxtLifetime) Then
        m_GOL.LifeTime = Abs(CLng(TxtLifetime.Text))
    End If
    If Not m_GOL.bIsRunning Then
        BtnStartStop.Caption = "Start"
        Me.Caption = "Conway's Game Of Life " & Me.Caption
        Call BtnDrawNew_Click
    Else
        BtnStartStop.Caption = "Pause"
        Me.PBGOL.AutoRedraw = False
        m_Timer = Timer
        Call MGameOfLife.Run(m_GOL, Me.PBGOL)
        Me.PBGOL.AutoRedraw = True 'False
        BtnDrawNew_Click
    End If
End Sub

Private Sub ChkFixPointSize_Click()
    Call BtnDrawNew_Click
End Sub
Private Sub OptPtFormCircle_Click()
    Call BtnDrawNew_Click
End Sub
Private Sub OptPtFormRect_Click()
    Call BtnDrawNew_Click
End Sub
Private Sub BtnDrawNew_Click()
    On Local Error Resume Next
    With m_GOL
        With .Field
            .FieldColor = CLng(TxtFColr.Text)
            .LifeColor = CLng(TxtLColr.Text)
            .GittColor = CLng(TxtGColr.Text)
            .bFixPtSize = (ChkFixPointSize.Value = vbChecked)
            If OptPtFormCircle.Value Then
                .LifeForm = LifeFormCircle
            ElseIf OptPtFormRect.Value Then
                .LifeForm = LifeFormRectangle
            End If
            Me.PBGOL.BackColor = .FieldColor
            Me.PBGOL.ForeColor = .GittColor
            Me.PBGOL.FillColor = .LifeColor
        End With
        Call MField.CalcField(.Field, Me.PBGOL)
    End With
    Call MGameOfLife.DrawAll(m_GOL, Me.PBGOL)
End Sub
Private Sub BtnCreateNewField_Click()
    Dim X As Long
    Dim Y As Long
    If Len(TxtX.Text) = 0 Then Exit Sub
    If Len(TxtY.Text) = 0 Then Exit Sub
    If Not IsNumeric(TxtX.Text) Then
        MsgBox "Bitte Zahl eingeben!"
        Exit Sub
    End If
    If Not IsNumeric(TxtY.Text) Then
        MsgBox "Bitte Zahl eingeben!"
        Exit Sub
    End If
    X = CLng(TxtX.Text)
    Y = CLng(TxtY.Text)
    Call MGameOfLife.New_GameOfLife(m_GOL, X, Y, LblLifeRule.Caption)
    Call BtnDrawNew_Click
End Sub
Private Sub BtnRandom_Click()
    Dim sr As String
    sr = InputBox("Please give the population density in %", "Population Density", CStr(m_r))
    If StrPtr(sr) <> 0 Then
        If Len(sr) > 0 Then
            If IsNumeric(sr) Then
                m_r = CLng(sr)
                Call MGameOfLife.InitRandom(m_GOL, m_r)
                Call BtnDrawNew_Click
            Else
                MsgBox "Bitte Zahl eingeben!"
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub BtnDelRandom_Click()
    Dim sr As String
    sr = InputBox("Please give the delete density in %", "Delete Density", CStr(m_r))
    If StrPtr(sr) <> 0 Then
        If Len(sr) > 0 Then
            If IsNumeric(sr) Then
                m_r = CLng(sr)
                Call MGameOfLife.DeleteRandom(m_GOL, m_r)
                Call BtnDrawNew_Click
            Else
                MsgBox "Bitte Zahl eingeben!"
                Exit Sub
            End If
        End If
    End If
End Sub
Private Sub BtnClear_Click()
    Call MGameOfLife.Clear(m_GOL)
    Call BtnDrawNew_Click
End Sub
Private Sub BtnSetLifeRule_Click()
    If FrmLifeRule.ShowDialog(Me, m_GOL.LifeRule) = vbOK Then
        LblLifeRule.Caption = MLifeRule.LifeRuleToString(m_GOL.LifeRule)
    End If
End Sub

Private Sub PBGOL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As Point: pt = MField.GetIndexPoint(m_GOL.Field, X, Y)
    If Not MGameOfLife.IsPointInside(m_GOL, pt) Then Exit Sub
    Call MGameOfLife.SwitchLife(m_GOL, pt, (Button = vbLeftButton))
    With m_GOL
        Call MGeneration.DrawIndividual(.pThis.Generation, .Field, Me.PBGOL, pt)
    End With
    Me.PBGOL.Refresh
End Sub

Private Sub PBGOL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then Exit Sub
    Call PBGOL_MouseDown(Button, Shift, X, Y)
End Sub
