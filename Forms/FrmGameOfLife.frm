VERSION 5.00
Begin VB.Form FrmGameOfLife 
   Caption         =   "Conway's Game Of Life "
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "FrmGameOfLife.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton BtnRandom 
      Caption         =   "Fill Random %-Density"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton BtnLoadGeneration 
      Caption         =   "LoadGeneration"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton BtnSaveGeneration 
      Caption         =   "SaveGeneration"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton BtnCreateNewField 
      Caption         =   "Create New Field"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.TextBox TxtY 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   960
      TabIndex        =   11
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox TxtX 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton BtnSetLifeRule 
      Caption         =   "SetLifeRule"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton BtnDrawNew 
      Caption         =   "Draw"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
   End
   Begin VB.OptionButton OptPtFormCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   1080
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
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   1935
      Begin VB.TextBox TxtGColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtLColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtFColr 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Frame:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   495
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
   End
   Begin VB.CommandButton BtnStartStop 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox TxtLifetime 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox TxtDoevents 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox PBGOL 
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   6  'Innen ausgefüllt
      FillColor       =   &H000000FF&
      ForeColor       =   &H00004000&
      Height          =   8175
      Left            =   2160
      ScaleHeight     =   8115
      ScaleWidth      =   8115
      TabIndex        =   16
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label6 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "SizeY:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "SizeX:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label LblLifeRule 
      Alignment       =   2  'Zentriert
      Caption         =   "/"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1935
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
End
Attribute VB_Name = "FrmGameOfLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '2008_08_06 Zeilen:  124
Private m_GOL     As TGameOfLife
Private m_Timer   As Single
Private m_GOLFile As TGOLFile

Private Sub Form_Load()
    'einige Voreinstellungen vornehmen
    m_GOLFile.FileName = "C:\Test.golg"
    Me.PBGOL.AutoRedraw = True
    Me.PBGOL.ScaleMode = vbPixels '3 ' Pixel
    Me.TxtDoevents.Text = CStr(20) 'alle x Generationen ein Doevents
    Me.TxtLifetime.Text = CStr(0) 'in Millisekunden
    Me.TxtX.Text = CStr(180)
    Me.TxtY.Text = CStr(180)
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
    Call MGameOfLife.InitRandom(m_GOL, 5) '5 % Populationsdichte
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
Try: On Error GoTo Catch
    Dim d As Double: d = Timer - m_Timer
    m_Timer = Timer
    Me.Caption = "Generations: " & CStr(CountGenerations) & "   " & CStr(CLng(m_GOL.GenTilDoEv / d)) & " FPS"
Catch:
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
Private Sub BtnSetLifeRule_Click()
    If FrmLifeRule.ShowDialog(Me, m_GOL.LifeRule) = vbOK Then
        LblLifeRule.Caption = MLifeRule.LifeRuleToString(m_GOL.LifeRule)
    End If
End Sub
Private Sub BtnCreateNewField_Click()
    Dim x As Long
    Dim y As Long
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
    x = CLng(TxtX.Text)
    y = CLng(TxtY.Text)
    Call MGameOfLife.New_GameOfLife(m_GOL, x, y, MLifeRule.New_LifeRule(LblLifeRule.Caption))
    Call BtnDrawNew_Click
End Sub
Private Sub BtnSaveGeneration_Click()
    Dim FNm As String: FNm = InputBox("FileName: ", "FileName: ", m_GOLFile.FileName)
    If StrPtr(FNm) Then
        m_GOLFile.FileName = FNm
        Call MGOLFile.SaveGeneration(m_GOLFile, m_GOL)
    End If
End Sub
Private Sub BtnLoadGeneration_Click()
    Dim FNm As String: FNm = InputBox("FileName: ", "FileName: ", m_GOLFile.FileName)
    If StrPtr(FNm) Then
        m_GOLFile.FileName = FNm
        If MGOLFile.LoadGeneration(m_GOLFile, m_GOL) Then
            'und anpassen
            Call BtnDrawNew_Click
        End If
    End If
End Sub
Private Sub BtnRandom_Click()
    Dim sr As String
    Dim d  As Long: d = m_GOL.Density
    sr = InputBox("Please give the population density in %", "Population Density", CStr(d))
    If StrPtr(sr) <> 0 Then
        If Len(sr) > 0 Then
            If IsNumeric(sr) Then
                d = CLng(sr)
                Call MGameOfLife.InitRandom(m_GOL, d)
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

Private Sub PBGOL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt As Point: pt = MField.GetIndexPoint(m_GOL.Field, x, y)
    Call MGameOfLife.SwitchLife(m_GOL, pt)
    With m_GOL
        Call MGeneration.DrawIndividual(.pThis.Generation, .Field, Me.PBGOL, pt)
    End With
    Me.PBGOL.Refresh
End Sub
