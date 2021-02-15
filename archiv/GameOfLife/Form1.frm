VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Conway's Game Of Life "
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtY 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox TxtX 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton BtnRandom 
      Caption         =   "Randomize"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton BtnDrawNew 
      Caption         =   "Neuzeichnen"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton BtnStartStop 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox PBGOL 
      BackColor       =   &H00004000&
      FillColor       =   &H000000FF&
      ForeColor       =   &H00004000&
      Height          =   6975
      Left            =   1320
      ScaleHeight     =   6915
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '2008_07_31 Zeilen:  60
Private m_GOL As TGameOfLife
Private m_Timer As Double

Private Sub Form_Load()
    TxtX.Text = CStr(120)
    TxtY.Text = CStr(120)
    Call BtnRandom_Click
End Sub
Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    With Me.PBGOL
        L = .Left:             T = .Top
        W = Me.ScaleWidth - L: H = Me.ScaleHeight - T
        If W > 0 And H > 0 Then Call .Move(L, T, W, H)
    End With
    Call MGameOfLife.CalcField(m_GOL, Me.PBGOL)
    Me.PBGOL.Cls
    Call MGameOfLife.DrawAll(m_GOL, Me.PBGOL)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    m_GOL.bIsRunning = False
    Call MGameOfLife.Delete(m_GOL)
End Sub
Public Sub SetCounter(ByVal CountGenerations As Long)
    Dim d As Double: d = Timer - m_Timer
    m_Timer = Timer
    Me.Caption = "Generations: " & CStr(CountGenerations) & "   " & CStr(CLng(50 / d)) & " s^-1"
End Sub
Private Sub BtnStartStop_Click()
    m_GOL.bIsRunning = Not m_GOL.bIsRunning
    If m_GOL.bIsRunning Then
        m_Timer = Timer
        BtnStartStop.Caption = "Pause"
        Call MGameOfLife.Run(m_GOL, Me.PBGOL)
    Else
        BtnStartStop.Caption = "Start"
        Me.Caption = "Conway's Game Of Life "
    End If
End Sub
Private Sub BtnDrawNew_Click()
    Call MGameOfLife.CalcField(m_GOL, Me.PBGOL)
    Me.PBGOL.Cls
    Call MGameOfLife.DrawAll(m_GOL, Me.PBGOL)
End Sub
Private Sub BtnRandom_Click()
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
    Call MGameOfLife.New_GameOfLife(m_GOL, X, Y, 10)
    Call MGameOfLife.InitRandom(m_GOL, 10)
    Call BtnDrawNew_Click
End Sub
