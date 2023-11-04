VERSION 5.00
Begin VB.Form frm_Worldsize 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Die Welt verändern?"
   ClientHeight    =   3360
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   4470
   Icon            =   "frm_Worldsize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox txt_Birthcontrol_Rule 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   3780
      TabIndex        =   9
      Text            =   "3"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txt_Rule_Of_Life 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   "23"
      Top             =   2400
      Width           =   615
   End
   Begin VB.OptionButton opt_BorderType 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "&grenzenlose Life Welt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Value           =   -1  'True
      Width           =   4095
   End
   Begin VB.OptionButton opt_BorderType 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "&begrenzte Life Welt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1620
      Width           =   4095
   End
   Begin VB.TextBox txt_CellSize 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Text            =   "5"
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txt_Worldsize 
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Text            =   "100"
      Top             =   480
      Width           =   675
   End
   Begin VB.CommandButton cmd_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3180
      TabIndex        =   11
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   4215
      Begin VB.OptionButton opt_CellRunded 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "&Quadratische Zellen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   1
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   4095
      End
      Begin VB.OptionButton opt_CellRunded 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "&Runde Zellen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   1
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   4095
      End
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3780
      X2              =   3600
      Y1              =   2460
      Y2              =   2700
   End
   Begin VB.Label Label2 
      Caption         =   "Regel&werk  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   2595
   End
   Begin VB.Line Line2 
      X1              =   60
      X2              =   4380
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4380
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "&Größe der Zellen in Pixel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   3255
   End
   Begin VB.Label lblworldSize 
      Caption         =   "&Ausdehnung der Life Welt in Zellen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "frm_Worldsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_intWorldSize As Integer
Private m_intCellSize As Integer

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    m_intWorldSize = frm_GoL.gol.WorldSize
    m_intCellSize = frm_GoL.gol.CellSize
    txt_Worldsize.Text = m_intWorldSize
    txt_CellSize.Text = m_intCellSize
    txt_Rule_Of_Life.Text = Split(frm_GoL.gol.RulesDefinition, "/")(0)
    txt_Birthcontrol_Rule.Text = Split(frm_GoL.gol.RulesDefinition, "/")(1)
    opt_CellRunded(1).Value = frm_GoL.gol.CellsRounded
    opt_CellRunded(0).Value = Not opt_CellRunded(1).Value
    opt_BorderType(frm_GoL.gol.BorderType).Value = True
End Sub

Private Sub cmd_OK_Click()
Dim i As Byte
    frm_GoL.gol.WorldSize = m_intWorldSize
    frm_GoL.gol.CellSize = m_intCellSize
    frm_GoL.gol.CellsRounded = opt_CellRunded(1).Value
    frm_GoL.gol.RulesDefinition = Trim(txt_Rule_Of_Life.Text) & "/" & Trim(txt_Birthcontrol_Rule.Text)
    
    For i = 0 To opt_BorderType.UBound
        If opt_BorderType(i).Value = True Then
            frm_GoL.gol.BorderType = i
            Exit For
        End If
    Next i
    Unload Me
End Sub

Private Sub txt_Birthcontrol_Rule_Validate(Cancel As Boolean)
    If Len(txt_Birthcontrol_Rule.Text) > 0 Then
        If IsNumeric(Trim(txt_Birthcontrol_Rule.Text)) Then
            If Not Val(txt_Birthcontrol_Rule.Text) > 0 Then
                txt_Birthcontrol_Rule = Split(frm_GoL.gol.RulesDefinition, "/")(1)
            End If
        Else
            txt_Birthcontrol_Rule = Split(frm_GoL.gol.RulesDefinition, "/")(1)
        End If
    Else
        txt_Birthcontrol_Rule = Split(frm_GoL.gol.RulesDefinition, "/")(1)
    End If
End Sub

Private Sub txt_CellSize_Change()
    m_intCellSize = Val(txt_CellSize.Text)
    txt_CellSize.Text = m_intCellSize
End Sub

Private Sub txt_CellSize_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txt_Worldsize.SetFocus
    End Select
End Sub

Private Sub txt_Rule_Of_Life_Validate(Cancel As Boolean)
    If Len(txt_Rule_Of_Life.Text) > 0 Then
        If IsNumeric(Trim(txt_Rule_Of_Life.Text)) Then
            If Not Val(txt_Rule_Of_Life.Text) > 0 Then
                txt_Rule_Of_Life = Split(frm_GoL.gol.RulesDefinition, "/")(0)
            End If
        Else
            txt_Rule_Of_Life = Split(frm_GoL.gol.RulesDefinition, "/")(0)
        End If
    Else
        txt_Rule_Of_Life = Split(frm_GoL.gol.RulesDefinition, "/")(0)
    End If
End Sub

Private Sub txt_Worldsize_Change()
    m_intWorldSize = Val(txt_Worldsize.Text)
    txt_Worldsize.Text = m_intWorldSize
End Sub

Private Sub txt_Worldsize_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmd_OK.SetFocus
    End Select
End Sub
