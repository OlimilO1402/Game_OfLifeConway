VERSION 5.00
Begin VB.Form frm_QuickHelp 
   Caption         =   "kleine Hilfe zu GoL"
   ClientHeight    =   6720
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   10725
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10725
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6195
      Left            =   0
      Picture         =   "frmHelp.frx":030A
      ScaleHeight     =   6135
      ScaleWidth      =   10515
      TabIndex        =   1
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton cmd_Rotate 
         DownPicture     =   "frmHelp.frx":0614
         Height          =   480
         Left            =   3000
         Picture         =   "frmHelp.frx":093D
         Style           =   1  'Grafisch
         TabIndex        =   8
         ToolTipText     =   "Figur um 90° nach rechts drehen  (Life Welt -> Rechtsklick)"
         Top             =   3180
         Width           =   480
      End
      Begin VB.CommandButton cmd_RemoveFigur 
         DownPicture     =   "frmHelp.frx":0C66
         Height          =   480
         Left            =   660
         Picture         =   "frmHelp.frx":0F8E
         Style           =   1  'Grafisch
         TabIndex        =   6
         ToolTipText     =   "Figur aus der Auswahlliste entfernen."
         Top             =   2580
         Width           =   480
      End
      Begin VB.CommandButton cmd_AddFigur 
         DownPicture     =   "frmHelp.frx":12B6
         Height          =   480
         Left            =   660
         Picture         =   "frmHelp.frx":14F2
         Style           =   1  'Grafisch
         TabIndex        =   4
         ToolTipText     =   "aktuelle Population als Figur in die Auswahlliste übernehmen!"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblTipLink 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "http://de.wikipedia.org/wiki/Game_of_Life"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   3960
         TabIndex        =   14
         Top             =   1380
         Width           =   4770
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelp.frx":172E
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   8
         Left            =   660
         TabIndex        =   13
         Top             =   660
         Width           =   9495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menü: ""Optionen -> Die &Welt verändern..."","
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
         Index           =   7
         Left            =   660
         TabIndex        =   12
         Top             =   420
         Width           =   8235
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voreinstellungen für die Life Welt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   660
         TabIndex        =   11
         Top             =   120
         Width           =   4020
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   2880
         Y1              =   5820
         Y2              =   4980
      End
      Begin VB.Shape Shape1 
         Height          =   555
         Left            =   420
         Shape           =   2  'Oval
         Top             =   5580
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   2280
         X2              =   1920
         Y1              =   5220
         Y2              =   4320
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   1500
         Y1              =   4320
         Y2              =   5280
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelp.frx":18E7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Index           =   5
         Left            =   2820
         TabIndex        =   10
         Top             =   4800
         Width           =   8235
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelp.frx":19AD
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   660
         TabIndex        =   9
         Top             =   4080
         Width           =   9750
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelp.frx":1A72
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   660
         TabIndex        =   7
         Top             =   3240
         Width           =   9480
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Figuren aus der Auswahlliste entfernen!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Top             =   2700
         Width           =   4800
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmHelp.frx":1B0B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   2100
         Width           =   8475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Figuren erstellen und in Figurauswahlliste Speichern."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   660
         TabIndex        =   2
         Top             =   1680
         Width           =   6375
      End
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   420
         Picture         =   "frmHelp.frx":1BBA
         Top             =   4740
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   9420
      TabIndex        =   0
      Top             =   6300
      Width           =   1215
   End
End
Attribute VB_Name = "frm_QuickHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                                                    ByVal hwnd As Long, _
                                                    ByVal Operation As String, _
                                                    ByVal FileName As String, _
                                                    Optional ByVal Parameters As String, _
                                                    Optional ByVal Directory As String, _
                                                    Optional ByVal WindowStyle As Long = vbMinimizedFocus _
                                                    ) As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub lblTipLink_Click(Index As Integer)
    ShellExecute 0, "Open", "http://de.wikipedia.org/wiki/Game_of_Life"
End Sub

