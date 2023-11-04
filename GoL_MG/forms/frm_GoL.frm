VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_GoL 
   AutoRedraw      =   -1  'True
   Caption         =   "Game of Life  (zelluläre Automaten nach John Horton Conway) @2008 marco@grossert.com"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   10170
   Icon            =   "frm_GoL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   628
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   678
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.PictureBox picTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   480
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   20
      Top             =   9540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fra_Controls 
      BorderStyle     =   0  'Kein
      Height          =   7815
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2010
      Begin VB.CheckBox chk_Stop 
         DownPicture     =   "frm_GoL.frx":030A
         Height          =   480
         Left            =   960
         Picture         =   "frm_GoL.frx":07DF
         Style           =   1  'Grafisch
         TabIndex        =   34
         ToolTipText     =   "Life Welt anhalten"
         Top             =   1320
         Value           =   1  'Aktiviert
         Width           =   480
      End
      Begin VB.CheckBox chk_Start 
         DownPicture     =   "frm_GoL.frx":0CB4
         Height          =   480
         Left            =   0
         Picture         =   "frm_GoL.frx":119C
         Style           =   1  'Grafisch
         TabIndex        =   33
         ToolTipText     =   "Life Welt starten [F5]"
         Top             =   1320
         Width           =   480
      End
      Begin VB.CommandButton cmd_PlaySteps 
         DownPicture     =   "frm_GoL.frx":1684
         Height          =   480
         Left            =   480
         Picture         =   "frm_GoL.frx":1BB5
         Style           =   1  'Grafisch
         TabIndex        =   0
         ToolTipText     =   "Start mit Autostop [F8] oder [Leertaste]  Rechtsklick um die Anzahl generationen bis zum Stop festzulegen!"
         Top             =   1320
         Width           =   480
      End
      Begin MSComDlg.CommonDialog cdlgLoadSave 
         Left            =   1380
         Top             =   6780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Clear 
         DownPicture     =   "frm_GoL.frx":20E6
         Height          =   480
         Left            =   1500
         Picture         =   "frm_GoL.frx":25FB
         Style           =   1  'Grafisch
         TabIndex        =   1
         ToolTipText     =   "Population löschen "
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox pic_Tools 
         Height          =   3915
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   1935
         TabIndex        =   26
         Top             =   1920
         Width           =   1995
         Begin VB.OptionButton opt_Tools 
            DownPicture     =   "frm_GoL.frx":2B10
            Height          =   480
            Index           =   5
            Left            =   480
            MouseIcon       =   "frm_GoL.frx":3026
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":3330
            Style           =   1  'Grafisch
            TabIndex        =   8
            ToolTipText     =   "Auswahlbereich"
            Top             =   720
            Width           =   480
         End
         Begin VB.CommandButton cmd_AddFigur 
            DownPicture     =   "frm_GoL.frx":3846
            Height          =   480
            Left            =   960
            Picture         =   "frm_GoL.frx":3A82
            Style           =   1  'Grafisch
            TabIndex        =   32
            ToolTipText     =   "aktuelle Population als Figur in die Auswahlliste übernehmen!"
            Top             =   1380
            Width           =   480
         End
         Begin VB.CommandButton cmd_RemoveFigur 
            DownPicture     =   "frm_GoL.frx":3CBE
            Height          =   480
            Left            =   1440
            Picture         =   "frm_GoL.frx":3FE6
            Style           =   1  'Grafisch
            TabIndex        =   31
            ToolTipText     =   "Figur aus der Auswahlliste entfernen."
            Top             =   1380
            Width           =   480
         End
         Begin VB.OptionButton opt_Tools 
            DisabledPicture =   "frm_GoL.frx":430E
            DownPicture     =   "frm_GoL.frx":4623
            Height          =   480
            Index           =   3
            Left            =   480
            MouseIcon       =   "frm_GoL.frx":4926
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":4A78
            Style           =   1  'Grafisch
            TabIndex        =   4
            ToolTipText     =   "Invertieren Pinsel"
            Top             =   240
            Width           =   480
         End
         Begin VB.CommandButton cmd_Rotate 
            DownPicture     =   "frm_GoL.frx":4D7B
            Height          =   480
            Left            =   1440
            Picture         =   "frm_GoL.frx":50A4
            Style           =   1  'Grafisch
            TabIndex        =   11
            ToolTipText     =   "Figur um 90° nach rechts drehen  (Life Welt od. Figurvorschau -> Rechtsklick)"
            Top             =   720
            Width           =   480
         End
         Begin VB.PictureBox pic_Figures 
            BackColor       =   &H8000000C&
            Height          =   2055
            Left            =   0
            ScaleHeight     =   1995
            ScaleWidth      =   1875
            TabIndex        =   27
            Top             =   1860
            Width           =   1935
            Begin VB.PictureBox pic_Figure 
               Appearance      =   0  '2D
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               FillStyle       =   0  'Ausgefüllt
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   0
               ScaleHeight     =   1545
               ScaleWidth      =   1785
               TabIndex        =   12
               ToolTipText     =   "Einfügen mit linker Maustaste, Routieren um 90° mit rechter Maustaste "
               Top             =   360
               Width           =   1815
               Begin GameOfLife.ctrl_MenuPictures ctrl_Menupic 
                  Left            =   1320
                  Top             =   1080
                  _ExtentX        =   847
                  _ExtentY        =   847
               End
            End
            Begin VB.ComboBox cmbFigurs 
               Appearance      =   0  '2D
               ForeColor       =   &H00800000&
               Height          =   315
               ItemData        =   "frm_GoL.frx":53CD
               Left            =   0
               List            =   "frm_GoL.frx":53D4
               Style           =   2  'Dropdown-Liste
               TabIndex        =   10
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.OptionButton opt_Tools 
            DisabledPicture =   "frm_GoL.frx":53E3
            DownPicture     =   "frm_GoL.frx":56E6
            Height          =   480
            Index           =   4
            Left            =   0
            MouseIcon       =   "frm_GoL.frx":59FB
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":5B4D
            Style           =   1  'Grafisch
            TabIndex        =   3
            ToolTipText     =   "Zelle Invertieren"
            Top             =   240
            Width           =   480
         End
         Begin VB.OptionButton opt_Tools 
            DisabledPicture =   "frm_GoL.frx":5E62
            DownPicture     =   "frm_GoL.frx":6177
            Height          =   480
            Index           =   2
            Left            =   960
            MouseIcon       =   "frm_GoL.frx":649D
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":65EF
            Style           =   1  'Grafisch
            TabIndex        =   5
            ToolTipText     =   "Zeichnen Pinsel"
            Top             =   240
            Width           =   480
         End
         Begin VB.OptionButton opt_Tools 
            DownPicture     =   "frm_GoL.frx":6915
            Height          =   480
            Index           =   1
            Left            =   1440
            MouseIcon       =   "frm_GoL.frx":6C0E
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":6F18
            Style           =   1  'Grafisch
            TabIndex        =   6
            ToolTipText     =   "Radieren"
            Top             =   240
            Width           =   480
         End
         Begin VB.OptionButton opt_Tools 
            DownPicture     =   "frm_GoL.frx":7211
            Height          =   480
            Index           =   0
            Left            =   0
            MouseIcon       =   "frm_GoL.frx":7516
            MousePointer    =   4  'Symbol
            Picture         =   "frm_GoL.frx":7820
            Style           =   1  'Grafisch
            TabIndex        =   7
            ToolTipText     =   "Figurauswahl"
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblFigurs 
            BackStyle       =   0  'Transparent
            Caption         =   "&Figuren:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1755
            Left            =   0
            TabIndex        =   9
            Top             =   1380
            Width           =   1875
         End
         Begin VB.Label lbl_Tools 
            BackStyle       =   0  'Transparent
            Caption         =   "&Werkzeuge:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   60
            TabIndex        =   2
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox pic_Display 
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         Height          =   1275
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   2175
         TabIndex        =   22
         Top             =   0
         Width           =   2235
         Begin VB.TextBox txtInt 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   1980
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "10"
            ToolTipText     =   "[Umschal] + [Pfeil Hoch] oder [+] für Interval + 1 /  [Umschalt] + [Pfeil nach Unten] oder [-] für Interval - 1"
            Top             =   1080
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lbl_Interval 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            ToolTipText     =   "[Umschalt] + [Pfeil nach Oben] oder [+] für Interval + 1 /  [Umschalt] + [Pfeil nach Unten] oder [-] für Interval - 1"
            Top             =   960
            Width           =   435
         End
         Begin VB.Label lbl_TimerInt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " &Intervall:     ms"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   0
            TabIndex        =   29
            ToolTipText     =   "[Umschalt] + [Pfeil nach Oben] oder [+] für Interval + 1 /  [Umschalt] + [Pfeil nach Unten] oder [-] für Interval - 1"
            Top             =   960
            Width           =   2010
         End
         Begin VB.Label lbl_Generation 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "Generation: 0"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1980
         End
         Begin VB.Label lbl_FPS 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "0 FPS"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label lbl_Coordinates 
            Alignment       =   1  'Rechts
            BackStyle       =   0  'Transparent
            Caption         =   "X:0 / Y:0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   555
            Left            =   0
            TabIndex        =   23
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   1980
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmd_Quit 
         Cancel          =   -1  'True
         Caption         =   "&Beenden"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   7320
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frm_GoL.frx":7B25
         ForeColor       =   &H00800000&
         Height          =   1695
         Left            =   180
         TabIndex        =   15
         Top             =   5820
         Width           =   1575
      End
   End
   Begin VB.PictureBox pic_Scroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   7815
      Left            =   1980
      ScaleHeight     =   517
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   16
      Top             =   0
      Width           =   8235
      Begin VB.PictureBox picFill 
         BorderStyle     =   0  'Kein
         Height          =   435
         Left            =   7920
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   21
         Top             =   7440
         Width           =   315
      End
      Begin VB.HScrollBar hscScrollX 
         Height          =   255
         LargeChange     =   10
         Left            =   0
         SmallChange     =   10
         TabIndex        =   19
         Top             =   7500
         Width           =   7875
      End
      Begin VB.VScrollBar vscScrollY 
         Height          =   7455
         LargeChange     =   10
         Left            =   7920
         SmallChange     =   10
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic_World 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         DrawStyle       =   1  'Strich
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H00FF0000&
         Height          =   7455
         Left            =   60
         MouseIcon       =   "frm_GoL.frx":7BC5
         ScaleHeight     =   497
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   513
         TabIndex        =   17
         Top             =   0
         Width           =   7695
         Begin VB.Shape shp_Select 
            BorderStyle     =   3  'Punkt
            Height          =   195
            Left            =   3060
            Top             =   2640
            Visible         =   0   'False
            Width           =   375
         End
      End
   End
   Begin VB.Menu mnulife 
      Caption         =   "&Life"
      Begin VB.Menu mnuStartStop 
         Caption         =   "&S&tart"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_AutoStopPlay 
         Caption         =   "Start mit &Autostop"
         Shortcut        =   {F8}
      End
      Begin VB.Menu tr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_CounterReset 
         Caption         =   "Generationen -Zähler &Rücksetzen"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_Random 
         Caption         =   "&zufällige Start-Population"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu_Clear 
         Caption         =   "Population l&öschen"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu tr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_LoadFigur 
         Caption         =   "&Figur laden..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu tr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_LoadPopulation 
         Caption         =   "Population &laden..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_SavePopulation 
         Caption         =   "Population &speichern unter..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu tr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnu_EditCopy 
         Caption         =   "&Kopieren"
         Enabled         =   0   'False
         Begin VB.Menu mnu_EditCopyArea 
            Caption         =   "&Auswahlrechteck"
         End
         Begin VB.Menu mnu_EditCopyFigur 
            Caption         =   "&Figur"
         End
      End
      Begin VB.Menu mnu_EditCut 
         Caption         =   "&Ausschneiden"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_EditInsert 
         Caption         =   "&Einfügen"
         Enabled         =   0   'False
         Begin VB.Menu mnu_EditInsertLeftTop 
            Caption         =   "&Links && Oben "
         End
         Begin VB.Menu mnu_EditInsertCenter 
            Caption         =   "&Zentriert"
         End
      End
      Begin VB.Menu Tr6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EditRotate 
         Caption         =   "&Drehen um 90° rechts"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "&Werkzeuge"
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "&Figurauswahl"
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "&Radieren"
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "&Zeichnen Pinsel"
         Index           =   2
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "Invertieren &Pinsel"
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "&Zelle Invertieren"
         Checked         =   -1  'True
         Index           =   4
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_Tools_Arr 
         Caption         =   "&Auswahlbereich Aktivieren"
         Index           =   5
      End
      Begin VB.Menu mnu_Tools_Rotate 
         Caption         =   "Figur um 90°  nach rechts &Drehen"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Optionen"
      Begin VB.Menu mnu_Grid 
         Caption         =   "&Raster anzeigen"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu Tr5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_WorldSize 
         Caption         =   "Die &Welt verändern..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnu_SetAutostop 
         Caption         =   "&Autostop Einstellung..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_SetInterval 
         Caption         =   "Generation -&Lebensdauer(Interval) ändern..."
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnu_Question 
      Caption         =   "&?"
      Begin VB.Menu mnu_Help 
         Caption         =   "&Hilfe"
      End
      Begin VB.Menu mnu_Info 
         Caption         =   "&Über"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frm_GoL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Modul       : frm_GoL
' Datum/Zeit  : 05.08.2008 10:00
' Autor       : Marco Großert
' Zweck       : game of life       (Spaßprojekt initiiert durch einen gleichnamigen
'                                   Forumthread im www.ActiveVB.de VB5/VB6 Forum)
'---------------------------------------------------------------------------------------

Option Explicit
'Private Declare Function Rectangle Lib "gdi32.dll" ( _
'                                                ByVal hdc As Long, _
'                                                ByVal X1 As Long, _
'                                                ByVal Y1 As Long, _
'                                                ByVal X2 As Long, _
'                                                ByVal Y2 As Long) As Long

Private Declare Function RoundRect Lib "gdi32.dll" ( _
                                                ByVal hdc As Long, _
                                                ByVal X1 As Long, _
                                                ByVal Y1 As Long, _
                                                ByVal X2 As Long, _
                                                ByVal Y2 As Long, _
                                                ByVal X3 As Long, _
                                                ByVal Y3 As Long) As Long
                                        

Private m_X_Start As Single         'X Cache für Mausaktionen
Private m_Y_Start As Single         'Y Cache für Mausaktionen


Private m_SelectArea As RECT

Private m_lng_LX As Long        'Letzte X Position an der per Klick eine Aktion durchgeführt wurde
Private m_lng_LY As Long        'Letzte Y Position an der per Klick eine Aktion durchgeführt wurde

Private m_str_GolPFileName As String
Private m_str_GolFFileName As String

Private m_int_SelectedTool As Integer
Private m_bln_FigurListChanged As Boolean

Public WithEvents gol As clsGameOfLife
Attribute gol.VB_VarHelpID = -1

Private Sub cmd_AddFigur_Click()
    Dim fN As Double
    Dim tmpLine As String
    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim rSrc As RECT
    Dim rDst As RECT
    Dim Figurname As String
    Dim tmpFigurs() As String
    fN = FreeFile
    On Error GoTo ErrHandler


    ' Erste Spalte mit Lebender Zelle suchen
    With rSrc
        .Left = 0
        .Right = gol.WorldSize
        .Top = 0
        .Bottom = gol.WorldSize
    End With
    If shp_Select.Visible Then
        rSrc = m_SelectArea
    End If
    rDst = gol.GetFigurRect(rSrc)

    With rDst

        If (.Bottom - .Top > 60) Or (.Right - .Left > 60) Then
            Call MsgBox("Die Größe der aktuellen Population überschreitet die maximale Figurgröße(60x60 Zellen)!" _
                        & vbCrLf & "Wenn Sie die aktuelle Population dennoch sichern möchten," _
                        & vbCrLf & "so können Sie diese als """"Game of Life Population *.golp"""", " _
                        & vbCrLf & "im Menü: ""&Life -> Population &speichern unter..."" , speichern!" _
                        , vbExclamation, App.Title)
    
            Exit Sub
        End If
    
        Figurname = InputBox("Unter welchem namen soll die Figur in der Auswahlliste gespeichert werden?")
        If Len(Trim(Figurname)) < 1 Then Exit Sub
    
        tmpLine = vbNullString
        For Y = .Top To .Bottom
            For X = .Left To .Right
                tmpLine = tmpLine & gol.GetCell(X, Y) & "|"
            Next X
            tmpLine = tmpLine & ","
        Next Y
    End With 'rDst
    
    tmpLine = Replace(tmpLine, "|,", ",")
    tmpLine = Mid(tmpLine, 1, Len(tmpLine) - 1)
    ReDim tmpFigurs(0 To UBound(g_str_Figurs, 1) + 1, 0 To 1)
    For i = 0 To UBound(g_str_Figurs, 1)
        tmpFigurs(i, 0) = g_str_Figurs(i, 0)
        tmpFigurs(i, 1) = g_str_Figurs(i, 1)
    Next i
    tmpFigurs(i, 0) = Figurname
    tmpFigurs(i, 1) = tmpLine
    g_str_Figurs = tmpFigurs
    cmbFigurs.AddItem tmpFigurs(i, 0)
    cmbFigurs.ItemData(cmbFigurs.ListCount - 1) = i
    cmbFigurs.ListIndex = cmbFigurs.ListCount - 1
    m_bln_FigurListChanged = True

    Exit Sub
ErrHandler:

End Sub

Private Sub cmd_RemoveFigur_Click()
    On Error Resume Next
    Me.MousePointer = vbHourglass
    Dim toDel As Integer
    Dim Index As Integer
    
    If cmbFigurs.ListCount > 0 Then
        toDel = cmbFigurs.ItemData(cmbFigurs.ListIndex)
        If MsgBox("Möchten Sie die Figur: """ & g_str_Figurs(toDel, 0) & """ aus der Auswahlliste entfernen?", vbQuestion Or vbYesNo) = vbYes Then
            g_str_Figurs(toDel, 0) = "delete"
            g_str_Figurs(toDel, 1) = "delete"
            Index = cmbFigurs.ListIndex
            cmbFigurs.RemoveItem Index
            If Index > 1 Then
                cmbFigurs.ListIndex = Index - 1
            ElseIf cmbFigurs.ListCount > 0 Then
                cmbFigurs.ListIndex = 0
            Else
                cmd_RemoveFigur.Enabled = False
            End If
            m_bln_FigurListChanged = True
        End If
    End If
    Me.MousePointer = vbDefault

End Sub

Private Sub GoL_Progress(ByVal Generation As Long, ByVal FPS As Single)
    lbl_Generation.Caption = "Generation: " & Generation
    lbl_FPS.Caption = Round(FPS, 1) & " FPS"
End Sub

Private Sub GoL_Started(ByVal AutoStop As Boolean)
    chk_Start.Value = 1
    chk_Stop.Value = 0
    DoEvents: Sleep 1
    If gol.Generation >= &H7FFFFFFF Then
        gol.Generation = 0
    End If
    mnuStartStop.Caption = "&Stop"
End Sub

Private Sub GoL_Stopped(ByVal AutoStop As Boolean)
    mnuStartStop.Caption = "&Start"
    chk_Start.Value = 0
    chk_Stop.Value = 1
    If AutoStop Then
        cmd_PlaySteps.SetFocus
    Else
        chk_Start.SetFocus
    End If
End Sub

Private Sub mnu_AutoStopPlay_Click()
    gol.WorldStop
    chk_Stop.SetFocus
    gol.GoSteps
    cmd_PlaySteps.SetFocus
End Sub

Private Sub cmd_PlaySteps_Click()
    mnu_AutoStopPlay_Click
End Sub

Private Sub cmd_PlaySteps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        mnu_SetAutostop_Click
    End If
End Sub

Private Sub mnu_EditCopy_Click()
    'g_byte_Figur = gol.Copy_Rect(m_SelectArea)
    'Call DrawFigurePreview
End Sub

Private Sub mnu_EditCopyArea_Click()
    g_byte_Figur = gol.Copy_Rect(m_SelectArea)
    Call DrawFigurePreview
End Sub

Private Sub mnu_EditCopyFigur_Click()
    g_byte_Figur = gol.Copy_Rect(gol.GetFigurRect(m_SelectArea))
    Call DrawFigurePreview
End Sub

Private Sub mnu_EditCut_Click()
    g_byte_Figur = gol.Copy_Rect(m_SelectArea)
    gol.Clear_Rect m_SelectArea
    'gol.Fill_Rect m_SelectArea
    Call DrawFigurePreview
End Sub

Private Sub mnu_EditInsert_Click()
    If Not (mnu_EditInsertCenter.Visible Or mnu_EditInsertLeftTop.Visible) Then
        gol.InsertFigur2Rect g_byte_Figur, m_SelectArea, True
    End If
End Sub

Private Sub mnu_EditInsertCenter_Click()
    gol.InsertFigur2Rect g_byte_Figur, m_SelectArea, True
End Sub

Private Sub mnu_EditInsertLeftTop_Click()
    gol.InsertFigur2Rect g_byte_Figur, m_SelectArea, False
End Sub

Private Sub mnu_EditRotate_Click()
    m_SelectArea = gol.Rotate_Rect(m_SelectArea)
    Call ShowSelection
End Sub

Private Sub mnuStartStop_Click()
    If gol.isRunning Then
        chk_Start.SetFocus
        gol.WorldStop
    Else
        chk_Stop.SetFocus
        gol.WorldStart
    End If
End Sub

Private Sub chk_Start_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        chk_Start_MouseDown 1, 0, 0, 0
    End Select
End Sub
Private Sub chk_Start_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chk_Stop.SetFocus
    gol.WorldStart
End Sub

Private Sub chk_Stop_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
        chk_Stop_MouseDown 1, 0, 0, 0
    End Select
End Sub
Private Sub chk_Stop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chk_Start.SetFocus
    gol.WorldStop
End Sub
 
Private Sub mnu_Help_Click()
    frm_QuickHelp.Show
End Sub

Private Sub mnu_Info_Click()
    Call MsgBox("Dieses Programm entstand auf Grund der Faszination des Autors: Marco Großert, " _
                & vbCrLf & "über den Forenthreat zum Thema ""Game of Life"" im VB5/VB6 Forum auf www.ActiveVB.de" _
                & vbCrLf & "Informationen zum Thema: Game of Life entworfen 1970 vom Mathematiker John Horton Conway," _
                & vbCrLf & "sowie zum Thema ""zelluläre Automaten"" auf ""http://de.wikipedia.org/wiki/Game_of_Life"" !" _
                & vbCrLf _
                & vbCrLf & "Autor: Marco Großert " _
                & vbCrLf & "Mail:  marco@grossert.com " _
                & vbCrLf & "©2008 by Marco Großert & Game of Life Algorithm by www.ActiveVB.de " _
                , vbInformation, "Über GoL")
    
End Sub


Private Sub mnu_SetAutostop_Click()
    Dim tmpInput As String
    tmpInput = Val(InputBox("Geben Sie die Anzahl Generationen für Automatischen Stop an:", "Autostop ändern", gol.Steps2Go))
    If tmpInput + gol.Steps2Go = 0 Then
        gol.Steps2Go = 1
    ElseIf tmpInput > 0 Then
        gol.Steps2Go = tmpInput
    End If
End Sub

Private Sub mnu_SetInterval_Click()
    Dim tmpInput As String
    tmpInput = Val(InputBox("Geben Sie den Interval für die Generationslebensdauer in Millisekunden an:", "Lebensdauer ändern", gol.Interval))
    If tmpInput + gol.Interval = 0 Then
        gol.Interval = 10
    ElseIf tmpInput > 0 Then
        gol.Interval = tmpInput
    End If
    txtInt.Text = gol.Interval
End Sub

Private Sub cmd_Clear_Click()
    mnu_Clear_Click
    cmd_Clear.Value = False
End Sub

Private Sub mnu_CounterReset_Click()
    lbl_Generation.Caption = "Generation: 0"
    gol.Generation = 0
End Sub

Private Sub mnu_Grid_Click()
    mnu_Grid.Checked = Not mnu_Grid.Checked
    gol.DrawGrid = mnu_Grid.Checked
    pic_World.Cls
    gol.DrawAll
End Sub

Public Sub DrawFigurePreview()
On Error Resume Next
    ' zeichnet das ganze Spielfeld neu
    Dim X As Long
    Dim X1 As Long
    Dim Y As Long
    Dim Y1 As Long
    Dim ext As Long
    Dim xSize As Long
    Dim ySize As Long
    Dim pSize As Long
    
    ext = 5
    xSize = UBound(g_byte_Figur, 1)
    ySize = UBound(g_byte_Figur, 2)
    pSize = IIf(xSize > ySize, xSize + 1, ySize + 1)
    
    pic_Figure.AutoRedraw = True
    pic_Figure.ScaleMode = 3
    pic_Figure.FillStyle = 0
    pic_Figure.Cls
    pic_Figure.Height = (ext * pSize + 2) * Screen.TwipsPerPixelY
    pic_Figure.Width = (ext * pSize + 2) * Screen.TwipsPerPixelX

    For X = 0 To xSize
        For Y = 0 To ySize
            X1 = X * ext
            Y1 = Y * ext
            pic_Figure.FillColor = gol.GetColor(g_byte_Figur(X, Y))
            If gol.DrawGrid Then
                pic_Figure.ForeColor = RGB(200, 200, 250)
                'Rectangle pic_Figure.hdc, X1, Y1, X1 + ext - 1, Y1 + ext - 1
            Else
                pic_Figure.ForeColor = pic_Figure.FillColor
            End If
            RoundRect pic_Figure.hdc, X1, Y1, X1 + ext, Y1 + ext, 90, 90
        Next
    Next
    pic_Figure.Refresh
  
End Sub

Private Sub Select_Tool(ByVal Index As Byte)
Dim i As Byte
    For i = 0 To 5
        mnu_Tools_Arr(i).Checked = False
    Next i
    shp_Select.Visible = False
    Select Case Index
        Case 1 To 5
            pic_World.MousePointer = vbCustom

            Select Case Index
                Case 5              'Zelle invertieren
                    pic_World.MouseIcon = LoadResPicture(106, vbResCursor)
                Case 4              'Zelle invertieren
                    pic_World.MouseIcon = LoadResPicture(105, vbResCursor)
                Case 3              'Invertieren Stift
                    pic_World.MouseIcon = LoadResPicture(103, vbResCursor)
                Case 2              'Zeichnen Stift
                    pic_World.MouseIcon = LoadResPicture(101, vbResCursor)
                Case 1              'Radieren
                    pic_World.MouseIcon = LoadResPicture(104, vbResCursor)
            End Select
            picTmp.MouseIcon = pic_World.MouseIcon
            picTmp.MousePointer = vbCustom
            pic_World.MousePointer = vbCustom
        Case 0              'Figurauswahl
            pic_World.MousePointer = vbCrosshair
            If Me.Visible Then
                If (Me.ActiveControl.Name = "opt_Tools") Then SendMessage cmbFigurs.hwnd, &H14F, True, 0
            End If

    End Select
    
    opt_Tools(Index).Value = True
    mnu_Tools_Arr(Index).Checked = True

End Sub

Private Sub cmbFigurs_Click()
On Error Resume Next
    Me.MousePointer = vbHourglass
    Dim X As Long
    Dim Y As Long
    Dim xSize As Long
    Dim ySize As Long
    Dim pSize As Long
    Dim tmpLines() As String
    Dim tmpFields() As String
    If Not mnu_Tools_Arr(5).Checked Then
        Select_Tool 0
    End If
    tmpLines = Split(g_str_Figurs(cmbFigurs.ItemData(cmbFigurs.ListIndex), 1), ",")
    tmpFields = Split(tmpLines(0), "|")

    pic_World.MousePointer = vbCrosshair
    ySize = UBound(tmpLines) + 1
    xSize = UBound(tmpFields) + 1
    ReDim g_byte_Figur(0 To xSize - 1, 0 To ySize - 1)
    For Y = 0 To ySize - 1
        If Len(Trim(tmpLines(Y))) > 0 Then
        tmpFields = Split(tmpLines(Y), "|")
        For X = 0 To xSize - 1
            g_byte_Figur(X, Y) = Val(tmpFields(X))
        Next X
        End If
    Next Y
    Call DrawFigurePreview
    If cmbFigurs.ListCount > 0 Then
        cmd_RemoveFigur.Enabled = True
    Else
        cmd_RemoveFigur.Enabled = False
    End If
    Me.MousePointer = vbDefault
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim rSrc As RECT
    Dim rDst As RECT
    'vbLeftButton
    'vbRightButton
    'vbMiddleButton
    '
    'vbShiftMask
    'vbCtrlMask
    'vbAltMask
        
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case vbKeySpace
                    Call cmd_PlaySteps_Click
                    'Call mnuStartStop_Click
                    DoEvents: Sleep 1
                Case vbKeyDelete ', vbKeyBack
                    Call mnu_Clear_Click
            End Select
            
        Case vbShiftMask, vbCtrlMask
            Select Case KeyCode
                Case vbKeyUp, vbKeyAdd       'Interval um Eins erhöhen
                    If gol.Interval + 1 <= 1000 Then
                        txtInt.Text = gol.Interval + 1
                    End If
                Case vbKeyDown, vbKeySubtract    'Interval um Eins verringern
                    If gol.Interval - 1 > 0 Then
                        txtInt.Text = gol.Interval - 1
                    End If
                Case vbKeyRight       'Interval um Zehn erhöhen
                    If gol.Interval + 10 <= 1000 Then
                        txtInt.Text = gol.Interval + 10
                    End If
                Case vbKeyLeft   'Interval um zehn verringern
                    If gol.Interval - 10 > 0 Then
                        txtInt.Text = gol.Interval - 10
                    End If
                Case vbKeyV
                    If Shift = vbCtrlMask Then
                        Me.MousePointer = vbHourglass
                        If shp_Select.Visible Then
                            gol.InsertFigur2Rect g_byte_Figur, m_SelectArea, False
                        Else
                            gol.Insert_Figur g_byte_Figur, m_lng_LX, m_lng_LY
                        End If
                        Me.MousePointer = vbDefault
                    End If
                Case vbKeyC
                    If Shift = vbCtrlMask Then
                        Me.MousePointer = vbHourglass
                        ' Erste Spalte mit Lebender Zelle suchen
                        With rSrc
                            .Left = 0
                            .Right = gol.WorldSize
                            .Top = 0
                            .Bottom = gol.WorldSize
                        End With
                        If shp_Select.Visible Then
                            rSrc = m_SelectArea
                            g_byte_Figur = gol.Copy_Rect(rSrc)
                        Else
                            rDst = gol.GetFigurRect(rSrc)
                            g_byte_Figur = gol.Copy_Rect(rDst)
                        End If

                        Call DrawFigurePreview
 
                        Me.MousePointer = vbDefault
                    End If
                Case vbKeyX
                    If Shift = vbCtrlMask Then
                        Me.MousePointer = vbHourglass
                        ' Erste Spalte mit Lebender Zelle suchen
                        With rSrc
                            .Left = 0
                            .Right = gol.WorldSize
                            .Top = 0
                            .Bottom = gol.WorldSize
                        End With
                        If shp_Select.Visible Then
                            rSrc = m_SelectArea
                            g_byte_Figur = gol.Copy_Rect(rSrc)
                            gol.Clear_Rect rSrc
                        Else
                            rDst = gol.GetFigurRect(rSrc)
                            g_byte_Figur = gol.Copy_Rect(rDst)
                        End If
                        
                        Call DrawFigurePreview
 
                        Me.MousePointer = vbDefault
                    End If
            End Select

        'Case vbCtrlMask
        
        Case vbAltMask
        
        Case Else
        
        End Select
End Sub

Private Sub FillFigursCombo()
Dim i As Integer
Dim figCount As Integer
    figCount = UBound(g_str_Figurs, 1)
    cmbFigurs.Clear
    If figCount > 0 Then
        For i = 0 To figCount
            cmbFigurs.AddItem g_str_Figurs(i, 0)
            cmbFigurs.ItemData(cmbFigurs.ListCount - 1) = i
        Next i
    End If
End Sub

Private Sub mnu_Random_Click()
    'Zufällige Start-Populationen erstellen
    Me.MousePointer = vbHourglass
    gol.Randomize_World
    Me.MousePointer = vbDefault
End Sub

Private Sub mnu_WorldSize_Click()
    frm_Worldsize.Show vbModal
    Me.MousePointer = vbHourglass
    Call gol.Resize_World(gol.CellSize, gol.WorldSize)
    Call Form_Resize
    Me.MousePointer = vbDefault
End Sub

Private Sub mnu_Clear_Click()
    Me.MousePointer = vbHourglass
    If shp_Select.Visible Then
        gol.Clear_Rect m_SelectArea
    Else
        gol.Clear_World
        lbl_Generation.Caption = "Generation: 0"
        gol.Generation = 0
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub mnu_Tools_Arr_Click(Index As Integer)
    opt_Tools_Click Index
    Select_Tool Index
End Sub

Private Sub opt_Tools_Click(Index As Integer)
    opt_Tools(Index).Value = True
    
    Select Case Index
        Case 0              'Figurauswahl
            If Me.Visible Then
                If (Me.ActiveControl.Name = "opt_Tools") Then SendMessage cmbFigurs.hwnd, &H14F, True, 0
            End If
    End Select

    Select_Tool Index
    m_int_SelectedTool = Index
End Sub

Private Sub opt_Tools_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then opt_Tools_Click Index
End Sub

Private Sub pic_Figure_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call pic_Figures_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    fra_Controls.Left = (((Me.Width / Screen.TwipsPerPixelX) - Me.ScaleWidth) / 2)
    fra_Controls.Top = 0
    pic_Scroll.Top = 0

    If (Me.WindowState <> vbMinimized) And (Me.WindowState <> vbMaximized) Then
'        If Me.Width < (Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX) + ((pic_World.ScaleWidth + fra_Controls.Width + fra_Controls.Left) * Screen.TwipsPerPixelX) + 150 Then
'            Me.Width = (Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX) + ((pic_World.ScaleWidth + fra_Controls.Width + fra_Controls.Left) * Screen.TwipsPerPixelX) + 150
'        End If
'        If Me.Height < (Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY) + (pic_World.ScaleHeight * Screen.TwipsPerPixelY) + 100 Then
'            Me.Height = (Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY) + (pic_World.ScaleHeight * Screen.TwipsPerPixelY) + 100
'        End If
        
        If Me.Width < 640 * Screen.TwipsPerPixelX Then
            Me.Width = 640 * Screen.TwipsPerPixelX
        End If
        If Me.Height < 480 * Screen.TwipsPerPixelY Then
            Me.Height = 480 * Screen.TwipsPerPixelY
        End If
    End If
    
    pic_Scroll.Left = fra_Controls.Width + 10
    pic_Scroll.Width = Me.ScaleWidth - pic_Scroll.Left - ((Me.Width / Screen.TwipsPerPixelX) - Me.ScaleWidth) / 2
    pic_Scroll.Height = Me.ScaleHeight - pic_Scroll.Top - ((Me.Width / Screen.TwipsPerPixelX) - Me.ScaleWidth) / 2
    fra_Controls.Height = Me.ScaleHeight - (((Me.ScaleHeight / Screen.TwipsPerPixelY) - Me.ScaleHeight) / 2)
 
    hscScrollX_Change
    vscScrollY_Change
End Sub

Private Sub pic_Scroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pic_Scroll.MouseIcon = LoadResPicture(1001, vbResCursor)
    pic_Scroll.MousePointer = vbCustom
    If Button = vbLeftButton Then
        m_X_Start = X
        m_Y_Start = Y
    End If
End Sub

Private Sub pic_Scroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XDiff As Single
Dim YDiff As Single
On Error Resume Next
    If Button = vbLeftButton Then
        XDiff = m_X_Start - X
        YDiff = m_Y_Start - Y
        If pic_World.Left - XDiff < 0 Then
            pic_World.Left = pic_World.Left - XDiff
        Else
            pic_World.Left = 4
        End If
        If pic_World.Top - YDiff < 0 Then
            pic_World.Top = pic_World.Top - YDiff
        Else
            pic_World.Top = 4
        End If
        If -(pic_World.Left - 4) <= hscScrollX.Max Then
            hscScrollX.Value = -(pic_World.Left - 4)
        Else
            hscScrollX.Value = hscScrollX.Max
        End If
        If -(pic_World.Top - 4) <= vscScrollY.Max Then
            vscScrollY.Value = -(pic_World.Top - 4)
        Else
            vscScrollY.Value = vscScrollY.Max
        End If
        
        m_X_Start = X
        m_Y_Start = Y
        pic_Scroll.MouseIcon = LoadResPicture(1001, vbResCursor)
        pic_Scroll.MousePointer = vbCustom
    End If
    On Error GoTo 0
End Sub

Private Sub pic_Scroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        pic_Scroll.MouseIcon = LoadResPicture(1000, vbResCursor)
        pic_Scroll.MousePointer = vbCustom
    Else
        pic_Scroll.MouseIcon = LoadResPicture(1001, vbResCursor)
        pic_Scroll.MousePointer = vbCustom
    End If
    m_X_Start = 0
    m_Y_Start = 0
End Sub

Private Sub pic_Scroll_Resize()
    vscScrollY.Top = 0
    hscScrollX.Left = 0

    vscScrollY.Height = pic_Scroll.Height - hscScrollX.Height - 4
    vscScrollY.Left = pic_Scroll.ScaleWidth - vscScrollY.Width
    hscScrollX.Width = pic_Scroll.Width - vscScrollY.Width - 4
    hscScrollX.Top = pic_Scroll.ScaleHeight - hscScrollX.Height
    hscScrollX.Min = 0
    vscScrollY.Min = 0
    hscScrollX.Max = pic_World.Width
    vscScrollY.Max = pic_World.Height
    If (pic_World.Width <= pic_Scroll.Width - vscScrollY.Width - 4) Then
        'hscScrollX.Enabled = False
        hscScrollX.Value = 0
    Else
        hscScrollX.Enabled = True
    End If
    If (pic_World.Height <= pic_Scroll.Height - hscScrollX.Height - 4) Then
        'vscScrollY.Enabled = False
        vscScrollY.Value = 0
    Else
        vscScrollY.Enabled = True
    End If
    picFill.Left = vscScrollY.Left
    picFill.Top = hscScrollX.Top
    hscScrollX_Change
    vscScrollY_Change
End Sub

Private Sub pic_Tools_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.MousePointer = vbHourglass
        Call gol.Rotate_Figur(g_byte_Figur)
        Call DrawFigurePreview
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub pic_Figures_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.MousePointer = vbHourglass
        Call gol.Rotate_Figur(g_byte_Figur)
        Call DrawFigurePreview
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub ShowSelection()
    Dim X As Long
    Dim Y As Long
    Dim X1 As Long
    Dim Y1 As Long

    With m_SelectArea
        If .Left < .Right Then
            X = .Left * gol.CellSize
            X1 = (.Right + 1) * gol.CellSize
        Else
            X1 = (.Left + 1) * gol.CellSize
            X = .Right * gol.CellSize
        End If
        If .Top < .Bottom Then
            Y = .Top * gol.CellSize
            Y1 = (.Bottom + 1) * gol.CellSize
        Else
            Y1 = (.Top + 1) * gol.CellSize
            Y = .Bottom * gol.CellSize
        End If
    
        shp_Select.Left = X
        shp_Select.Width = X1 - X
        shp_Select.Top = Y
        shp_Select.Height = Y1 - Y
    End With
End Sub

Private Sub pic_World_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim iX As Long
    Dim iY As Long
    Dim iX1 As Long
    Dim iY1 As Long
    Dim lXLen As Long
    Dim lYLen As Long
    Select Case Shift
        Case vbCtrlMask
            pic_World.MouseIcon = LoadResPicture(1001, vbResCursor)
            pic_World.MousePointer = vbCustom
            Call pic_Scroll_MouseDown(Button, 0, X, Y)
        Case Else
            Select Case Button
                Case vbLeftButton
                    iX = (CLng(X) \ gol.CellSize)
                    iY = (CLng(Y) \ gol.CellSize)
                    m_lng_LX = iX
                    m_lng_LY = iY
                    Select Case m_int_SelectedTool
                        Case 0
                            Me.MousePointer = vbHourglass
                            Call gol.Insert_Figur(g_byte_Figur, iX - 1, iY - 1)
                            Me.MousePointer = vbDefault
                        Case 1  'Radieren
                            gol.SetCell(iX, iY) = 0
                        Case 2  'Zeichnen Stift
                            gol.SetCell(iX, iY) = 1
                        Case 3  'Invertieren Stift
                            gol.SetCell(iX, iY) = IIf(gol.GetCell(iX, iY) >= 1, 0, 1)
                        Case 4  'Zelle invertieren
                            gol.SetCell(iX, iY) = IIf(gol.GetCell(iX, iY) >= 1, 0, 1)
                        Case 5  'Select modus
                            With m_SelectArea
                                .Left = (CLng(X) \ gol.CellSize)
                                .Right = .Left
                                .Top = (CLng(Y) \ gol.CellSize)
                                .Bottom = .Top
                            End With
                            shp_Select.Width = 0
                            shp_Select.Height = 0
                            shp_Select.Visible = False
                    End Select
                Case vbRightButton
                    Select Case m_int_SelectedTool
                        Case 0
                            'Figur drehen
                            Me.MousePointer = vbHourglass
                            Call gol.Rotate_Figur(g_byte_Figur)
                            Call DrawFigurePreview
                            Me.MousePointer = vbDefault
                        Case 5  'Select modus
                            If (X > shp_Select.Left) And (X < shp_Select.Width + shp_Select.Left) And _
                                (Y > shp_Select.Top) And (Y < shp_Select.Height + shp_Select.Top) Then
                                'innerhalb des Auswahlbereichs
                                mnu_EditCopy.Enabled = True
                                mnu_EditCut.Enabled = True
                                mnu_EditInsert.Enabled = True
                                mnu_EditRotate.Enabled = True
                                Print pic_World.Width
                                Call PopupMenu(mnu_Edit, vbPopupMenuLeftAlign, X + pic_World.Left + pic_Scroll.Left, Y + pic_World.Top)
                                 
                            Else
                                'ausserhalb des Auswahlbereichs
                                'Auswahlrechteck drehen
                                With m_SelectArea
                                    If .Left > .Right Then
                                        iX = .Left
                                        .Left = .Right
                                        .Right = iX
                                    End If
                                    If .Top > .Bottom Then
                                        iY = .Top
                                        .Top = .Bottom
                                        .Bottom = iY
                                    End If
                                    lXLen = .Right - .Left
                                    lYLen = .Bottom - .Top
                                    .Left = (.Left + lXLen * 0.5) - (lYLen * 0.5)
                                    .Top = (.Top + lYLen * 0.5) - (lXLen * 0.5)
                                    .Right = .Left + lYLen
                                    .Bottom = .Top + lXLen

                                    iX = .Left * gol.CellSize
                                    iX1 = .Right * gol.CellSize
                                    iY = .Top * gol.CellSize
                                    iY1 = .Bottom * gol.CellSize
                                    
                                    If (Abs(iX1 - iX) > 0) And (Abs(iY1 - iY) > 0) Then
                                        lbl_Coordinates.Caption = "L(" & .Left & ") R(" & .Right & ") B(" & .Right - .Left + 1 & ")" & vbCrLf & _
                                                                  "O(" & .Top & " U(" & .Bottom & ") H(" & .Bottom - .Top + 1 & ")"
                                        mnu_EditCopy.Enabled = True
                                        mnu_EditCut.Enabled = True
                                        mnu_EditInsert.Enabled = True
                                        mnu_EditRotate.Enabled = True
                                        shp_Select.Visible = True
                                        Call ShowSelection
                                    Else
                                        mnu_EditCopy.Enabled = False
                                        mnu_EditCut.Enabled = False
                                        mnu_EditInsert.Enabled = False
                                        mnu_EditRotate.Enabled = False
                                        Call ShowSelection
                                    End If
                                End With 'm_SelectArea
                            End If
                        Case Else

                    End Select
            End Select  'Button
        End Select  'Shift
        If shp_Select.Visible Then
            mnu_Clear.Caption = "Auswahlbereich l&öschen"
        Else
            mnu_Clear.Caption = "Population l&öschen"
        End If
    pic_World.Refresh
End Sub

Private Sub pic_World_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim iX As Long
    Dim iY As Long
    Dim iX1 As Long
    Dim iY1 As Long
    If Shift = vbCtrlMask Then
    
        If Button = vbLeftButton Then
            pic_World.MouseIcon = LoadResPicture(1001, vbResCursor)
            pic_World.MousePointer = vbCustom
        End If
        pic_Scroll.Refresh
        Call pic_Scroll_MouseMove(Button, 0, X, Y)
    Else
        pic_World.MouseIcon = picTmp.MouseIcon
        iX = (CLng(X) \ gol.CellSize)
        iY = (CLng(Y) \ gol.CellSize)
        
        Select Case Button
            Case vbLeftButton
                If (m_lng_LX = iX) And (m_lng_LY = iY) Then Exit Sub
                m_lng_LX = iX
                m_lng_LY = iY
                
                Select Case m_int_SelectedTool
                    Case 1  'Radieren
                        gol.SetCell(iX, iY) = 0
                    Case 2  'Zeichnen Stift
                        gol.SetCell(iX, iY) = 1
                    Case 3  'Invertieren Stift
                        gol.SetCell(iX, iY) = IIf(gol.GetCell(iX, iY) >= 1, 0, 1)
                    Case 4  'Zelle invertieren
                        'GoL.SetCell(iX, iY) = IIf(GoL.GetCell(iX, iY) >= 1, 0, 1)
                    Case 5  'Select Modus
                        With m_SelectArea
                            .Right = (CLng(X) \ gol.CellSize)
                            .Bottom = (CLng(Y) \ gol.CellSize)
           
                            iX = .Left * gol.CellSize
                            iX1 = .Right * gol.CellSize
                            iY = .Top * gol.CellSize
                            iY1 = .Bottom * gol.CellSize
                            
                            .Right = .Right - 1
                            .Bottom = .Bottom - 1
                            
                            If (Abs(iX1 - iX) > 0) And (Abs(iY1 - iY) > 0) Then
                                lbl_Coordinates.Caption = "L(" & .Left & ") R(" & .Right & ") B(" & .Right - .Left + 1 & ")" & vbCrLf & _
                                                          "O(" & .Top & " U(" & .Bottom & ") H(" & .Bottom - .Top + 1 & ")"
                                shp_Select.Visible = True
                                Call ShowSelection
                                mnu_EditCopy.Enabled = True
                                mnu_EditCut.Enabled = True
                                mnu_EditInsert.Enabled = True
                                mnu_EditRotate.Enabled = True
                                cmd_Clear.ToolTipText = "Auswahlbereich leeren"
                                mnu_Clear.Caption = "Auswahlbereich &leeren"
                                cmd_Rotate.ToolTipText = "Auswahlbereich um 90° nach rechts drehen  (Auswahlbereich -> Rechtsklick -> &Drehen um 90° rechts)"
                            Else
                                mnu_EditCopy.Enabled = False
                                mnu_EditCut.Enabled = False
                                mnu_EditInsert.Enabled = False
                                mnu_EditRotate.Enabled = False
                                cmd_Clear.ToolTipText = "Population löschen"
                                mnu_Clear.Caption = "Population &löschen"
                                cmd_Rotate.ToolTipText = "Figur um 90° nach rechts drehen  (Life Welt od. Figurvorschau -> Rechtsklick)"
                            End If
                        End With 'm_SelectArea
                End Select

            Case vbRightButton
                Select Case m_int_SelectedTool
                    Case 0
'                    Case 1     'Radieren
'                    Case 2     'Zeichnen Stift
'                    Case 3     'Invertieren Stift
'                    Case 4     'Zelle invertieren
'                    Case 5     'Select Modus
                    Case Else

                End Select
            
            Case Else
                If Not shp_Select.Visible Then
                    iX = (CLng(X) \ gol.CellSize)
                    iY = (CLng(Y) \ gol.CellSize)
                    lbl_Coordinates.Caption = "X:" & iX & "(" & m_lng_LX & ")" & " / Y:" & iY & "(" & m_lng_LY & ")"
                End If
        End Select
    End If
    If shp_Select.Visible Then
        mnu_Clear.Caption = "Auswahlbereich l&öschen"
    Else
        mnu_Clear.Caption = "Population l&öschen"
    End If

    pic_World.Refresh
End Sub

Private Sub pic_World_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iX As Long
Dim iY As Long
Dim iX1 As Long
Dim iY1 As Long

    Select Case Shift
        Case vbCtrlMask
            If Button = vbLeftButton Then
                pic_World.MouseIcon = LoadResPicture(1000, vbResCursor)
                pic_World.MousePointer = vbCustom
            Else
                pic_World.MouseIcon = picTmp.MouseIcon
                pic_World.MousePointer = vbCustom
            End If
        
            Call pic_Scroll_MouseMove(Button, 0, X, Y)
        Case Else
        
    End Select
    
    Select Case Button
        Case vbLeftButton
            Select Case m_int_SelectedTool
                Case 1  'Radieren
                Case 2  'Zeichnen Stift
                Case 3  'Invertieren Stift
                Case 4  'Zelle invertieren
                Case 5  'Select Modus
                    With m_SelectArea
                        .Right = (CLng(X) \ gol.CellSize)
                        .Bottom = (CLng(Y) \ gol.CellSize)
                        If .Left > .Right Then
                            iX = .Left
                            .Left = .Right
                            .Right = iX
                        End If
                        If .Top > .Bottom Then
                            iY = .Top
                            .Top = .Bottom
                            .Bottom = iY
                        End If

                        
                        iX = .Left * gol.CellSize
                        iX1 = .Right * gol.CellSize
                        iY = .Top * gol.CellSize
                        iY1 = .Bottom * gol.CellSize
                        
                        
                        .Right = .Right - 1
                        .Bottom = .Bottom - 1
                        
                        If (Abs(iX1 - iX) > 0) And (Abs(iY1 - iY) > 0) Then
                            lbl_Coordinates.Caption = "L(" & .Left & ") R(" & .Right & ") B(" & .Right - .Left + 1 & ")" & vbCrLf & _
                                                                  "O(" & .Top & " U(" & .Bottom & ") H(" & .Bottom - .Top + 1 & ")"
                            shp_Select.Visible = True
                            Call ShowSelection
                            mnu_EditCopy.Enabled = True
                            mnu_EditCut.Enabled = True
                            mnu_EditInsert.Enabled = True
                            mnu_EditRotate.Enabled = True
                        Else
                            mnu_EditCopy.Enabled = False
                            mnu_EditCut.Enabled = False
                            mnu_EditInsert.Enabled = False
                            mnu_EditRotate.Enabled = False
                        End If
                    End With 'm_SelectArea
            End Select
        Case vbRightButton
            
    End Select
    If shp_Select.Visible Then
        mnu_Clear.Caption = "Auswahlbereich l&öschen"
    Else
        mnu_Clear.Caption = "Population l&öschen"
    End If

End Sub

Private Sub mnu_Tools_Rotate_Click()
    Call cmd_Rotate_Click
End Sub

Private Sub cmd_Rotate_Click()
    Me.MousePointer = vbHourglass
    If shp_Select.Visible Then
        m_SelectArea = gol.Rotate_Rect(m_SelectArea)
        Call ShowSelection
    Else
        Call gol.Rotate_Figur(g_byte_Figur)
        Call DrawFigurePreview
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub txtInt_Change()
    gol.Interval = Val(txtInt.Text)
    If gol.Interval < 1 Then gol.Interval = 1
    If gol.Interval > 1000 Then gol.Interval = 1000
    txtInt.Text = gol.Interval
    lbl_Interval.Caption = gol.Interval
End Sub

Private Sub hscScrollX_Change()
    pic_World.Left = -hscScrollX.Value + 4
End Sub
Private Sub vscScrollY_Change()
    pic_World.Top = -vscScrollY.Value + 4
End Sub

'Programmstart und Ende------------------------------------------------------------------------------------------------

Private Sub mnuQuit_Click()
    gol.WorldStop
    Unload Me
End Sub

Private Sub cmd_Quit_Click()
    gol.WorldStop
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Caption = "Game of Life(" & App.Major & "." & App.Minor & ")  (zelluläre Automaten nach John Horton Conway) @2008 marco@grossert.com"
    Me.MousePointer = vbHourglass
    g_str_IniFile = fc_AppPath & "GoL.ini"
    
    pic_World.AutoRedraw = True
    pic_World.ScaleMode = 3
    pic_World.FillStyle = 0
    pic_World.DrawStyle = vbInsideSolid
    pic_World.DrawWidth = 1
        
    Set gol = New clsGameOfLife
    Set gol.DestPic = pic_World
    Call GetINISettings(gol)
    mnu_Grid.Checked = gol.DrawGrid

    txtInt.Text = gol.Interval
    Call gol.Resize_World(gol.CellSize, gol.WorldSize)
    gol.DrawGrid = mnu_Grid.Checked
    Call gol.DrawAll
    gol.Generation = 0
    gol.WorldStop
    
    Call LoadFigursList(fc_AppPath & "figurs.golfl")
    Call FillFigursCombo
    
    Call Form_Resize

    DoEvents: Sleep 1
    Me.Show
    'Kleiner Bug oder ich bin zu Blöd es war immer der falsche Eintrag der Combobox ausgewählt
    Call opt_Tools_Click(3)
    Call opt_Tools_Click(4)
    cmbFigurs.ListIndex = 0
    Select_Tool 4
    LoadMenuIcons
    Me.MousePointer = vbDefault
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gol.WorldStop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gol.WorldStop
    If m_bln_FigurListChanged Then Call SaveFigursList(g_str_Figurs)
    Call SaveIniSettings(gol)
    DoEvents: Sleep 100
    Set gol = Nothing
    End
End Sub
'------------------------------------------------------------------------------------------------Programmstart und Ende

'Laden und Speichern---------------------------------------------------------------------------------------------------
Private Sub mnu_SavePopulation_Click()
On Error GoTo Fehler
    If Len(m_str_GolPFileName) = 0 Then
        m_str_GolPFileName = "NewFile.golp"
    End If
    cdlgLoadSave.DialogTitle = "Datei Speichern"
    cdlgLoadSave.Flags = cdlOFNOverwritePrompt
    cdlgLoadSave.InitDir = App.Path
    cdlgLoadSave.FileName = m_str_GolPFileName
    cdlgLoadSave.Filter = "Game of Life Population *.golp|*.golp|Game of Life Figur *.golf|*.golf;"
    cdlgLoadSave.CancelError = True
    cdlgLoadSave.ShowSave   'Standarddialogfeld "Öffnen" anzeigen.
    m_str_GolPFileName = cdlgLoadSave.FileName
    
    If fc_FileExtension(m_str_GolPFileName, False) = "golp" Then
        SaveGolPFile m_str_GolPFileName, gol.World
    ElseIf fc_FileExtension(m_str_GolFFileName, False) = "golf" Then
        SaveGolFigurFile m_str_GolFFileName, gol.World
    End If

Fehler:
    'bla bla
End Sub

Private Sub mnu_LoadFigur_Click()
Dim tmpFigur As String
Dim strFigurFile As String
Dim strFigurName As String
Dim i As Integer
Dim tmpFigures() As String
On Error GoTo Fehler
        
    If Len(m_str_GolFFileName) = 0 Then
        m_str_GolFFileName = "*.golf"
    End If
    cdlgLoadSave.DialogTitle = "Datei Öffnen"
    cdlgLoadSave.Flags = cdlOFNOverwritePrompt
    cdlgLoadSave.InitDir = App.Path
    cdlgLoadSave.FileName = m_str_GolFFileName
    cdlgLoadSave.Filter = "Game of Life Figur *.golf|*.golf"
    cdlgLoadSave.CancelError = True
    cdlgLoadSave.ShowOpen   'Standarddialogfeld "Öffnen" anzeigen.
    strFigurFile = cdlgLoadSave.FileName
    
    If LoadGolFigurFile(strFigurFile, tmpFigur, strFigurName) Then
        pic_World.MousePointer = vbCrosshair
        ReDim tmpFigures(0 To UBound(g_str_Figurs, 1) + 1, 0 To 1)
        For i = 0 To UBound(g_str_Figurs, 1)
            tmpFigures(i, 0) = g_str_Figurs(i, 0)
            tmpFigures(i, 1) = g_str_Figurs(i, 1)
        Next i
        tmpFigures(i, 0) = strFigurName
        tmpFigures(i, 1) = tmpFigur
        g_str_Figurs = tmpFigures
        cmbFigurs.AddItem g_str_Figurs(i, 0)
        cmbFigurs.ItemData(cmbFigurs.ListCount - 1) = i
        cmbFigurs.ListIndex = cmbFigurs.ListCount - 1
        DrawFigurePreview
        m_bln_FigurListChanged = True
    End If
    
Fehler:
    'bla bla
End Sub

Private Sub mnu_LoadPopulation_Click()
Dim tmpWelt() As Byte
On Error GoTo Fehler
    If Len(m_str_GolPFileName) = 0 Then m_str_GolPFileName = "*.golp"
    cdlgLoadSave.DialogTitle = "Datei Öffnen"
    cdlgLoadSave.Flags = cdlOFNOverwritePrompt
    cdlgLoadSave.InitDir = App.Path
    cdlgLoadSave.FileName = m_str_GolPFileName
    cdlgLoadSave.Filter = "Game of Life Population *.golp|*.golp"
    cdlgLoadSave.CancelError = True
    cdlgLoadSave.ShowOpen   'Standarddialogfeld "Öffnen" anzeigen.
    m_str_GolPFileName = cdlgLoadSave.FileName
    
    If LoadGolPFile(m_str_GolPFileName, tmpWelt()) Then
        Me.MousePointer = vbHourglass
        gol.World = tmpWelt
        'GoL.WorldSize = UBound(tmpWelt, 1) + 1
        gol.Resize_World gol.CellSize, gol.WorldSize
        Call Form_Resize
        Me.MousePointer = vbDefault
    End If
    
Fehler:
    'bla bla
End Sub

'---------------------------------------------------------------------------------------------------Laden und Speichern

Private Sub LoadMenuIcons()
    
    With ctrl_Menupic
        .DestHeight = 14
        .DestWidth = 14
        .SourceLeft = 6
        .SourceWidth = 26
        .SourceTop = 6
        .SourceHeight = 26
        
        .Add_MenuIcon "0, 0", chk_Start.Picture             'mnuStartStop
        .Add_MenuIcon "0, 1", cmd_PlaySteps.Picture         'mnu_AutoStopPlay
        .Add_MenuIcon "0, 5", cmd_Clear.Picture             'mnu_Clear
        
        .Add_MenuIcon "2, 0", opt_Tools(0).Picture          '&Figurauswahl (mnu_Tools_Arr(0))
        .Add_MenuIcon "2, 1", opt_Tools(1).Picture          '&Radieren (mnu_Tools_Arr(1))
        .Add_MenuIcon "2, 2", opt_Tools(2).Picture          '&Zeichnen Pinsel (mnu_Tools_Arr(2))
        .Add_MenuIcon "2, 3", opt_Tools(3).Picture          'Invertieren &Pinsel (mnu_Tools_Arr(3))
        .Add_MenuIcon "2, 4", opt_Tools(4).Picture          '&Zelle Invertieren (mnu_Tools_Arr(4))
        .Add_MenuIcon "2, 5", opt_Tools(5).Picture          '&Auswahlbereich Aktivieren (mnu_Tools_Arr(5))
        .Add_MenuIcon "2, 6", cmd_Rotate.Picture            'Figur um 90°  nach rechts &Drehen (mnu_Tools_Rotate)
                
        picTmp.Picture = LoadResPicture(101, vbResBitmap)
        .Add_MenuIcon "0, 12", picTmp.Picture
        
        .DestHeight = 14
        .DestWidth = 14
        .SourceLeft = 1
        .SourceWidth = 24
        .SourceTop = 0
        .SourceHeight = 24
                
        
       
        .DestHeight = 15
        .DestWidth = 15
        .SourceLeft = 1
        .SourceWidth = 16
        .SourceTop = 1
        .SourceHeight = 16
        
        picTmp.Picture = LoadResPicture(108, vbResBitmap)       'Generationen -Zähler &Rücksetzen (mnu_CounterReset)
        .Add_MenuIcon "0, 3", picTmp.Picture
        
        picTmp.Picture = LoadResPicture(109, vbResBitmap)       '&zufällige Start-Population (mnu_Random)
        .Add_MenuIcon "0, 4", picTmp.Picture
        
        picTmp.Picture = LoadResPicture(102, vbResBitmap)
        .Add_MenuIcon "1, 0", picTmp.Picture
        picTmp.Picture = LoadResPicture(103, vbResBitmap)
        .Add_MenuIcon "1, 1", picTmp.Picture
        picTmp.Picture = LoadResPicture(104, vbResBitmap)
        .Add_MenuIcon "1, 2", picTmp.Picture
        
        picTmp.Picture = LoadResPicture(106, vbResBitmap)       'Figur &laden... (mnu_LoadFigur)
        .Add_MenuIcon "0, 7", picTmp.Picture
        picTmp.Picture = LoadResPicture(107, vbResBitmap)       'Population &laden... (mnu_LoadPopulation)
        .Add_MenuIcon "0, 9", picTmp.Picture
         picTmp.Picture = LoadResPicture(105, vbResBitmap)      'Population &speichern unter... (mnu_SavePopulation)
        .Add_MenuIcon "0, 10", picTmp.Picture
               
        picTmp.Picture = LoadResPicture(110, vbResBitmap)       'Die &Welt verändern... (mnu_WorldSize)
        .Add_MenuIcon "3, 2", picTmp.Picture
        picTmp.Picture = LoadResPicture(111, vbResBitmap)       '&Autostop Einstellung... (mnu_SetAutostop)
        .Add_MenuIcon "3, 3", picTmp.Picture
        picTmp.Picture = LoadResPicture(112, vbResBitmap)       'Generation -&Lebensdauer(Interval) ändern... (mnu_SetInterval)
        .Add_MenuIcon "3, 4", picTmp.Picture
                
        .DestHeight = 14
        .DestWidth = 14
        .SourceLeft = 6
        .SourceWidth = 26
        .SourceTop = 6
        .SourceHeight = 26
        .Add_MenuIcon "1, 4", cmd_Rotate.Picture
        DoEvents: Sleep 10



    End With
End Sub
