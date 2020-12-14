VERSION 5.00
Begin VB.Form FrmLifeRule 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "LifeRule"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "FrmLifeRule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnAnti 
      Caption         =   "Anti"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton BtnNot 
      Caption         =   "Not"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ListBox LstNewBorn 
      Height          =   1815
      ItemData        =   "FrmLifeRule.frx":000C
      Left            =   1440
      List            =   "FrmLifeRule.frx":002B
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox LstSurvive 
      Height          =   1815
      ItemData        =   "FrmLifeRule.frx":004A
      Left            =   120
      List            =   "FrmLifeRule.frx":0069
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "New Born"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Survive"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label LblLifeRule 
      Alignment       =   2  'Zentriert
      Caption         =   "/"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmLifeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_RetVal   As VbMsgBoxResult
Private m_LifeRule As LifeRule

Private Sub Form_Load()
    '
End Sub
Friend Function ShowDialog(aOwner As Form, ByRef aLifeRuleInOut As LifeRule) As VbMsgBoxResult
    m_LifeRule = aLifeRuleInOut
    Call MLifeRule.LifeRuleToListBox(m_LifeRule, Me.LstSurvive, Me.LstNewBorn)
    Me.LblLifeRule.Caption = MLifeRule.LifeRuleToString(m_LifeRule)
    Me.Show vbModal, aOwner
    ShowDialog = m_RetVal
    If ShowDialog = vbOK Then
        aLifeRuleInOut = m_LifeRule
    End If
End Function
Private Sub BtnOK_Click()
    m_RetVal = vbOK
    Unload Me
End Sub
Private Sub BtnCancel_Click()
    m_RetVal = vbCancel
    Unload Me
End Sub
Private Sub BtnNot_Click()
    Dim i As Long
    For i = 0 To 8
        Me.LstSurvive.Selected(i) = Not Me.LstSurvive.Selected(i)
        Me.LstNewBorn.Selected(i) = Not Me.LstNewBorn.Selected(i)
    Next
    Call NewLifeRule
End Sub
Private Sub BtnAnti_Click()
    Dim i As Long, n As Long: n = 8
    ReDim b(0 To 1, 0 To n) As Boolean
    For i = 0 To n
        b(0, i) = Me.LstSurvive.Selected(i)
        b(1, i) = Me.LstNewBorn.Selected(i)
    Next
    For i = 0 To n
        Me.LstSurvive.Selected(i) = Not b(1, n - i)
        Me.LstNewBorn.Selected(i) = Not b(0, n - i)
    Next
    Call NewLifeRule
End Sub

Private Sub NewLifeRule()
    Dim s As String
    s = s & GetStrFromListBox(Me.LstSurvive) & "/"
    s = s & GetStrFromListBox(Me.LstNewBorn)
    m_LifeRule = MLifeRule.New_LifeRule(s)
    Me.LblLifeRule.Caption = MLifeRule.LifeRuleToString(m_LifeRule)
End Sub
Private Function GetStrFromListBox(aLB As ListBox) As String
    Dim s As String
    Dim i As Long
    If aLB.SelCount > 0 Then
        For i = 0 To aLB.ListCount - 1
            If aLB.Selected(i) Then s = s & aLB.List(i)
        Next
    End If
    GetStrFromListBox = s
End Function

Private Sub LstSurvive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call NewLifeRule
End Sub
Private Sub LstNewBorn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call NewLifeRule
End Sub


