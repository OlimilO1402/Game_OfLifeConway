Attribute VB_Name = "MLifeRule"
Option Explicit

Public Enum ERule
    RNB0 = &H1   ' = 2 ^ 0 (Bit 1 für 0)
    RNB1 = &H2   ' = 2 ^ 1 (Bit 2 für 1)
    RNB2 = &H4   ' = 2 ^ 2
    RNB3 = &H8   ' = 2 ^ 3
    RNB4 = &H10  ' = 2 ^ 4
    RNB5 = &H20  ' = 2 ^ 5
    RNB6 = &H40  ' = 2 ^ 6
    RNB7 = &H80  ' = 2 ^ 7
    RNB8 = &H100 ' = 2 ^ 8
End Enum

Public Type LifeRule
    RuleSurvive As ERule
    RuleNewBorn As ERule
End Type

'
'Rule 23/3 heißt:
' * bei RuleSurvive ist Bit 2(+1) und Bit 3(+1) gesetzt
' * bei RuleNewBorn ist Bit 3(+1) gesetzt
'

Public Function New_LifeRule(ByVal sLifeRule As String) As LifeRule
    'übernimmt einen String
    'die beiden Regeln für Überleben und Neugeboren
    'werden aus dem String geparst.
    
    Dim s() As String
    s = Split(sLifeRule, "/")
    With New_LifeRule
        .RuleSurvive = ParseRule(s(0))
        If UBound(s) > 0 Then
            .RuleNewBorn = ParseRule(s(1))
        End If
    End With
End Function

Private Function ParseRule(sRule As String) As ERule
'eine Regel aus einem String lesen
'es werden nur Ziffern zwischen 0 und 8 akzeptiert
'aber eigetnlich ist eine Überprüfung auf Ziffern nicht notwendig, weil ohnehin
'keine anderen Zeichen möglich sind, durch das Formular FrmLifeRule wird dies ja
'effizient verhindert.
    Dim i As Long, nNeighBours As Long
    Dim c As Integer 'ein Character
    For i = 1 To Len(sRule)
        c = Asc(Mid$(sRule, i, 1))
        Select Case c
        '   "0", "8"
        Case 48 To 56
            nNeighBours = c - 48
            ParseRule = ParseRule Or CERule(nNeighBours)
        End Select
    Next
End Function

Private Function CERule(ByVal nNeighBours As Long) As ERule
'eine Anzahl Nachbarn als Long in die zugehörige Enum-Konstante
'verwandeln diese Funktion sollte in der run-Funktion geinlined werden
    Select Case nNeighBours
    Case 0: CERule = RNB0
    Case 1: CERule = RNB1
    Case 2: CERule = RNB2
    Case 3: CERule = RNB3
    Case 4: CERule = RNB4
    Case 5: CERule = RNB5
    Case 6: CERule = RNB6
    Case 7: CERule = RNB7
    Case 8: CERule = RNB8
    End Select
End Function

Public Function LifeRuleToString(this As LifeRule) As String
    'zur Anzeige wird die Regel in einen String verwandelt
    With this
        LifeRuleToString = RuleToString(.RuleSurvive) & "/" & _
                           RuleToString(.RuleNewBorn)
    End With
End Function
Public Function RuleToString(ByVal this As ERule) As String
    'eine einzelne Regel in einen String verwandeln
    Dim s As String
    Dim i As Long
    For i = 0 To 8
        If this And 2 ^ i Then s = s & CStr(i)
    Next
    RuleToString = s
End Function
Public Sub LifeRuleToListBox(this As LifeRule, aLBSurvive As ListBox, aLBNewBorn As ListBox)
    With this
        Call RuleToListBox(.RuleSurvive, aLBSurvive)
        Call RuleToListBox(.RuleNewBorn, aLBNewBorn)
    End With
End Sub
Private Sub RuleToListBox(ByVal this As ERule, aLB As ListBox)
    Dim i As Long
    For i = 0 To aLB.ListCount - 1
        aLB.Selected(i) = (this And (2 ^ i))
    Next
End Sub


