Attribute VB_Name = "MGOLFile"
Option Explicit
Public Type TGOLFile
    FileName As String
End Type
'was braucht man alles?
' man muß nur den Dateinamen speichern werden
' in einem Feld, bzw in einer Generation gibt es nur zewei verschiedene Zustände
' Life oder Field (Berg oder Tal zum exportieren/importieren)
'
'
' das ist doch eigentlich auch ein Schmarrn
' das gehört zum UI und damit ins Formular
'

Public Function SaveGeneration(this As TGOLFile, aGOL As TGameOfLife)
    Dim FNr As Integer: FNr = FreeFile
TryE: On Error GoTo CatchE
    With this
        Open .FileName For Output As FNr
            Print #FNr, MGeneration.GenerationToString(aGOL.ThisG)
        Close #FNr
        SaveGeneration = True
    End With
CatchE:
    If Err <> 0 Then MsgBox Err.Description
FinallyE:
    Close #FNr
End Function
Public Function LoadGeneration(this As TGOLFile, aGOL As TGameOfLife) As Boolean
    Dim FNr As Integer: FNr = FreeFile
TryE: On Error GoTo CatchE
    With this
        Open .FileName For Binary As FNr
            Dim s As String: s = String$(LOF(FNr), vbNullChar)
            Get #FNr, , s
        Close #FNr
        'so jetzt bis zum nächsten vbcrlf zählen
        Dim nx As Long: nx = InStr(1, s, vbCrLf) - 1
        Dim ny As Long: ny = Len(s)
        If nx > 0 Then ny = ny / (nx + 2)
        'With aGOL
        'so jetzt das Feld dimensionieren
        'überprüfen ob die Größe übereinsttimmt,
        If Not ((aGOL.SizeX = nx) And (aGOL.SizeY = ny)) Then
            'wenn nicht dann neu anlegen
            Call New_GameOfLife(aGOL, nx, ny, aGOL.LifeRule)
        End If
        'dann einlesen
        Call MGeneration.ParseGeneration(aGOL.ThisG, s)
        LoadGeneration = True
    End With
CatchE:
    If Err <> 0 Then MsgBox Err.Description
FinallyE:
    Close #FNr

End Function

