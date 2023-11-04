Attribute VB_Name = "mod_GOL_Files"
'---------------------------------------------------------------------------------------
' Modul       : GoL_Files
' Datum/Zeit  : 05.08.2008 10:04
' Autor       : Marco Groﬂert
' Zweck       : game of life       (Spaﬂprojekt initiiert durch einen gleichnamigen
'                                   Forumthread im www.ActiveVB.de VB5/VB6 Forum)
'             Laden und Speichern von Game of Life Dateien
'---------------------------------------------------------------------------------------

Option Explicit

Public Function LoadFigursList(Optional ByVal FullPath As String = vbNullString) As Integer
    Dim fN As Double
    Dim tmpLine As String
    Dim i As Integer
    Dim figCount As Integer
    Dim tmpFigurs() As String
    Dim tmpFigNames() As String
    fN = FreeFile
    If FullPath = vbNullString Then FullPath = fc_AppPath & "figurs.golfl"
    On Error GoTo ErrNoFigurs
    If fc_FileExist(FullPath) Then
        Open FullPath For Input As fN#
            Do While Not EOF(fN)
                Line Input #fN, tmpLine
                ReDim Preserve tmpFigNames(figCount)
                ReDim Preserve tmpFigurs(figCount)
                tmpFigNames(figCount) = Split(tmpLine, ":")(0)
                tmpFigurs(figCount) = Split(tmpLine, ":")(1)
                figCount = figCount + 1
            Loop
        Close #fN
    End If

ErrNoFigurs:
    Err.Clear
    On Error Resume Next
    
    ReDim g_str_Figurs(0 To figCount - 1, 0 To 1)

    For i = 0 To figCount - 1
        g_str_Figurs(i, 0) = tmpFigNames(i)
        tmpFigurs(i) = Replace(tmpFigurs(i), "0-", "0|")
        tmpFigurs(i) = Replace(tmpFigurs(i), "1-", "1|")
        tmpFigurs(i) = Replace(tmpFigurs(i), "|,", ",")
        If InStr(1, tmpFigurs(i), "|") = 0 Then
            tmpFigurs(i) = Replace(tmpFigurs(i), "0", "0|")
            tmpFigurs(i) = Replace(tmpFigurs(i), "1", "1|")
            tmpFigurs(i) = Replace(tmpFigurs(i), "|,", ",")
        End If
        g_str_Figurs(i, 1) = tmpFigurs(i)
    Next i

    LoadFigursList = figCount
End Function

Public Function SaveFigursList(ByRef FigLst() As String, Optional ByVal FullPath As String = vbNullString) As Boolean
    Dim fN As Double
    Dim tmpLine As String
    Dim BackupFile As String
    Dim i As Integer
    On Error GoTo ErrHandler
    
    If FullPath = vbNullString Then FullPath = fc_AppPath & "figurs.golfl"
    BackupFile = fc_Path(FullPath) & fc_FileTitel(fc_FileName(FullPath)) & ".golflbk"
    If fc_FileExist(FullPath) Then
        If fc_FileExist(BackupFile) Then fc_DelFile BackupFile
        fc_FileCopy FullPath, BackupFile, True, False
    End If
    
    fN = FreeFile
    Open FullPath For Output As fN#
        For i = 0 To UBound(FigLst, 1)
            If FigLst(i, 1) <> "delete" Then
                tmpLine = Replace(FigLst(i, 0) & ":" & FigLst(i, 1), "|", "")
                Print #fN, tmpLine
            End If
        Next i
    Close #fN
    SaveFigursList = True
    Exit Function
ErrHandler:
    SaveFigursList = False
    Close #fN
End Function

Public Function SaveGolFigurFile(ByVal FullPath As String, World() As Byte, Optional ByVal Figurname As String) As Boolean
    Dim fN As Double
    Dim tmpLine As String
    Dim X As Integer
    Dim Y As Integer
    Dim StartCol As Integer
    Dim EndCol As Integer
    Dim StartRow As Integer
    Dim EndRow As Integer
    On Error GoTo ErrHandler
    If Len(Figurname) = 0 Then Figurname = fc_FileTitel(fc_FileName(FullPath))
    StartCol = UBound(World, 1) - 1
    StartRow = UBound(World, 1) - 1
    EndCol = 0
    EndRow = 0
    ' Erste Spalte mit Lebender Zelle suchen
    For Y = 0 To UBound(World, 1)
        For X = 0 To UBound(World, 1) - 1
            If World(X, Y) = 1 Then
                If X < StartRow Then StartRow = X
                Exit For
            End If
        Next X
        For X = UBound(World, 1) - 1 To 0 Step -1
            If World(X, Y) = 1 Then
                If X > EndRow Then EndRow = X
                Exit For
            End If
        Next X
    Next Y
    
    For X = 0 To UBound(World, 1)
        For Y = 0 To UBound(World, 1) - 1
            If World(X, Y) = 1 Then
                If Y < StartCol Then StartCol = Y
                Exit For
            End If
        Next Y
        For Y = UBound(World, 1) - 1 To 0 Step -1
            If World(X, Y) = 1 Then
                If Y > EndCol Then EndCol = Y
                Exit For
            End If
        Next Y
    Next X
    
    fN = FreeFile
    Open FullPath For Output As fN#
        tmpLine = Figurname & ":"
        For Y = StartCol To EndCol
            For X = StartRow To EndRow
                tmpLine = tmpLine & World(X, Y)
            Next X
            tmpLine = tmpLine & ","
        Next Y
        tmpLine = Mid(tmpLine, 1, Len(tmpLine) - 1)
        Print #fN, tmpLine
    Close #fN
    SaveGolFigurFile = True
    Exit Function
ErrHandler:
    Close #fN
    SaveGolFigurFile = False
End Function

Public Function SaveGolPFile(ByVal FullPath As String, World() As Byte) As Boolean
    Dim fN As Double
    Dim tmpLine As String
    Dim X As Integer
    Dim Y As Integer

    On Error GoTo ErrHandler
'    If fc_FileExist(FullPath) Then
'        If MsgBox("Die Datei: """ & fc_FileName(FullPath) & """ existiert bereits am Angegebenen Ort! " & vbCrLf & _
'                    "Soll die vorhandene Datei ersetzt werden?", vbQuestion Or vbYesNo) = vbNo Then Exit Function
'    End If

    fN = FreeFile
    Open FullPath For Output As fN#
        For Y = 0 To UBound(World, 1) - 1
            tmpLine = vbNullString
            For X = 0 To UBound(World, 1) - 1
                tmpLine = tmpLine & World(X, Y)
            Next X
            Print #fN, tmpLine
        Next Y
    Close #fN
    SaveGolPFile = True
    Exit Function
ErrHandler:
    SaveGolPFile = False
    Close #fN
End Function

Public Function LoadGolFigurFile(ByVal FullPath As String, Figur As String, ByRef Figurname As String) As Boolean
    Dim fN As Double
    Dim tmpLine As String

    On Error GoTo ErrHandler
    If Not fc_FileExist(FullPath) Then GoTo ErrHandler

    fN = FreeFile
    Open FullPath For Input As fN#
        If Not EOF(fN) Then
            Line Input #fN, tmpLine
        End If
        Figurname = Split(tmpLine, ":")(0)
        Figur = Split(tmpLine, ":")(1)
        Figur = Replace(Figur, "0", "0|")
        Figur = Replace(Figur, "1", "1|")
        Figur = Replace(Figur, "|,", ",")
    Close #fN
    
    LoadGolFigurFile = True
    Exit Function
ErrHandler:
    Close #fN
    LoadGolFigurFile = False
End Function

Public Function LoadGolPFile(ByVal FullPath As String, World() As Byte) As Boolean
    Dim fN As Double
    Dim tmpLines() As String
    Dim X As Integer
    Dim Y As Integer

    On Error GoTo ErrHandler
    If Not fc_FileExist(FullPath) Then GoTo ErrHandler
    
    fN = FreeFile
    Open FullPath For Input As fN#
        Y = -1
        Do While Not EOF(fN)
            Y = Y + 1
            ReDim Preserve tmpLines(Y)
            Line Input #fN, tmpLines(Y)
        Loop
        ReDim World(Len(tmpLines(0)) - 1, 0 To Y)
        For Y = 0 To UBound(World, 1)
            For X = 0 To UBound(World, 1)
               World(X, Y) = Val(Mid(tmpLines(Y), X + 1, 1))
            Next X
        Next Y
    Close #fN
    
    LoadGolPFile = True
    Exit Function
ErrHandler:
    Close #fN
    LoadGolPFile = False
End Function
