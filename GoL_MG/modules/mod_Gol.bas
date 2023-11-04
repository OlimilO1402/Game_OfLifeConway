Attribute VB_Name = "mod_Gol"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public g_byte_Figur() As Byte
Public g_str_Figurs() As String
Public g_str_IniFile As String

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Sub GetINISettings(ByRef gol As clsGameOfLife)
On Error Resume Next
Dim Value As String
 
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "Worldsize", Value) Then
        gol.WorldSize = IIf(Val(Value) > 0, CLng(Val(Value)), 100)
    Else
        gol.WorldSize = 100
    End If
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "Cellsize", Value) Then
        gol.CellSize = IIf(Val(Value) > 0, CLng(Val(Value)), 5)
    Else
        gol.CellSize = 5
    End If
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "Interval", Value) Then
        gol.Interval = IIf(Val(Value) > 0, Val(Value), 10)
    Else
        gol.Interval = 10
    End If
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "Steps2Play", Value) Then
        gol.Steps2Go = IIf(Val(Value) > 0, Val(Value), 10)
    Else
        gol.Steps2Go = 10
    End If
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "BorderType", Value) Then
        gol.BorderType = IIf(Val(Value) > 0, Val(Value), 0)
    Else
        gol.BorderType = 0
    End If
    
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "DrawGrid", Value) Then
        If (Val(Value) = True) Or (Value = "Wahr") Or (Value = "True") Then
            gol.DrawGrid = True
        Else
            gol.DrawGrid = False
        End If
    Else
        gol.DrawGrid = True
    End If
    
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "RoundedCells", Value) Then
        If (Val(Value) = True) Or (Value = "Wahr") Or (Value = "True") Then
            gol.CellsRounded = True
        Else
            gol.CellsRounded = False
        End If
    Else
        gol.CellsRounded = True
    End If
        
    If Mod_INI.fcGetINIValue(g_str_IniFile, "GoL Settings", "RulesDefinition", Value) Then
        If (InStr(1, Value, "/") > 0) Then
            frm_GoL.gol.RulesDefinition = Value
        Else
            frm_GoL.gol.RulesDefinition = "23/3" 'Conways 'Game of Life' Regelwerk
        End If
    Else
        frm_GoL.gol.RulesDefinition = "23/3" 'Conways 'Game of Life' Regelwerk
    End If
End Sub

Public Sub SaveIniSettings(ByRef gol As clsGameOfLife)
On Error Resume Next
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "Worldsize", gol.WorldSize
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "Cellsize", gol.CellSize
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "Interval", gol.Interval
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "Steps2Play", gol.Steps2Go
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "BorderType", gol.BorderType
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "DrawGrid", Int(gol.DrawGrid)
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "RoundedCells", Int(gol.CellsRounded)
    Mod_INI.fcSetINIValue g_str_IniFile, "GoL Settings", "RulesDefinition", frm_GoL.gol.RulesDefinition
End Sub
