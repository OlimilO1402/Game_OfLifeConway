Attribute VB_Name = "Mod_INI"
' Modul für das Lesen und Schreiben von INI Dateien
' ©2001 by Marco Großert
' marco@grossert.com
Option Explicit
Option Base 0

' API Deklarationen--------------------------------------------------------------

Private Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long

'Private Declare Function WritePrivateProfileSection Lib "kernel32" _
'        Alias "WritePrivateProfileSectionA" _
'        (ByVal lpAppName As String, _
'        ByVal lpString As String, _
'        ByVal lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSection Lib "kernel32" _
        Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Private result As Long

' Public-Variablen-Deklarationen------------------------------------------------

' Private-Variablen-Deklarationen------------------------------------------------

' Funktionen--------------------------------------------------------------------

Public Function fcDeleteINIKey(ByVal Path As String, ByVal Sect As String, ByVal Key As String) As Boolean
    ' löscht INI-Key Einträge mittels API
    On Error GoTo Error
    result = WritePrivateProfileString(Sect, Key, 0&, Path)
    fcDeleteINIKey = True
    Exit Function
Error:
    fcDeleteINIKey = False
End Function
 
Public Function fcDeleteINISection(ByVal Path As String, ByVal Sect As String)
    ' löscht INI-Section Einträge mittels API
    On Error GoTo Error
    result = WritePrivateProfileString(Sect, 0&, 0&, Path)
    fcDeleteINISection = True
    Exit Function
Error:
    fcDeleteINISection = False
End Function

Public Function fcSetINIValue(ByVal Path As String, ByVal Sect As String, ByVal Key As String, ByVal Value As String) As Boolean
    ' Schreibt INI-Dateien mittels API
    On Error GoTo Error
    Dim result&
    'Wert schreiben
    result = WritePrivateProfileString(Sect, Key, Value, Path)
    fcSetINIValue = True
    Exit Function
Error:
    fcSetINIValue = False
End Function

Public Function fcGetINIValue(ByVal Path As String, ByVal Sect As String, ByVal Key As String, ByRef Value As String) As Boolean
    ' Liest INI-Dateien mittels API
    On Error GoTo Error
    Dim result As Long, Buffer As String
    'Wert lesen
    Buffer = Space$(128)
    result = GetPrivateProfileString(Sect, Key, vbNullString, Buffer, Len(Buffer), Path)
    Value = Left$(Buffer, result)
    Value = Trim(Value)
    If result <> 0 Then fcGetINIValue = True
    Exit Function
Error:
    fcGetINIValue = False
End Function

Public Function fcGetINIArray(ByVal Path As String, ByVal Sect As String, ByRef xArray() As String) As Boolean
On Error GoTo ErrOut
    fcGetINIArray = True
    ' Liest INI-Section mittels API in xArray()
    Dim Buffer As String
    Dim l, p, Z As Integer
    Buffer = Space(32767)
    result = GetPrivateProfileSection(Sect, Buffer, Len(Buffer), Path)
    Buffer = Left$(Buffer, result)
    If Buffer <> "" Then
        l = 1
        Z = 0
        Do While l < result
            ReDim Preserve xArray(Z)
            p = InStr(l, Buffer, Chr$(0))
            If p = 0 Then Exit Do
            xArray(Z) = Trim(Mid$(Buffer, l, p - l))
            Z = Z + 1
            l = p + 1
        Loop
    End If
    If result <> 0 And xArray(0) <> "" Then fcGetINIArray = True
Exit Function
ErrOut:
    fcGetINIArray = False
End Function
