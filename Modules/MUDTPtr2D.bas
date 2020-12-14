Attribute VB_Name = "MUDTPtr2D"
Option Explicit '2008_07_31 Zeilen:  99
' Ein SafeArray-Descriptor dient in VB als ein universaler Zeiger
Public Type TUDTPtr2D
    pSA         As Long
    Reserved    As Long ' z.B. für IRecordInfo
    cDims       As Integer
    fFeatures   As Integer
    cbElements  As Long
    cLocks      As Long
    pvData      As Long
    cElements2  As Long
    lLBound2    As Long
    cElements1  As Long
    lLBound1    As Long
End Type

Public Enum SAFeature
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4

    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20
    FADF_HAVEIID = &H40
    FADF_HAVEVARTYPE = &H80

    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800
    FADF_RESERVED = &HF008
End Enum

Public Declare Sub GetMem4 Lib "msvbvm60" ( _
    ByRef pSrc As Any, ByRef pDst As Any)
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" ( _
    ByRef pArr() As Any) As Long

Public Sub New_UDTPtr2D(ByRef this As TUDTPtr2D, _
                        ByVal Feature As SAFeature, _
                        ByVal bytesPerElement As Long, _
                        Optional ByVal CountElements1 As Long = 1, _
                        Optional ByVal lLBound1 As Long = 0, _
                        Optional ByVal CountElements2 As Long = 1, _
                        Optional ByVal lLBound2 As Long = 0)
'kann nur eine Sub sein darf keine Function sein
    With this
        .pSA = VarPtr(.cDims)
        .cDims = 2
        .cbElements = bytesPerElement
        .fFeatures = CInt(Feature)
        .cElements2 = CountElements2
        .lLBound2 = lLBound2
        .cElements1 = CountElements1
        .lLBound1 = lLBound1
    End With
    'Debug.Print UDTPtr2DToString(this)
End Sub

' Um zu überprüfen ob der UDTPtr auch das enthält was er soll
' kann man folgende Funktion verwenden
Public Function UDTPtr2DToString(this As TUDTPtr2D) As String
    Dim s As String
    With this
        s = s & "pSA         : " & CStr(.pSA) & vbCrLf
        s = s & "Reserved    : " & CStr(.Reserved) & vbCrLf
        s = s & "cDims       : " & CStr(.cDims) & vbCrLf
        s = s & "fFeatures   : " & FeaturesToString(CLng(.fFeatures)) & vbCrLf
        s = s & "cbElements  : " & CStr(.cbElements) & vbCrLf
        s = s & "cLocks      : " & CStr(.cLocks) & vbCrLf
        s = s & "pvData      : " & CStr(.pvData) & vbCrLf
        s = s & "cElements1  : " & CStr(.cElements1) & vbCrLf
        s = s & "lLBound1    : " & CStr(.lLBound1) & vbCrLf
        s = s & "cElements2  : " & CStr(.cElements2) & vbCrLf
        s = s & "lLBound2    : " & CStr(.lLBound2) & vbCrLf
    End With
    UDTPtr2DToString = s
End Function

Private Function FeaturesToString(ByVal Feature As SAFeature) As String
    Const o As String = " Or "
    Dim s As String
    Dim f As SAFeature
    f = Feature
    s = s & IIf((f And FADF_AUTO), IIf(Len(s), o, "") & "FADF_AUTO", "")
    s = s & IIf((f And FADF_STATIC), IIf(Len(s), o, "") & "FADF_STATIC", "")
    s = s & IIf((f And FADF_EMBEDDED), IIf(Len(s), o, "") & "FADF_EMBEDDED", "")
    s = s & IIf((f And FADF_FIXEDSIZE), IIf(Len(s), o, "") & "FADF_FIXEDSIZE", "")
    s = s & IIf((f And FADF_RECORD), IIf(Len(s), o, "") & "FADF_RECORD", "")
    s = s & IIf((f And FADF_HAVEVARTYPE), IIf(Len(s), o, "") & "FADF_HAVEVARTYPE", "")
    s = s & IIf((f And FADF_BSTR), IIf(Len(s), o, "") & "FADF_BSTR", "")
    s = s & IIf((f And FADF_UNKNOWN), IIf(Len(s), o, "") & "FADF_UNKNOWN", "")
    s = s & IIf((f And FADF_DISPATCH), IIf(Len(s), o, "") & "FADF_DISPATCH", "")
    s = s & IIf((f And FADF_VARIANT), IIf(Len(s), o, "") & "FADF_VARIANT", "")
    s = s & IIf((f And FADF_RESERVED), IIf(Len(s), o, "") & "FADF_RESERVED", "")
    s = s & IIf((f And FADF_UNKNOWN), IIf(Len(s), o, "") & "FADF_UNKNOWN", "")
    FeaturesToString = s
End Function
