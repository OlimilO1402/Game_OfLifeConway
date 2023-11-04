Attribute VB_Name = "mod_File_and_Path"
Option Explicit

'############################-Deklarationen-###############################################################
'Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Datei OP's (zB. in Papierkorb)--##########################################################################
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                                (lpFileOp As SHFILEOPSTRUCT) As Long

Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Public Const FO_COPY = &H2                 ' Kopiert das File in pFROM nach pTo
Public Const FO_DELETE = &H3               ' Löscht das File in pFrom (pTo wird ignoriert)
Public Const FO_MOVE = &H1                 ' Verschiebt das File in pFROM nach pTo
Public Const FO_RENAME = &H4               ' Umbenennen des Files in pTo
' KONSTANTEN DER FLAGS

Public Const FOF_ALLOWUNDO = &H40           ' Undo Information -> Schiebt beim Löschen das (die) File(s)
                                            ' in den Papierkorb
Public Const FOF_NOERRORUI = &H400
Public Const FOF_CONFIRMMOUSE = &H2        ' Bislang keine bekannte Funktion
Public Const FOF_CREATEPROGRESSDLG = &H0   ' Handle zum Eltern-Fenster der Progress-Dialogbox (also Me.hwnd)
Public Const FOF_FILESONLY = &H80          ' Nur Files - KEINE ORDNER - wenn *.* als Source

Public Const FOF_MULTIDESTFILES = &H1      ' Für diverse Stellen bei DEST (der "pTo" muss dann die
                                            ' gleiche Anzahl von Zielen aufweisen wie "pFrom"
Public Const FOF_NOCONFIRMATION = &H10     ' ANTWORTET AUTOMATISCH MIT 'JA für alle'
Public Const FOF_NOCONFIRMMKDIR = &H200    ' Keine Abfrage für einen neuen Ordner, falls benötigt
Public Const FOF_RENAMEONCOLLISION = &H8   ' Bei Namenskollisionen im ZIEL wird ein neuer Name
                                            ' erzeugt (z.B. Kopie(2) von xy.tmp)
Public Const FOF_SILENT = &H4              ' Zeigt keine Fortschritts-Dialogbox (fliegende Blätter)
Public Const FOF_SIMPLEPROGRESS = &H100    ' Zeigt die Fortschritts-Dialogbox an, aber ohne Filenamen


Public Const FOF_WANTMAPPINGHANDLE = &H20   ' Wenn FOF_RENAMECOLLISION gewählt wird,
                                            ' hNameMappings wird gefüllt (Anzahl)

'##########################################################################--Datei OP's (zB. in Papierkorb)

'Systempfade ermitteln--##################################################################################
' Ordner-Auflistung
Public Enum SpecialFolderIDs
    myidNone = -1
    myidAppPath = -2
    CSIDL_ADMINTOOLS = &H30               ' \Start Menu\Programme\Verwaltung
'    CSIDL_ALTSTARTUP = &H1D               ' non localizedstartup
    CSIDL_APPDATA = &H1A                  ' <BenutzerName>\Anwendungsdaten
    CSIDL_BITBUCKET = &HA                 ' \Recycle Bin
    CSIDL_CDBURN_AREA = &H3B              'NEU
    CSIDL_COMMON_ADMINTOOLS = &H2F        ' All Users\Start Menu\Programme\Verwaltung
'    CSIDL_COMMON_ALTSTARTUP = &H1E        ' non localizedCommon startup
    CSIDL_COMMON_APPDATA = &H23           ' All Users\Anwendungsdaten
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19  ' All Users\Desktop
    CSIDL_COMMON_DOCUMENTS = &H2E         ' All Users\Dokumente
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_MUSIC = &H35             'NEU
'    CSIDL_COMMON_OEM_LINKS = &H3A         'NEU
    CSIDL_COMMON_PictureS = &H36          'NEU
    CSIDL_COMMON_PROGRAMS = &H17          ' All Users\Programme
    CSIDL_COMMON_STARTMENU = &H16         ' All Users\StartMenu
    CSIDL_COMMON_STARTUP = &H18           ' All Users\Startmenü\Programme\Autostart
    CSIDL_COMMON_TEMPLATES = &H2D         ' All Users\Templates
    CSIDL_COMMON_VIDEO = &H37             'NEU
    CSIDL_COMPUTERSNEARME = &H3D          'NEU
    CSIDL_CONNECTIONS = &H31              'NEU
'    CSIDL_CONTROLS = &H3                  ' My Computer\Control Panel
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0                   ' Desktop
    CSIDL_DESKTOPDIRECTORY = &H10         ' name>\Desktop
'    CSIDL_DRIVERS = &H11                  ' My Computer
    CSIDL_FAVORITES = &H6                 ' \Favorites
    CSIDL_FONTS = &H14                    ' windows\fonts
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1                  ' Internet Explorer (icon on desktop)
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C            ' name>\Local Settings\Applicaiton Data (non roaming)
'    CSIDL_MYDOCUMENTS = &HC                               'NEU
    CSIDL_MYMUSIC = &HD                               'NEU
    CSIDL_MYPictureS = &H27               ' C:\Program Files\My Pictures
    CSIDL_MYVIDEO = &HE                               'NEU
    CSIDL_NETHOOD = &H13                  ' \nethood
    CSIDL_NETWORK = &H12                  ' Network Neighborhood
    CSIDL_PERSONAL = &H5                  ' My Dokumente
    CSIDL_PRINTERS = &H4                  ' My Computer\Printers
    CSIDL_PRINTHOOD = &H1B                ' name>\PrintHood
    CSIDL_PROFILE = &H28                  ' USERPROFILE
    CSIDL_PROGRAM_FILES = &H26            ' C:\Program Files
    CSIDL_PROGRAM_FILES_COMMON = &H2B     ' C:\Program Files\Common
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C  ' x86 Program Files\Common on RISC
    CSIDL_PROGRAM_FILESX86 = &H2A         ' x86 C:\Program Files on RISC
    CSIDL_PROGRAMS = &H2                  ' Start Menu\Programme
    CSIDL_RECENT = &H8                    ' \Recent
    CSIDL_RESOURCES = &H38                'NEU
'    CSIDL_RESOURCES_LOCALIZED = &H39      'NEU
    CSIDL_SENDTO = &H9                    ' \SendTo
    CSIDL_STARTMENU = &HB                 ' \StartMenu
    CSIDL_STARTUP = &H7                   ' StartMenu\Programme\Startup
    CSIDL_SYSTEM = &H25                   ' GetSystemDirectory()
    CSIDL_SYSTEMX86 = &H29                ' x86 system directory on RISC
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24                  ' GetWindowsDirectory()
End Enum

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

'Private Const MAX_PATH = 260
Private Enum eULFlAGS
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_RETURNFSANCESTORS = &H8
    BIF_RETURNONLYFSDIRS = &H1
    BIF_STATUSTEXT = &H4
End Enum

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As eULFlAGS
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                            ByVal lpString2 As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, _
                                                            ByVal lpbuffer As String) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BROWSEINFO) As Long
                                
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" _
                                                            (ByVal hWndOwner As Long, _
                                                            ByVal nFolder As Long, _
                                                            pidl As ITEMIDLIST) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
                                                            (ByVal nBufferLength As Long, _
                                                            ByVal lpbuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
                                                            ByVal lpbuffer As String, _
                                                            ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                                                            (ByVal lpbuffer As String, _
                                                            ByVal nSize As Long) As Long
                                    
                                    
Private Const RETURNONLYFSDIRS = &H3
Private Const WM_SETTEXT = &HC
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                                                            (ByVal hwnd As Long, _
                                                            ByVal wMsg As Long, _
                                                            ByVal wParam As Long, _
                                                            ByVal lParam As String) As Long
                        
'##################################################################################--Systempfade ermitteln

'Datei-Extension Registrieren--###########################################################################
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
                                                            (ByVal hKey As Long, _
                                                            ByVal lpSubKey As String, _
                                                            phkResult As Long) As Long

Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
                                                            (ByVal hKey As Long, _
                                                            ByVal lpSubKey As String, _
                                                            ByVal dwType As Long, _
                                                            ByVal lpData As String, _
                                                            ByVal cbData As Long) As Long

Const HKEY_CLASSES_ROOT = &H80000000
Const REG_SZ As Long = 1
'###########################################################################--Datei-Extension Registrieren

'Datei-Versions-Info--####################################################################################
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
                                                            (ByVal lptstrFilename As String, _
                                                            ByVal dwhandle As Long, _
                                                            ByVal dwlen As Long, _
                                                            lpData As Any) As Long

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
                                                            (ByVal lptstrFilename As String, _
                                                            lpdwHandle As Long) As Long

Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
                                                            (pBlock As Any, _
                                                            ByVal lpSubBlock As String, _
                                                            lplpBuffer As Any, _
                                                            puLen As Long) As Long

'Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
'                                                            ByVal lpBuffer As String, _
'                                                            ByVal nSize As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                                            (Dest As Any, _
                                                            ByVal Source As Long, _
                                                            ByVal length As Long)

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer          'dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer         'dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer         'dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer      'dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer      'dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long             ' = 0x3F for version "0.42"
    dwFileFlags As VSFileFlags        ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As VSOSFlags             ' e.g. VOS_DOS_WINDOWS16
    dwFileType As VSFileTypeFlags    ' e.g. VFT_DRIVER
    dwFileSubtype As VSFileSubTypeFlags               ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long                ' e.g. 0
    dwFileDateLS As Long                ' e.g. 0
End Type

Private Enum VSFileInfoFlags
    VS_FFI_SIGNATURE = &HFEEF04BD
    VS_FFI_STRUCVERSION = &H10000
    VS_FFI_FILEFLAGSMASK = &H3F&
End Enum
Private Enum VSFileFlags
    VS_FF_DEBUG = &H1
    VS_FF_PRERELEASE = &H2
    VS_FF_PATCHED = &H4
    VS_FF_PRIVATEBUILD = &H8
    VS_FF_INFOINFERRED = &H10
    VS_FF_SPECIALBUILD = &H20
End Enum
Private Enum VSOSFlags
    VOS_UNKNOWN = &H0
    VOS_DOS = &H10000
    VOS_OS216 = &H20000
    VOS_OS232 = &H30000
    VOS_NT = &H40000
    VOS__BASE = &H0
    VOS__WINDOWS16 = &H1
    VOS__PM16 = &H2
    VOS__PM32 = &H3
    VOS__WINDOWS32 = &H4
    VOS_DOS_WINDOWS16 = &H10001
    VOS_DOS_WINDOWS32 = &H10004
    VOS_OS216_PM16 = &H20002
    VOS_OS232_PM32 = &H30003
    VOS_NT_WINDOWS32 = &H40004
End Enum
Private Enum VSFileTypeFlags
    VFT_UNKNOWN = &H0
    VFT_APP = &H1
    VFT_DLL = &H2
    VFT_DRV = &H3
    VFT_FONT = &H4
    VFT_VXD = &H5
    VFT_STATIC_LIB = &H7
End Enum
Private Enum VSFileSubTypeFlags
    VFT2_UNKNOWN = &H0
    VFT2_DRV_PRINTER = &H1
    VFT2_DRV_KEYBOARD = &H2
    VFT2_DRV_LANGUAGE = &H3
    VFT2_DRV_DISPLAY = &H4
    VFT2_DRV_MOUSE = &H5
    VFT2_DRV_NETWORK = &H6
    VFT2_DRV_SYSTEM = &H7
    VFT2_DRV_INSTALLABLE = &H8
    VFT2_DRV_SOUND = &H9
    VFT2_DRV_COMM = &HA
    VFT2_FONT_RASTER = &H1
    VFT2_FONT_VECTOR = &H2
    VFT2_FONT_TRUETYPE = &H3
End Enum

Public Type FILEVERSIONINFO
    StructVer As String
    FileVer As String
    ProductVer As String
    VerFlags As String
    OS As String
    typ As String
    SubTyp As String
End Type
'####################################################################################--Datei-Versions-Info

'Dateiinformationen--#####################################################################################
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, _
                                                                    riid As CLSIdType, _
                                                                    ByVal fown As Long, _
                                                                    lpUnk As Object) _
                                                                    As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
                                                                    (ByVal pszPath As String, _
                                                                    ByVal dwFileAttributes As Long, _
                                                                    psfi As SHFILEINFO, _
                                                                    ByVal cbFileInfo As Long, _
                                                                    ByVal uFlags As SHGetFileInfo_Flags) _
                                                                    As Long
Private Type IconType
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Public Enum IconSize
    Large = &H100
    Small = &H101
End Enum

Private Type CLSIdType
    id(16) As Byte
End Type
'Private Const MAX_PATH = 260
Private Enum SHGetFileInfo_Flags
    SHGFI_LARGEICON = &H0
    SHGFI_SMALLICON = &H1
    SHGFI_OPENICON = &H2
    SHGFI_SHELLICONSIZE = &H4
    SHGFI_PIDL = &H8
    SHGFI_USEFILEATTRIBUTES = &H10
    SHGFI_ADDOVERLAYS = &H20
    SHGFI_OVERLAYINDEX = &H40
    SHGFI_ICON = &H100
    SHGFI_DISPLAYNAME = &H200
    SHGFI_TYPENAME = &H400
    SHGFI_ATTRIBUTES = &H800
    SHGFI_ICONLOCATION = &H1000
    SHGFI_EXETYPE = &H2000
    SHGFI_SYSICONINDEX = &H4000
    SHGFI_LINKOVERLAY = &H8000
    SHGFI_SELECTED = &H10000
    SHGFI_ATTR_SPECIFIED = &H20000
End Enum

Private Const EXE_WIN16 = &H454E
Private Const EXE_DOS16 = &H5A4D
Private Const EXE_WIN32 = &H4550

Private Const MAX_PATH = 260

Private Type SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
    Reserved As String * 80
    reserved1 As String * 80
End Type
'#####################################################################################--Dateiinformationen

'Erweiterte Datei-Infos auslesen--########################################################################
'Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
'                                                        (ByVal lptstrFilename As String, _
'                                                        ByVal dwhandle As Long, _
'                                                        ByVal dwlen As Long, _
'                                                        lpData As Any) _
'                                                        As Long
'Public Declare Function GetFileVersionInfoSize Lib "Version.dll" _
'                                                        Alias "GetFileVersionInfoSizeA" ( _
'                                                        ByVal lptstrFilename As String, _
'                                                        lpdwHandle As Long) _
'                                                        As Long
'Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" _
'                                                        (pBlock As Any, _
'                                                        ByVal lpSubBlock As String, _
'                                                        lplpBuffer As Any, _
'                                                        puLen As Long) _
'                                                        As Long
'Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
'                                                        (dest As Any, _
'                                                        ByVal Source As Long, _
'                                                        ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
                                                        (ByVal lpString1 As String, _
                                                        ByVal lpString2 As Long) _
                                                        As Long

Public Type exFileInfo
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
End Type
'#########################################################################--Erweiterte Datei-Infos auslesen

'ActiveX-Komponenten registrieren--########################################################################
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
                                                        (ByVal lpLibFileName As String) _
                                                        As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) _
                                                        As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
                                                        ByVal dwMilliseconds As Long) _
                                                        As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
                                                        ByVal dwStackSize As Long, _
                                                        ByVal lpStartAddress As Long, _
                                                        ByVal lParameter As Long, _
                                                        ByVal dwCreationFlags As Long, _
                                                        lpThreadID As Long) _
                                                        As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
                                                        lpExitCode As Long) _
                                                        As Long
                                                    
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Const STATUS_WAIT_0 = &H0
'########################################################################--ActiveX-Komponenten registrieren

'Netzwerkmappings & Connections--##########################################################################
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, _
                                                        ByVal dwType As Long) _
                                                        As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, _
                                                        ByVal dwType As Long) _
                                                        As Long
Private Declare Function WNetUseConnection Lib "mpr.dll" _
                                                        Alias "WNetUseConnectionA" _
                                                        (ByVal hWndOwner As Long, _
                                                        ByRef lpNetResource As NETRESOURCE, _
                                                        ByVal lpUsername As String, _
                                                        ByVal lpPassword As String, _
                                                        ByVal dwFlags As Long, _
                                                        ByVal lpAccessName As Any, _
                                                        ByRef lpBufferSize As Long, _
                                                        ByRef lpResult As Long) _
                                                        As Long

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
                                                        (ByVal lpszLocalName As String, _
                                                        ByVal lpszRemoteName As String, _
                                                        cbRemoteName As Long) _
                                                        As Long
                                                        
Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
                                                        Alias "GetLogicalDriveStringsA" _
                                                        (ByVal nBufferLength As Long, _
                                                        ByVal lpbuffer As String) _
                                                        As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" _
                                                        Alias "WNetCancelConnection2A" _
                                                        (ByVal lpName As String, _
                                                        ByVal dwFlags As Long, _
                                                        ByVal fForce As Long) _
                                                        As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" _
                                                        (ByVal lpszName As String, _
                                                        ByVal bForce As Long) _
                                                        As Long
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" _
(ByVal lpszNetPath As String, _
                                                        ByVal lpszPassword As String, _
                                                        ByVal lpszLocalName As String) _
                                                        As Long
'Private Declare Function NetAccessCheck Lib "SVRAPI.dll" (ByVal pszReserved As String, _
'                                                        ByVal pszUserName As String, _
'                                                        ByVal pszResource As String, _
'                                                        ByVal usOperation As Integer, _
'                                                        ByRef pusResult As Integer) _
'                                                        As Long

' --- Notwendige Deklarationen
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF&
Private Const LANG_USER_DEFAULT             As Long = &H400&
Private Declare Function FormatMessage Lib "kernel32" _
                                                        Alias "FormatMessageA" _
                                                        (ByVal dwFlags As Long, _
                                                        ByRef lpSource As Any, _
                                                        ByVal dwMessageId As Long, _
                                                        ByVal dwLanguageId As Long, _
                                                        ByVal lpbuffer As String, _
                                                        ByVal nSize As Long, _
                                                        ByRef Arguments As Long) _
                                                        As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long
Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_ABSENT = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
Private Enum edw_Scope
    RESOURCE_CONNECTED = &H1
    RESOURCE_GLOBALNET = &H2
    RESOURCE_REMEMBERED = &H3
End Enum
Private Enum edw_Type
    RESOURCETYPE_ANY = &H0
    RESOURCETYPE_DISK = &H1
    RESOURCETYPE_PRINT = &H2
End Enum
Private Enum edw_DisplayType
    RESOURCEDISPLAYTYPE_GENERIC = &H0
    RESOURCEDISPLAYTYPE_DOMAIN = &H1
    RESOURCEDISPLAYTYPE_SERVER = &H2
    RESOURCEDISPLAYTYPE_SHARE = &H3
    RESOURCEDISPLAYTYPE_FILE = &H4
    RESOURCEDISPLAYTYPE_GROUP = &H5
    RESOURCEDISPLAYTYPE_NETWORK = &H6
    RESOURCEDISPLAYTYPE_ROOT = &H7
    RESOURCEDISPLAYTYPE_SHAREADMIN = &H8
    RESOURCEDISPLAYTYPE_DIRECTORY = &H9
    RESOURCEDISPLAYTYPE_TREE = &HA
    RESOURCEDISPLAYTYPE_NDSCONTAINER = &HB
End Enum
Private Enum edw_Usage
    RESOURCEUSAGE_CONNECTABLE = &H1
    RESOURCEUSAGE_CONTAINER = &H2
    RESOURCEUSAGE_NOLOCALDEVICE = &H4
    RESOURCEUSAGE_ATTACHED = &H10
    RESOURCEUSAGE_ALL = (RESOURCEUSAGE_CONNECTABLE Or RESOURCEUSAGE_CONTAINER Or RESOURCEUSAGE_ATTACHED)
    RESOURCEUSAGE_RESERVED = &H80000000
    RESOURCEUSAGE_SIBLING = &H8
End Enum
Private Type NETRESOURCE
    dwScope As edw_Scope
    dwType As edw_Type
    dwDisplayType As edw_DisplayType
    dwUsage As edw_Usage
    lpLocalName As String
    'Points to the name of a local device if the dwScope member is RESOURCE_CONNECTED or RESOURCE_REMEMBERED.
    'This member is NULL if the connection does not use a device. Otherwise, it is undefined.
    lpRemoteName As String
    'Points to a remote network name if the entry is a network resource.
    'If the entry is a current or persistent connection, lpRemoteName points
    'to the network name associated with the name pointed to by the lpLocalName member.
    lpComment As String
    'Points to a provider-supplied comment.
    lpProvider As String
    'Points to the name of the provider owning this resource.
    'This member can be NULL if the provider name is unknown.
End Type
Private Enum edw_Flags
    CONNECT_INTERACTIVE = &H8
    CONNECT_PROMPT = &H10
    CONNECT_REDIRECT = &H80
    CONNECT_UPDATE_PROFILE = &H1
    CONNECT_UPDATE_RECENT = &H2
    CONNECT_CURRENT_MEDIA = &H200
    CONNECT_DEFERRED = &H400
    CONNECT_LOCALDRIVE = &H100
    CONNECT_NEED_DRIVE = &H20
    CONNECT_REFCOUNT = &H40
    CONNECT_RESERVED = &HFF000000
    CONNECT_TEMPORARY = &H4
'    CONNECT_E_ADVISELIMIT = (CONNECT_E_FIRST + 1)
'    CONNECT_E_CANNOTCONNECT = (CONNECT_E_FIRST + 2)
'    CONNECT_E_NOCONNECTION = (CONNECT_E_FIRST + 0)
'    CONNECT_E_OVERRIDDEN = (CONNECT_E_FIRST + 3)
End Enum
Public Enum edw_Errors
    NO_ERROR = 0
    ERROR_CONNECTION_UNAVAIL = 1201&
    ERROR_NOT_SUPPORTED = 50&
    ERROR_NOT_CONNECTED = 2250&
    ERROR_ACCESS_DENIED = 5&
    ERROR_ALREADY_ASSIGNED = 85&
    ERROR_BAD_DEVICE = 1200&
    ERROR_BAD_NET_NAME = 67&
    ERROR_BAD_PROVIDER = 1204&
    ERROR_CANCELLED = 1223&
    ERROR_EXTENDED_ERROR = 1208&
    ERROR_INVALID_ADDRESS = 487&
    ERROR_INVALID_PARAMETER = 87
    ERROR_INVALID_PASSWORD = 86&
    ERROR_MORE_DATA = 234
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_NO_NET_OR_BAD_PATH = 1203&
    ERROR_NO_NETWORK = 1222&
    ERROR_NO_SPECIFIED = -1&
End Enum
Private Declare Function IsNetDrive Lib "Shell32" (ByVal iDrive As Long) As Long
Private Const NERR_Success = 0
Private Const NERR_BASE = 2100
Private Const NERR_InvalidComputer = (NERR_BASE + 251)
Private Const NERR_UseNotFound = (NERR_BASE + 150)
Private Const CP_ACP = 0
Private Type USER_INFO_3
    usri3_name As Long
    usri3_password As Long
    usri3_password_age As Long
    usri3_priv As Long
    usri3_home_dir As Long
    usri3_comment As Long
    usri3_flags As Long
    usri3_script_path As Long
    usri3_auth_flags As Long
    usri3_full_name As Long
    usri3_usr_comment As Long
    usri3_parms As Long
    usri3_workstations As Long
    usri3_last_logon As Long
    usri3_last_logoff As Long
    usri3_acct_expires As Long
    usri3_max_storage As Long
    usri3_units_per_week As Long
    usri3_logon_hours As Byte
    usri3_bad_pw_count As Long
    usri3_num_logons As Long
    usri3_logon_server As String
    usri3_country_code As Long
    usri3_code_page As Long
    usri3_user_id As Long
    usri3_primary_group_id As Long
    usri3_profile As Long
    usri3_home_dir_drive As Long
    usri3_password_expired As Long
End Type
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Any, Username As Byte, ByVal Level As Long, lpbuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Const NUL As Long = 0
'Netzwerkmappings & Connections--##########################################################################

Private Declare Function GetShortPathName Lib "kernel32" _
                                                        Alias "GetShortPathNameA" _
                                                        (ByVal lpszLongPath As String, _
                                                        ByVal lpszShortPath As String, _
                                                        ByVal cchBuffer As Long) _
                                                        As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" _
                                                        Alias "GetLongPathNameA" _
                                                        (ByVal lpszShortPath As String, _
                                                        ByVal lpszLongPath As String, _
                                                        ByVal cchBuffer As Long) _
                                                        As Long
'Datei und Verzeichnis Browsing--##########################################################################
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                                                        (ByVal lpFileName As String, _
                                                        lpFindFileData As WIN32_FIND_DATA) _
                                                        As Long
        
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                                                        (ByVal hFindFile As Long, _
                                                        lpFindFileData As WIN32_FIND_DATA) _
                                                        As Long
        
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'##########################################################################--Datei und Verzeichnis Browsing
        
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Type tWshShortcut  ' Verknüpfung (Link) auslesen Shortcut-Struktur
    Arguments As String
    Description As String
    FullName As String
    Hotkey As String
    IconLocation As String
    TargetPath As String
    WindowStyle As Long
    WorkingDirectory As String
End Type

Private strStartPath As String

Public Function ExecuteShellLink(ByVal FullPath As String) As Boolean ' Verknüpfung (Link) ausführen
On Error GoTo ErrHandler
'Dim wsh As WshShell
'Dim Sht As WshShortcut
'
'    Set wsh = New WshShell
'    Set Sht = wsh.CreateShortcut(FullPath)
Dim wsh As Object
Dim Sht As Object

    Set wsh = CreateObject("Wscript.Shell")
    Set Sht = wsh.CreateShortcut(FullPath)
    
    wsh.Exec Sht.TargetPath & " " & Sht.Arguments
    
    Set wsh = Nothing
    Set Sht = Nothing
    ExecuteShellLink = True
Exit Function
ErrHandler:
    ExecuteShellLink = False
    Set wsh = Nothing
    Set Sht = Nothing
End Function

Public Function ReadShellLink(ByVal FullPath As String) As tWshShortcut  ' Verknüpfung (Link) auslesen
On Error GoTo ErrHandler
'Dim wsh As WshShell
'Dim Sht As WshShortcut
'
'    Set wsh = New WshShell
'    Set Sht = wsh.CreateShortcut(FullPath)
Dim wsh As Object
Dim Sht As Object

    Set wsh = CreateObject("Wscript.Shell")
    Set Sht = wsh.CreateShortcut(FullPath)
    
    With ReadShellLink
        .Arguments = Sht.Arguments
        .Description = Sht.Description
        .FullName = Sht.FullName
        .Hotkey = Sht.Hotkey
        .IconLocation = Sht.IconLocation
        .TargetPath = Sht.TargetPath
        .WindowStyle = Sht.WindowStyle
        .WorkingDirectory = Sht.WorkingDirectory
    End With
    wsh.Exec Sht.TargetPath & " " & Sht.Arguments
Set wsh = Nothing
Set Sht = Nothing

Exit Function
ErrHandler:
    Set wsh = Nothing
    Set Sht = Nothing
End Function

Public Sub Wait(TimeToWait)
Dim sTime As Double
    sTime = timeGetTime
    Do
        DoEvents
        Sleep 1
    Loop While Abs(timeGetTime - TimeToWait) < sTime
End Sub

Public Function TrimNull(strString As String) As String
Dim intPos As Integer
    ' Remove any Nulls that strings returned from the Windows API
    ' might happen to have embedded.  Don't send this function a
    ' Null string.  It won't like it.
    intPos = InStr(strString, vbNullChar)
    If intPos > 0 Then
        TrimNull = Trim(Left(strString, intPos - 1))
    Else
        TrimNull = Trim(strString)
    End If
End Function

Public Function ANSIToUni(varAnsi As Variant) As Variant
    ' Convert an ANSI string to Unicode.
    ANSIToUni = StrConv(varAnsi, vbUnicode)
End Function
Public Function UniToAnsi(varUni As Variant) As Variant
    ' Convert a Unicode string to ANSI.
    UniToAnsi = StrConv(varUni, vbFromUnicode)
End Function
' Returns an ANSI string from a pointer to a Unicode string.
Public Function GetStrFromPtrW(lpszW As Long) As String
    Dim sRtn As String
    sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
    ' WideCharToMultiByte also returns Unicode string length
    Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
    GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function
' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
    If InStr(sz, vbNullChar) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
    Else
        ' If sz had no null char, the Left$ function
        ' above would return a zero length string ("").
        GetStrFromBufferA = sz
    End If
End Function

Public Function APIErrorDescription(ByVal ErrLastDllError As Long) As String
' Liefert die Klartextbeschreibung zu einer API Fehlernummer, die
' unter Visual Basic über Err.LastDllError ermittelt wurde.
' HINWEIS: Die API-Funktion GetLastError ist für Visual Basic tabu!
Dim sBuffer    As String  ' String für die Rückgabe des Fehlertexts
Dim lBufferLen As Long    ' Länge des reservierten Strings
    ' Stringbuffer für die Rückgabe reservieren
    sBuffer = Space$(1024)
    ' Fehlernummer in einen Fehlertext wandeln
    lBufferLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                FORMAT_MESSAGE_MAX_WIDTH_MASK Or _
                                FORMAT_MESSAGE_IGNORE_INSERTS, _
                                ByVal 0&, ErrLastDllError, _
                                LANG_USER_DEFAULT, _
                                sBuffer, Len(sBuffer), 0)
    If lBufferLen > 0 Then
      ' Fehler wurde identifiziert, der Fehlertext liegt vor
      APIErrorDescription = Left$(sBuffer, lBufferLen)
    Else
      ' Der Fehlertext konnte nicht ermittelt werden
      APIErrorDescription = "Unbekannter Fehler: &H" & Hex$(ErrLastDllError)
    End If
End Function

Public Function fc_Get_WNetErrorDescription(ByVal dw_Error As edw_Errors) As String
    fc_Get_WNetErrorDescription = dw_Error & " " & APIErrorDescription(dw_Error)
End Function

'Netzwerkmappings & Connections--##########################################################################
Public Function fc_IsNetDrive(ByVal Drive As String) As Boolean
On Error GoTo ErrOut
    'Drive = LCase(Left$(Drive, 1) & ":")
    fc_IsNetDrive = (IsNetDrive(Asc(UCase(Drive)) - 65) <> 0)
ErrOut:
End Function

Public Function fc_NetGetUNCPath(MappedDrive As String, ByRef UncPath As String, Optional ErrInfo As Long) As Boolean
    On Local Error GoTo NetGetUNCPath_Err
Dim lpszRemoteName As String
Dim cbRemoteName As Long
    fc_NetGetUNCPath = False
    lpszRemoteName = String$(255, Chr$(32))
    cbRemoteName = Len(lpszRemoteName)
    MappedDrive = LCase(Left$(MappedDrive, 1) & ":")
    ErrInfo = WNetGetConnection(MappedDrive, lpszRemoteName, cbRemoteName)
    
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
    ' Check for success
    If (ErrInfo = NO_ERROR Or ErrInfo = 1201) Then
        UncPath = TrimNull(Left$(lpszRemoteName, cbRemoteName))
        fc_NetGetUNCPath = True
    End If
Exit Function
NetGetUNCPath_Err:
    fc_NetGetUNCPath = False
    ErrInfo = ERROR_NO_SPECIFIED
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
End Function

Public Function fc_NetReConnect(ByVal MappedDrive As String, Optional ByVal pwd As String, _
                                Optional ByRef UNC As String, Optional ByRef ErrInfo As Long) As Boolean
    fc_NetReConnect = False
    MappedDrive = Left$(MappedDrive, 1) & ":"
    If fc_NetGetUNCPath(MappedDrive, UNC, ErrInfo) Then
        If fc_NetCancelConnection(MappedDrive, True, ErrInfo) Then
            DoEvents: Sleep 10
            If fc_NetAddConnection(UNC, MappedDrive, pwd, ErrInfo) Then
                DoEvents: Sleep 10
                fc_NetReConnect = True
            End If
        End If
    End If
End Function

Public Function fc_NetAddConnection(ByVal UNC As String, ByVal MappedDrive As String, Optional ByVal pwd As String, _
                             Optional ErrInfo As Long) As Boolean
' Versucht, ein freigegebenes Verzeichnis im Netzwerk (UNCPath)
' als lokales Netzwerklaufwerk (LocalPath) einzubinden. Im Fall
' kennwortgeschützter freigegebener Verzeichnisse im Netzwerk
' kann im Parameter Password ein Kennwort übergeben werden.
On Local Error GoTo AddConnection_Err
    fc_NetAddConnection = False
    MappedDrive = Left$(MappedDrive, 1) & ":"
    ErrInfo = WNetAddConnection(UNC, pwd, MappedDrive)
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
    ' Check for success
    If (ErrInfo = NO_ERROR) Then
        fc_NetAddConnection = True
    End If
Exit Function
AddConnection_Err:
    fc_NetAddConnection = False
    ErrInfo = ERROR_NO_SPECIFIED
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
End Function

Public Function fc_NetCancelConnection(ByVal MappedDrive As String, Optional ByVal Force As Boolean = False, Optional ErrInfo As Long) As Boolean
' Versucht, ein Netzwerklaufwerk abzumelden. Wird Force zu True
' gesetzt, werden mögliche Probleme ignoriert (ggf. "auf Kosten"
' eines Netzwerkteilnehmers).
On Local Error GoTo CancelConnection_Err
    fc_NetCancelConnection = False
    MappedDrive = Left$(MappedDrive, 1) & ":"
    ErrInfo = WNetCancelConnection(MappedDrive, CLng(Force))
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
    ' Check for success
    If (ErrInfo = NO_ERROR) Then
        fc_NetCancelConnection = True
    End If
Exit Function
CancelConnection_Err:
    fc_NetCancelConnection = False
    ErrInfo = ERROR_NO_SPECIFIED
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
End Function

Public Function fc_NetCancelConnection2(ByVal MappedDrive As String, Optional ErrInfo As Long) As Boolean
On Local Error GoTo CancelConnection_Err
    fc_NetCancelConnection2 = False
    ' Call API to disconnect the drive
    MappedDrive = Left$(MappedDrive, 1) & ":"
    ErrInfo = WNetCancelConnection2(MappedDrive, CONNECT_UPDATE_PROFILE, False)
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
    ' Check for success
    If (ErrInfo = NO_ERROR) Then
        fc_NetCancelConnection2 = True
    End If
Exit Function
CancelConnection_Err:
    fc_NetCancelConnection2 = False
    ErrInfo = ERROR_NO_SPECIFIED
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
End Function


Public Function fc_NetUseConnection(ByVal UNC As String, ByVal pwd As String, ByVal User As String, ByRef MappedDrive As String) As Boolean
    Dim NetR As NETRESOURCE    ' NetResouce structure
    Dim ErrInfo As Long        ' Return value from API
    Dim Buffer As String       ' Drive letter assigned to resource
    Dim bufferlen As Long      ' Size of the buffer
    Dim Success As Long        ' Additional info about API call
    fc_NetUseConnection = False
    ' Initialize the NetResouce structure
    MappedDrive = Left$(MappedDrive, 1) & ":"
    NetR.dwScope = RESOURCE_CONNECTED
    NetR.dwType = RESOURCETYPE_DISK
    NetR.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    NetR.dwUsage = RESOURCEUSAGE_CONNECTABLE
    NetR.lpLocalName = MappedDrive
    NetR.lpRemoteName = UNC

    ' Initialize the return buffer and buffer size
    Buffer = Space(32)
    bufferlen = Len(Buffer)

    ' Call API to map the drive
    ErrInfo = WNetUseConnection(0&, NetR, pwd, User, CONNECT_REDIRECT, Buffer, bufferlen, Success)
    Debug.Print fc_Get_WNetErrorDescription(ErrInfo)
    ' Check if call to API failed. According to the MSDN help, there
    ' are some versions of the operating system that expect the userid
    ' as the 3rd parameter and the password as the 4th, while other
    ' versions of the operating system have them in reverse order, so
    ' if first call to API fails, try reversing these two parameters.
    If ErrInfo <> NO_ERROR Then
        ' Call API with userid and password switched
        ErrInfo = WNetUseConnection(0&, NetR, pwd, User, CONNECT_REDIRECT, Buffer, bufferlen, Success)
        fc_Get_WNetErrorDescription (ErrInfo)
    End If
    
    ' Check for success
    If (ErrInfo = NO_ERROR) And (Success = CONNECT_LOCALDRIVE) Then
        ' Store the mapped drive letter for later usage
        MappedDrive = Left$(Buffer, InStr(1, Buffer, ":"))
        fc_NetUseConnection = True
    End If
End Function

'##########################################################################--Netzwerkmappings & Connections


Public Function fc_DirectoryPathAccess(ByVal DirPath As String) As Boolean
    fc_DirectoryPathAccess = False
    If Len(DirPath) > 0 Then
        fc_CheckDriveAccess fc_Path(DirPath), fc_DirectoryPathAccess
    End If
End Function

Public Function fc_GetDrives() As String
'Returns all mapped drives
    Dim lngRet As Long
    Dim strDrives As String * 255
    Dim lngTmp As Long
    lngTmp = Len(strDrives)
    lngRet = GetLogicalDriveStrings(lngTmp, strDrives)
    fc_GetDrives = Left(strDrives, lngRet)
End Function

Public Function fc_DriveTypeDescription(ByVal DriveName As String, Optional ByRef DriveType As Long) As String
    Dim strDrive As String
    DriveType = GetDriveType(DriveName)
    Select Case DriveType
        Case DRIVE_UNKNOWN 'The drive type cannot be determined.
            strDrive = "Unknown Drive Type"
        Case DRIVE_ABSENT 'The root directory does not exist.
            strDrive = "Drive does not exist"
        Case DRIVE_REMOVABLE 'The drive can be removed from the drive.
            strDrive = "Removable Media"
        Case DRIVE_FIXED 'The disk cannot be removed from the drive.
            strDrive = "Fixed Drive"
        Case DRIVE_REMOTE  'The drive is a remote (network) drive.
            strDrive = "Network Drive"
        Case DRIVE_CDROM 'The drive is a CD-ROM drive.
            strDrive = "CD Rom"
        Case DRIVE_RAMDISK 'The drive is a RAM disk.
            strDrive = "Ram Disk"
    End Select
    fc_DriveTypeDescription = strDrive
End Function

Public Function fc_ListAllDrives(ByRef DrivesList() As String) As Long
Dim strAllDrives As String
Dim strTmp As String
Dim Count As Integer
Dim UncPath As String
    strAllDrives = fc_GetDrives
    If strAllDrives <> "" Then
        Do
            strTmp = Mid$(strAllDrives, 1, InStr(strAllDrives, vbNullChar) - 1)
            strAllDrives = Mid$(strAllDrives, InStr(strAllDrives, vbNullChar) + 1)
            Count = Count + 1
        Loop While strAllDrives <> ""
        ReDim DrivesList(Count - 1, 2)
        Count = 0
        strAllDrives = fc_GetDrives
        Do
            strTmp = Mid$(strAllDrives, 1, InStr(strAllDrives, vbNullChar) - 1)
            strAllDrives = Mid$(strAllDrives, InStr(strAllDrives, vbNullChar) + 1)
            DrivesList(Count, 0) = fc_DriveTypeDescription(strTmp)
            DrivesList(Count, 1) = strTmp
            Select Case fc_DriveTypeDescription(strTmp)
                Case "Network Drive":
                    fc_NetGetUNCPath Left$(strTmp, Len(strTmp) - 1), UncPath
                    DrivesList(Count, 2) = UncPath
            End Select
            Count = Count + 1
        Loop While strAllDrives <> ""

        fc_ListAllDrives = Count
    End If
End Function

Public Function fc_CheckDriveAccess(ByVal Path As String, Optional ByRef AccessRead As Boolean, Optional ByRef AccessWrite As Boolean) As String
On Error GoTo NoAccess
Dim fN As Long
    If Len(Dir(Path, vbDirectory)) > 0 Then
        AccessRead = True
        On Error Resume Next
        fN = FreeFile
        Open fc_Korrect_Path(Path) & "CheckDriveAccess.txt" For Output As #fN
            Print #fN, "CheckDriveAccess Teststring"
        Close #fN
        If Err.Number <> 0 Then
            AccessWrite = False
            Kill fc_Korrect_Path(Path) & "CheckDriveAccess.txt"
        Else
            AccessWrite = True
            Kill fc_Korrect_Path(Path) & "CheckDriveAccess.txt"
        End If
    Else
        GoTo NoAccess
    End If
    fc_CheckDriveAccess = IIf(AccessRead, "Lesezugriff, ", "kein Lesezugriff, ") & _
                        IIf(AccessRead, "Schreibzugriff, ", "Schreibzugriff, ") & _
                        "auf """ & Path & """"
Exit Function
NoAccess:
    fc_CheckDriveAccess = "kein Zugriff auf """ & Path & """"
    AccessRead = False
    AccessWrite = False
End Function

Public Function fc_GetFileDescription(ByVal FullPath As String) As String
On Error GoTo ErrOut
Dim lF As Long
Dim ShellInfo As SHFILEINFO
    Call SHGetFileInfo(FullPath, 0, ShellInfo, Len(ShellInfo), SHGFI_TYPENAME)
    lF = InStr(1, ShellInfo.szTypeName, Chr$(0)) - 1
    fc_GetFileDescription = Left(ShellInfo.szTypeName, lF)
Exit Function
ErrOut:
    fc_GetFileDescription = ""
End Function
  
Public Function fc_GetFileIcon(ByVal FullPath As String, Optional IconSize As IconSize = Large) As IPictureDisp
On Error GoTo ErrOut
  Dim result As Long
  Dim Unkown As IUnknown
  Dim Icon As IconType
  Dim CLSID As CLSIdType
  Dim ShellInfo As SHFILEINFO
    FullPath = fc_Get_Short_Path(FullPath)
    Call SHGetFileInfo(FullPath, 0, ShellInfo, Len(ShellInfo), IconSize)
    
    Icon.cbSize = Len(Icon)
    Icon.picType = vbPicTypeIcon
    Icon.hIcon = ShellInfo.hIcon
    CLSID.id(8) = &HC0
    CLSID.id(15) = &H46
    result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    
    Set fc_GetFileIcon = Unkown
Exit Function
ErrOut:
    Set fc_GetFileIcon = Nothing
End Function

Function fc_GetExeType(ByVal FullPath As String) As String
On Error GoTo ErrOut
   Dim dwExeVal As Long
   Dim shfi As SHFILEINFO
   Dim dwLowWord As Long
   Dim dwHighWord As Long
   Dim bHighWordLowByte As Byte
   Dim bHighWordHighByte As Byte
   Dim sRtn As String

   dwExeVal = SHGetFileInfo(FullPath, 0&, shfi, Len(shfi), SHGFI_EXETYPE)
   dwLowWord = dwExeVal And &HFFFF&

   Select Case dwLowWord
      Case 0
         sRtn = "(nicht ausführbar)"
      Case EXE_WIN16
         sRtn = "16 Bit Windows"
      Case EXE_DOS16
         sRtn = "DOS"
      Case EXE_WIN32
         sRtn = "32 Bit Windows"
      Case Else
         sRtn = "(unbekannt)"
   End Select
ErrOut:
   fc_GetExeType = sRtn
End Function

'Erweiterte Datei-Infos auslesen--#####################################################################
Public Function fc_GetExtendedFileInfo(ByVal FullPath As String, Optional ByRef Success As Boolean) As exFileInfo
Dim lngBufferlen As Long
Dim lngDummy As Long
Dim lngRc As Long
Dim lngVerPointer As Long
Dim lngHexNumber As Long
Dim bytBuffer() As Byte
Dim bytBuff(255) As Byte
Dim strBuffer As String
Dim strLangCharset As String
Dim strVersionInfo(7) As String
Dim strTemp As String
Dim intTemp As Integer
Success = False
    On Error GoTo 0
    lngBufferlen = GetFileVersionInfoSize(FullPath, lngDummy)
    If lngBufferlen > 0 Then
        ReDim bytBuffer(lngBufferlen)
        lngRc = GetFileVersionInfo(FullPath, 0&, lngBufferlen, bytBuffer(0))
        If lngRc <> 0 Then
            lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
            If lngRc <> 0 Then
                Success = True
                MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
                lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
                strLangCharset = Hex(lngHexNumber)
                Do While Len(strLangCharset) < 8
                    strLangCharset = "0" & strLangCharset
                Loop
                ' assign propertienames
                strVersionInfo(0) = "CompanyName"
                strVersionInfo(1) = "FileDescription"
                strVersionInfo(2) = "FileVersion"
                strVersionInfo(3) = "InternalName"
                strVersionInfo(4) = "LegalCopyright"
                strVersionInfo(5) = "OriginalFileName"
                strVersionInfo(6) = "ProductName"
                strVersionInfo(7) = "ProductVersion"
                ' loop and get EXFILEINFOs
                For intTemp = 0 To 7
                    strBuffer = String$(255, 0)
                    strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(intTemp)
                    lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                    If lngRc <> 0 Then
                        ' get and format data
                        lstrcpy strBuffer, lngVerPointer
                        strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                        strVersionInfo(intTemp) = strBuffer
                    Else
                        ' property not found
                        strVersionInfo(intTemp) = "?"
                    End If
                Next intTemp
            End If
        End If
    End If
    ' assign array to user-defined-type
    fc_GetExtendedFileInfo.CompanyName = strVersionInfo(0)
    fc_GetExtendedFileInfo.FileDescription = strVersionInfo(1)
    fc_GetExtendedFileInfo.FileVersion = strVersionInfo(2)
    fc_GetExtendedFileInfo.InternalName = strVersionInfo(3)
    fc_GetExtendedFileInfo.LegalCopyright = strVersionInfo(4)
    fc_GetExtendedFileInfo.OriginalFileName = strVersionInfo(5)
    fc_GetExtendedFileInfo.ProductName = strVersionInfo(6)
    fc_GetExtendedFileInfo.ProductVersion = strVersionInfo(7)
End Function
'#####################################################################--Erweiterte Datei-Infos auslesen
 

Public Function fc_Path(ByVal FullPath As String) As String
Dim i As Integer
    ' Gibt den Pfad einer Datei aus dem FullPath zurück
    ' ©2002 by Marco Großert
    On Error GoTo Error
    If Len(FullPath) > 0 Then
        i = InStrRev(FullPath, "\")
        If i = 0 Then
            fc_Path = FullPath
        Else
            fc_Path = Left$(FullPath, i)
        End If
    End If
Error:

End Function

'18.11.2007 fc_CreatePath funktionierte Fehlerhaft bei UNC-Pfaden mit \\Server\...----------------------------------
'Alte Version:
'Public Function fc_CreatePath(FullPath) As Boolean
'' Erstellt die in Fullpath enthaltene Verzeichnisstruktur wenn nicht vorhanden
'' ©2004 by Marco Großert       email: marco@grossert.com
'On Error GoTo ErrOut
'Dim tmpFolders() As String
'Dim CurrentPath As String
'Dim i As Integer
'
'    tmpFolders = Split(FullPath, "\")
'    For i = 0 To UBound(tmpFolders)
'        CurrentPath = fc_Korrect_Path(CurrentPath & tmpFolders(i))
'        If Not fc_Folder_Exists(CurrentPath) Then
'            MkDir CurrentPath
'        End If
'    Next i
'
'    fc_CreatePath = True
'Exit Function
'ErrOut:
'    fc_CreatePath = False
'End Function
'Neue Version:
Public Function fc_CreatePath(ByVal FullPath As String) As Boolean
' Erstellt die in Fullpath enthaltene Verzeichnisstruktur wenn nicht vorhanden
' ©2007 by Marco Großert       email: marco@grossert.com
On Error GoTo ErrOut
Dim tmpNewFolders() As String
Dim CurrentPath As String
Dim i As Integer, j As Integer

    FullPath = fc_Korrect_Path(FullPath)
    FullPath = Mid(FullPath, 1, (Len(FullPath) - 1))
    If InStr(1, FullPath, fc_AppPath) > 0 Then
        tmpNewFolders() = Split(Mid(FullPath, Len(fc_AppPath)), "\")
        CurrentPath = fc_AppPath
        j = 1
    Else
        tmpNewFolders() = Split(FullPath, "\")
        If Left(FullPath, 2) = "\\" Then
            j = 3
            CurrentPath = "\\" & tmpNewFolders(2) & "\"
        ElseIf Mid(FullPath, 2, 2) = ":\" Then
            j = 1
            CurrentPath = Left(FullPath, 3)
        Else
            j = 1
            CurrentPath = fc_AppPath
        End If
    End If

    For i = j To UBound(tmpNewFolders)
        CurrentPath = fc_Korrect_Path(CurrentPath & tmpNewFolders(i))
        If Not fc_Folder_Exists(CurrentPath) Then
            MkDir CurrentPath
        End If
    Next i
    
    fc_CreatePath = True
Exit Function
ErrOut:
    fc_CreatePath = False
End Function
'18.11.2007 fc_CreatePath funktionierte Fehlerhaft bei UNC-Pfaden mit \\Server\...----------------------------------


Public Function fc_FileSize(ByVal FullPath As String) As Long
    ' Gibt die Größe einer Datei in Byte zurück
    ' ©2001 by Marco Großert
    On Error GoTo Error
    fc_FileSize = FileLen(FullPath)
    Exit Function
Error:
    fc_FileSize = "0"
End Function

Public Function FormatMemorySize(ByVal lFileSize As Long) As String
    Select Case lFileSize
        Case Is < (2 ^ 10)
            FormatMemorySize = lFileSize & " Byte"
        Case Is < (2 ^ 20)
            FormatMemorySize = Round((lFileSize / (2 ^ 10)), 2) & " KB"
        Case Is < (2 ^ 30)
            FormatMemorySize = Round((lFileSize / (2 ^ 20)), 2) & " MB"
        Case Else
            FormatMemorySize = Round((lFileSize / (2 ^ 30)), 2) & " GB"
    End Select
End Function

Public Function fc_FileCreateDate(ByVal FullPath As String) As String
    ' Gibt das Erstelldatum einer Datei zurück
    ' ©2001 by Marco Großert
    On Error GoTo Error
  
    fc_FileCreateDate = FileDateTime(FullPath)
    Exit Function
Error:

End Function

Public Function fc_DirExist(ByVal FullPath As String) As Boolean
' Prüft ob das Verzeichnis 'FullPath'  Existiert
' ©2001 by Marco Großert       email: marco@grossert.com
    On Error Resume Next
    fc_DirExist = Not CBool(GetAttr(FullPath) And (vbDirectory Or vbVolume))
    Err.Clear
End Function

Public Function fc_FileExist(ByVal FullPath As String) As Boolean
' Prüft ob die Datei 'FullPath' im in 'FullPath' angegebenen  Pfad Existiert
' ©2001 by Marco Großert
On Error GoTo NoExist
    fc_FileExist = Not CBool(GetAttr(FullPath) And (vbDirectory Or vbVolume))
    Exit Function
NoExist:
    Err.Clear
    fc_FileExist = False
End Function

Public Function fc_FileTitel(ByVal FileName As String) As String
    ' Gibt den Dateinamen ohne Endung aus 'Filename' zurück
    ' ©2001 by Marco Großert
    On Error GoTo Error
    Dim i As Integer
    For i = Len(FileName) To 1 Step -1
        If Mid$(FileName, i, 1) = "." Then
            fc_FileTitel = Left$(FileName, i - 1)
            Exit For
        End If
    Next
    Exit Function
Error:

End Function

Public Function fc_FileName(ByVal FullPath As String) As String
' Gibt den kompletten Dateinamen aus 'FullPath' zurück
' ©2001 by Marco Großert
Dim i As Integer
    If FullPath > "" Then
        For i = 1 To Len(FullPath)
            fc_FileName = Right$(FullPath, i)
            If Left$(Right$(FullPath, i + 1), 1) = "\" Then Exit For
        Next
    End If
End Function

Public Function fc_FileExtension(ByVal FileName As String, Optional WithDot As Boolean = True) As String
' Gibt die Dateiextension aus 'Filename' zurück
' ©2006 by Marco Großert
Dim Pos As Integer
    If Len(FileName) > 0 Then
        Pos = InStrRev(FileName, ".", -1)
        If Pos > 0 Then fc_FileExtension = Mid(FileName, IIf(WithDot, Pos, Pos + 1))
    End If
End Function

Public Function fc_Get_Short_Path(ByVal LongPath As String) As String
'Wandelt langen Pfad(Windows) in kurzen Pfad (DOS mit Tilde)
' ©2001 by Marco Großert '29.08.2001'
Dim intPath_Len As Integer
Dim tmpPath As String * 255
    If InStr(1, LongPath, "~") = 0 Then
        intPath_Len = GetShortPathName(LongPath, tmpPath, 255)
        fc_Get_Short_Path = fc_Korrect_Path(Left(tmpPath, intPath_Len))
    Else
        fc_Get_Short_Path = fc_Korrect_Path(LongPath)
    End If
End Function

Public Function fc_Get_Long_Path(ByVal ShortPath As String) As String
'Wandelt kurzen Pfad(DOS mit Tilde) in langen Pfad(Windows)
' ©2001 by Marco Großert '29.08.2001'
Dim intPath_Len As Integer
Dim tmpPath As String * 255
    If InStr(1, ShortPath, "~") > 0 Then
        intPath_Len = GetLongPathName(ShortPath, tmpPath, 255)
        fc_Get_Long_Path = fc_Korrect_Path(Left(tmpPath, intPath_Len))
    Else
        fc_Get_Long_Path = fc_Korrect_Path(ShortPath)
    End If
End Function

Public Function fc_Parent_Path(ByVal FullPath As String) As String
' Gibt das Verzeichnis das im Verzeichnisbaum über dem Verzeichnis in 'Fullpath' steht zurück
' ©2003 by Marco Großert
Dim i As Integer
    For i = Len(FullPath) To 0 Step -1
        If Mid(FullPath, i, 1) = "\" Then
            fc_Parent_Path = Left(FullPath, i - 1)
            Exit For
        End If
    Next i
End Function

Public Function fc_TempPath() As String
Dim strTemp As String
    'Create a buffer
    strTemp = String(100, Chr$(0))
    'Get the temporary path
    GetTempPath 100, strTemp
    'strip the rest of the buffer
    strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
    fc_TempPath = fc_Korrect_Path(strTemp)
End Function

Public Function fc_UNC(ByVal FullPath As String) As String
' Gibt den in 'Fullpath' übergebemen Pfad als UNC-Pfad zurück
' ©2003 by Marco Großert
    fc_UNC = Replace(FullPath, "\", "/")
End Function

'Public Function fc_Folder(ByVal FullPath As String) As String
'' Gibt den Namen des aktuellen Verzeichnisses aus 'Fullpath' zurück
'' ©2003 by Marco Großert
'On Error GoTo ErrOut
'Dim i As Integer
'    For i = Len(FullPath) To 0 Step -1
'        If Mid(FullPath, i, 1) = "\" Then
'            fc_Folder = Right(FullPath, Len(FullPath) - i)
'            Exit For
'        End If
'    Next i
'    Exit Function
'ErrOut:
'    fc_Folder = vbNullString
'End Function

Public Function fc_Folder(ByVal FullPath As String) As String
' Gibt den Namen des aktuellen Verzeichnisses aus 'Fullpath' zurück
' ©2003 by Marco Großert
On Error GoTo ErrOut
Dim Pos As Integer
    Pos = InStrRev(FullPath, "\")
    If Pos = Len(FullPath) Then
        FullPath = Left(FullPath, Pos - 1)
        Pos = InStrRev(FullPath, "\")
    End If
    If Pos > 0 Then
        fc_Folder = Right(FullPath, Len(FullPath) - Pos)
    End If
 
    Exit Function
ErrOut:
    fc_Folder = vbNullString
End Function

Public Function fc_Folder_Exists(FullPath) As Boolean
On Error Resume Next
' Prüft ob der in 'Fullpath' übergbene Pfad existiert
' ©2004 by Marco Großert
   fc_Folder_Exists = Not (Len(Dir$(FullPath, vbDirectory)) = 0)
End Function

Public Function fc_AppPath() As String
' Gibt den erforderlichenfalls korregierten Application -Pfad mit abschließendem Backslash zurück
' ©2004 by Marco Großert
    fc_AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function

Public Function fc_Korrect_Path(FullPath As String) As String
' Gibt den erforderlichenfalls korregierten Pfad mit abschließendem Backslash zurück
' ©2004 by Marco Großert
    fc_Korrect_Path = IIf(Right(FullPath, 1) = "\", FullPath, FullPath & "\")
End Function

Public Function fc_Drive(ByVal FullPath As String) As String
' Gibt den Laufwerksbuchstaben (inclusive ':') aus dem in 'Fullpath' übergebenen Pfad zurück
' ©2004 by Marco Großert
On Error GoTo Err_Out
Dim i As Integer
    For i = Len(FullPath) To 0 Step -1
        If Mid(FullPath, i, 1) = ":" Then
            fc_Drive = Left(FullPath, i)
            Exit For
        End If
    Next i
Err_Out:
End Function

Public Function fc_Rel_UNC_Path(ByVal Path1 As String, ByVal Path2 As String) As String
Dim i As Integer
    For i = Len(Path1) To Len(Path2)
        If Mid(Path2, i, 1) = "/" Or Mid(Path2, i, 1) = "\" Then
            fc_Rel_UNC_Path = fc_Rel_UNC_Path & "../"
        End If
    Next i
End Function

Public Function fc_GetPathFromRelPath(ByVal FullPath As String, ByVal RelPath As String) As String
    Do
        If Left(RelPath, 3) = "..\" Then
            FullPath = fc_Parent_Path(FullPath)
            RelPath = Mid(RelPath, 4)
        Else
            Exit Do
        End If
    Loop While InStr(1, RelPath, "..\") > 0
    fc_GetPathFromRelPath = fc_Korrect_Path(FullPath) & RelPath
End Function

Public Function fc_DelFiles(ByRef FullPath() As Variant) As Boolean

' Löscht die im ByRef übergebenen Array 'Fullpath()' Dateien
' un zwar so das sie im Windows FileSystem Recycled werden können
' mit anderen Worten, sie landen im Papierkorb
' Wenn diese Funktion erfolgreich war gibt sie True zurück
' ©2001 by Marco Großert '29.08.2001'

'!! folgene API- incl. Typ & Konstanten  -Deklaration muß im Deklarationsbereich existieren----------
'''' Datei OP's (zB. in Papierkorb)
'''Public Declare Function SHFileOperation Lib "shell32.dll" _
'''        Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) _
'''        As Long
'''
'''Public Type SHFILEOPSTRUCT
'''        hwnd As Long
'''        wFunc As Long
'''        pFrom As String
'''        pTo As String
'''        fFlags As Integer
'''        fAnyOperationsAborted As Long
'''        hNameMappings As Long
'''        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
'''End Type
'''
'''Public Const FO_DELETE = &H3
'''Public Const FOF_ALLOWUNDO = &H40
'''Public Const FOF_SIMPLEPROGRESS = &H100
'''Public Const FOF_NOCONFIRMATION = &H10
'''Public Const FOF_SILENT = &H4
'''Public Const FOF_NOERRORUI = &H400
'----------------------------------------------------------------------------------------------------

On Error GoTo Error
Dim ShellInfo As SHFILEOPSTRUCT
Dim i As Integer

    With ShellInfo
        For i = 1 To UBound(FullPath())
            If fc_FileExist(FullPath(i)) Then
                .pFrom = .pFrom & FullPath(i) & vbNullChar
            End If
        Next i
        .pFrom = .pFrom & vbNullChar
        .fFlags = FOF_ALLOWUNDO
        .fFlags = .fFlags Or FOF_NOCONFIRMATION
        '.fFlags = .fFlags Or FOF_NOERRORUI
        '.fFlags = .fFlags Or FOF_SILENT
        '.hwnd = Screen.ActiveForm.hwnd          ' ein Handle der aktiven Form
                                                '(muss nicht unbedingt mit übergeben werden)
        .wFunc = FO_DELETE
    End With
    SHFileOperation ShellInfo
    
    fc_DelFiles = True
    Exit Function
Error:
    fc_DelFiles = False
    
End Function

Public Function fc_DelFile(FullPath As String) As Boolean
On Error GoTo Error
Dim ShellInfo As SHFILEOPSTRUCT
Dim i As Integer
    With ShellInfo
        .pFrom = .pFrom & FullPath & vbNullChar
        .fFlags = FOF_ALLOWUNDO
        .fFlags = .fFlags Or FOF_NOCONFIRMATION
        .fFlags = .fFlags Or FOF_NOERRORUI
        .wFunc = FO_DELETE
    End With
    SHFileOperation ShellInfo
    fc_DelFile = True
    Exit Function
Error:
    fc_DelFile = False
End Function

Public Function fc_FileCopy(ByVal srcFullPath As String, ByVal dstFullPath As String, _
                                    Optional ByVal ReplaceIfExist As Boolean = False, _
                                    Optional ByVal Silent As Boolean = True) As Boolean
Dim FileStructur As SHFILEOPSTRUCT
Dim FLAG As Integer
  ' ReplaceIfExist: True, wenn ohne Warnung überschrieben werden soll (Entspricht -y beim DOS copy BEFEHL)

  FLAG = 0
  If InStr(srcFullPath, vbNullChar + vbNullChar) > 0 Then FLAG = FLAG + FOF_MULTIDESTFILES
  If InStr(srcFullPath, "*") > 0 Then FLAG = FLAG + FOF_FILESONLY
  If ReplaceIfExist = True Then FLAG = FLAG + FOF_RENAMEONCOLLISION
  With FileStructur
    .wFunc = FO_COPY
    .pFrom = AppendNullChars(srcFullPath)
    .pTo = dstFullPath
    .fFlags = FLAG Or IIf(Silent, FOF_SILENT, 0&)
  End With
  fc_FileCopy = (SHFileOperation(FileStructur) = 0)
End Function

Public Function fc_DeleteFiles(ByVal FullPath As String, Optional ByVal Recycle As Boolean = False, _
                                Optional ByVal ShowDlg As Boolean = False, _
                                Optional ByVal Silent As Boolean = True) As Boolean
  ' Recycle: True, wenn in Papierkorb gelöscht
  ' ShowDlg: True, wenn zusätzlich Löschabfrage erfolgen soll
  Dim FileStructur As SHFILEOPSTRUCT
  Dim Flags As Long
  
  Flags = 0
  If Recycle Then Flags = FOF_ALLOWUNDO
  If Not ShowDlg Then Flags = Flags Or FOF_NOCONFIRMATION
  
  With FileStructur
    .wFunc = FO_DELETE
    .pFrom = AppendNullChars(FullPath)
    .fFlags = Flags Or IIf(Silent, FOF_SILENT, 0&)
  End With

  fc_DeleteFiles = (SHFileOperation(FileStructur) = 0)
End Function

Public Function fc_MoveFolder(ByVal srcFullPath As String, ByVal dstFullPath As String, Optional ByVal Silent As Boolean = True) As Boolean
Dim FileStructur As SHFILEOPSTRUCT
    srcFullPath = fc_Korrect_Path(fc_Path(srcFullPath))
    dstFullPath = fc_Korrect_Path(fc_Path(dstFullPath))
    With FileStructur
        .wFunc = FO_MOVE
        .pFrom = AppendNullChars(srcFullPath)
        .pTo = dstFullPath
        .fFlags = FOF_RENAMEONCOLLISION Or IIf(Silent, FOF_SILENT, 0&)
    End With
    fc_MoveFolder = (SHFileOperation(FileStructur) = 0)
End Function

Public Function fc_MoveFiles(ByVal srcFullPath As String, ByVal dstFullPath As String, Optional ByVal Silent As Boolean = True) As Boolean
Dim FileStructur As SHFILEOPSTRUCT
  With FileStructur
    .wFunc = FO_MOVE
    .pFrom = AppendNullChars(srcFullPath)
    .pTo = dstFullPath
    .fFlags = FOF_RENAMEONCOLLISION Or IIf(Silent, FOF_SILENT, 0&)
  End With
  fc_MoveFiles = (SHFileOperation(FileStructur) = 0)
End Function

Public Function fc_RenameFolder(ByVal srcFullPath As String, ByVal dstFullPath As String, Optional ByVal Silent As Boolean = True) As Boolean
Dim FileStructur As SHFILEOPSTRUCT
    srcFullPath = fc_Korrect_Path(fc_Path(srcFullPath))
    dstFullPath = fc_Korrect_Path(fc_Path(dstFullPath))
    With FileStructur
        .wFunc = FO_RENAME
        .pFrom = AppendNullChars(srcFullPath)
        .pTo = dstFullPath
        .fFlags = FOF_RENAMEONCOLLISION Or IIf(Silent, FOF_SILENT, 0&)
    End With
    
    fc_RenameFolder = (SHFileOperation(FileStructur) = 0)
End Function
Public Function fc_RenameFiles(ByVal srcFullPath As String, ByVal dstFullPath As String, Optional ByVal Silent As Boolean = True) As Boolean
Dim FileStructur As SHFILEOPSTRUCT
 
    With FileStructur
        .wFunc = FO_RENAME
        .pFrom = AppendNullChars(srcFullPath)
        .pTo = dstFullPath
        .fFlags = FOF_RENAMEONCOLLISION Or IIf(Silent, FOF_SILENT, 0&)
    End With
    
    fc_RenameFiles = (SHFileOperation(FileStructur) = 0)
End Function

Public Function fc_FilesFromArray(ByRef Liste() As String) As String
' Alle Dateinamen eines Array-Datenfeldes hintereinander - durch vbNullChar getrennt - zusammenfassen
Dim i As Long
Dim Temp As String
    For i = 0 To UBound(Liste)
        If fc_FileExist(Liste(i)) Then
            Temp = Temp + Liste(i) + vbNullChar     'Datei-Eintrag mit CHR(0) abschließen
        End If
    Next
    fc_FilesFromArray = Temp + vbNullChar           'Notwendig: Abschließendes CHR(0)
End Function

' Alle Angaben müssen mit vbNullChar+vbNullChar abgeschlossen werden. Hier wird's noch mal geprüft
Private Function AppendNullChars(ByVal str As String) As String
  If Right(str, 2) <> vbNullChar + vbNullChar Then
    If Right(str, 1) <> vbNullChar Then
      str = str + vbNullChar + vbNullChar
    Else
      str = str + vbNullChar
    End If
  End If
  AppendNullChars = str
End Function


Public Function fc_FolderBrowser(Optional ByVal Title As String, Optional ByVal Path As String) As String
Dim res As Long
Dim BI As BROWSEINFO
    If Not Len(Title) > 0 Then Title = "Bitte wählen Sie ein Verzeichnis aus!"
    If Not Len(Path) > 0 Then Path = "C:\"
    With BI
        .hOwner = GetActiveWindow()
        .lpszTitle = Title
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN
        
        strStartPath = Path
        .lpfn = FnPtrToLong(AddressOf cb_SetStartPath)

    End With
    
    res = SHBrowseForFolder(BI)
    If res Then
        Path = String$(MAX_PATH, vbNullChar)
        SHGetPathFromIDList res, Path
        Call CoTaskMemFree(res)
        Path = Replace(Trim(Path), vbNullChar, "")
    End If
    fc_FolderBrowser = Path
End Function

Private Function FnPtrToLong(ByVal lngFnPtr As Long) As Long
    FnPtrToLong = lngFnPtr
End Function

Private Function cb_SetStartPath(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    'Callback Funktion zum setzen des Startpath für den SHBrowseForFolder API-Dialog in Function fc_FolderBrowser
    If uMsg = BFFM_INITIALIZED Then
        If Len(strStartPath) > 1 Then
            SendMessage hwnd, BFFM_SETSELECTION, 1, strStartPath
        End If
    End If
End Function

'Systempfade ermitteln-------------------------------------------------------------------------------
Public Function GetProgramsPath() As String
Dim ProgramsPath As String, tmpPath As String
    ProgramsPath = GetSpecialFolder(CSIDL_PROGRAM_FILES)
    If Len(Trim(ProgramsPath)) <= 1 Then
        tmpPath = GetSystemDir
        tmpPath = fc_Drive(tmpPath)
        If fc_Folder_Exists(tmpPath & "\Programme") Then
            ProgramsPath = tmpPath & "\Programme"
        ElseIf fc_Folder_Exists(tmpPath & "\Program Files") Then
            ProgramsPath = tmpPath & "\Program Files"
        ElseIf fc_Folder_Exists(tmpPath & "\ProgramFiles") Then
            ProgramsPath = tmpPath & "\ProgramFiles"
        End If
    End If
    GetProgramsPath = ProgramsPath
End Function

Public Function GetWinSysPath() As String
Dim WinSysPath As String
    WinSysPath = GetSpecialFolder(CSIDL_SYSTEM)
    If Len(Trim(WinSysPath)) <= 1 Then
       WinSysPath = GetSystemDir
    End If
    GetWinSysPath = WinSysPath
End Function

Public Function GetWinPath() As String
Dim WinPath As String
    WinPath = GetSpecialFolder(CSIDL_WINDOWS)
    If Len(Trim(WinPath)) <= 1 Then
       WinPath = GetWindowsDir
    End If
    GetWinPath = WinPath
End Function

'Windows-Verzeichnis ermitteln
Public Function GetWindowsDir() As String
Dim Temp As String
Dim lResult As Integer
    Temp = Space$(256)
    lResult = GetWindowsDirectory(Temp, Len(Temp))
    Temp = Left$(Temp, lResult)
    GetWindowsDir = Temp
End Function

'Windows-System-Verzeichnis ermitteln
Public Function GetSystemDir() As String
Dim Temp As String
Dim lResult As Long
    Temp = Space$(256)
    lResult = GetSystemDirectory(Temp, Len(Temp))
    Temp = Left$(Temp, lResult)
    GetSystemDir = Temp
End Function

Public Function GetSpecialFolder(CSIDL As SpecialFolderIDs) As String
Dim lResult As Long
Dim IDL As ITEMIDLIST
Dim sPath As String
    Select Case CSIDL
        Case myidNone: GetSpecialFolder = "kein"
        Case myidAppPath: GetSpecialFolder = fc_AppPath
        Case Else
            lResult = SHGetSpecialFolderLocation(100, CSIDL, IDL)
            If lResult = 0 Then
                sPath = Space$(512)
                lResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
                GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
            End If
    End Select
End Function

'-------------------------------------------------------------------------------Systempfade ermitteln

'ActiveX-Komponenten registrieren--------------------------------------------------------------------

Public Function RegServeDLL(ByVal Path As String, Optional ByVal Unregister As Boolean = False) As Boolean
On Error GoTo ErrOut
Dim insthLib As Long
Dim lpLibAdr As Long
Dim hThd As Long
Dim lpExCode As Long
Dim procName As String
Dim result As Long
Dim okFlag As Boolean

    RegServeDLL = False
    
    'DLL in den Speicher laden
    insthLib = LoadLibrary(Path)
    
    'Aktion wählen
    If insthLib Then
        If Unregister Then
            procName = "DllUnregisterServer"
        Else
            procName = "DllRegisterServer"
        End If
        
        'Adresse der DLL im Speicher
        lpLibAdr = GetProcAddress(insthLib, procName)
        If lpLibAdr <> 0 Then
            
            'Aktion starten
            hThd = CreateThread(ByVal 0, 0, ByVal lpLibAdr, _
                                ByVal 0&, 0&, 0&)
            If hThd Then
                'Maximal 5 sec warten
                result = WaitForSingleObject(hThd, 5000)
                If result = STATUS_WAIT_0 Then
                    'Vorgang erfolgreich in 5 sec beendet
                    Call CloseHandle(hThd)
                    okFlag = True
                Else
                    '5 sec überschritten -> Thread schließen
                    Call GetExitCodeThread(hThd, lpExCode)
                    Call ExitThread(lpExCode)
                    Call CloseHandle(hThd)
                End If
            End If
        End If
        'Speicher wieder freigeben
        Call FreeLibrary(insthLib)
    End If
    
    If Not okFlag Then
        RegServeDLL = False
    Else
        RegServeDLL = True
    End If
Exit Function
ErrOut:
    RegServeDLL = False
End Function

Private Function RegServeEXE(ByVal Path As String, Optional ByVal Unregister As Boolean = False) As Boolean
On Error GoTo ErrOut
    'ActiveX-Exen besitzen von sich aus eine Methode, sich zu registrieren
    If Unregister Then
        Shell Path & " /unregserver"
    Else
        Shell Path & " /regserver"
    End If
    RegServeEXE = True
    Exit Function
ErrOut:
    RegServeEXE = False
End Function

'--------------------------------------------------------------------ActiveX-Komponenten registrieren

'Datei-Versions-Info-######################################################################################
Public Function GetVerInfo(ByVal FilePath As String, VerInfo As FILEVERSIONINFO) As Boolean
On Error GoTo ErrOut
    'alte Version der Funktion
    VerInfo = fc_GetVerInfo(FilePath)
    GetVerInfo = True
Exit Function
ErrOut:
    GetVerInfo = False
End Function


Public Function fc_GetVerInfo(ByVal FilePath As String) As FILEVERSIONINFO
  Dim StructVer As String
  Dim FileVer As String
  Dim ProductVer As String
  Dim VerFlags As String
  Dim typ As String
  Dim SubTyp As String
  Dim OS As String
  Dim Buff() As Byte
  Dim len_ As Long
  Dim BuffL As Long
  Dim pointer As Long
  Dim Version As VS_FIXEDFILEINFO
  Dim FileVerInfo As FILEVERSIONINFO
    len_ = GetFileVersionInfoSize(FilePath, 0&)
    If len_ >= 1 Then
        ReDim Buff(len_)
        Call GetFileVersionInfo(FilePath, 0&, len_, Buff(0))
        Call VerQueryValue(Buff(0), "\", pointer, BuffL)
        Call MoveMemory(Version, pointer, Len(Version))

        With Version
            StructVer = Format$(.dwStrucVersionh) & "." & Format$(.dwStrucVersionl)
            
            FileVer = Format$(.dwFileVersionMSh) & "." & Format$(.dwFileVersionMSl) & "." & _
                                        Format$(.dwFileVersionLSh) & "." & _
                                        Format$(.dwFileVersionLSl)
            
            ProductVer = Format$(.dwProductVersionMSh) & "." & _
                                        Format$(.dwProductVersionMSl) & "." & _
                                        Format$(.dwProductVersionLSh) & "." & _
                                        Format$(.dwProductVersionLSl)
            
            If .dwFileFlags And VS_FF_DEBUG Then VerFlags = "Debug "
            If .dwFileFlags And VS_FF_PRERELEASE Then VerFlags = VerFlags & "PreRel "
            If .dwFileFlags And VS_FF_PATCHED Then VerFlags = VerFlags & "Patched "
            If .dwFileFlags And VS_FF_PRIVATEBUILD Then VerFlags = VerFlags & "Private "
            If .dwFileFlags And VS_FF_INFOINFERRED Then VerFlags = VerFlags & "Info "
            If .dwFileFlags And VS_FF_SPECIALBUILD Then VerFlags = VerFlags & "Special "
            If .dwFileFlags And VFT2_UNKNOWN Then VerFlags = VerFlags & "Unknown "
        End With

        Select Case Version.dwFileOS
            Case VOS_DOS_WINDOWS16: OS = "DOS-Win16"
            Case VOS_DOS_WINDOWS32: OS = "DOS-Win32"
            Case VOS_OS216_PM16:    OS = "OS/2-16 PM-16"
            Case VOS_OS232_PM32:    OS = "OS/2-16 PM-32"
            Case VOS_NT_WINDOWS32:  OS = "NT-Win32"
            Case Else:              OS = "Unbekannt"
        End Select
        
        Select Case Version.dwFileType
            Case VFT_APP:               typ = "Anwendung"
            Case VFT_DLL:               typ = "Dynamic Link Library (DLL)"
            Case VFT_DRV:               typ = "Geräte Treiber"
                Select Case Version.dwFileSubtype
                    Case VFT2_DRV_PRINTER:     SubTyp = "Printer drv"
                    Case VFT2_DRV_KEYBOARD:    SubTyp = "Keyboard drv"
                    Case VFT2_DRV_LANGUAGE:    SubTyp = "Language drv"
                    Case VFT2_DRV_DISPLAY:     SubTyp = "Display drv"
                    Case VFT2_DRV_MOUSE:       SubTyp = "Mouse drv"
                    Case VFT2_DRV_NETWORK:     SubTyp = "Network drv"
                    Case VFT2_DRV_SYSTEM:      SubTyp = "System drv"
                    Case VFT2_DRV_INSTALLABLE: SubTyp = "Installable"
                    Case VFT2_DRV_SOUND:       SubTyp = "Sound drv"
                    Case VFT2_DRV_COMM:        SubTyp = "Comm drv"
                    Case VFT2_UNKNOWN:         SubTyp = "Unknown"
                End Select
            Case VFT_FONT:              typ = "Schriftart"
                Select Case Version.dwFileSubtype
                    Case VFT2_FONT_RASTER:     SubTyp = "Raster Font"
                    Case VFT2_FONT_VECTOR:     SubTyp = "Vector Font"
                    Case VFT2_FONT_TRUETYPE:   SubTyp = "TrueType Font"
                End Select
            Case VFT_VXD:               typ = "Virtueller Geräte Treiber"
            Case VFT_STATIC_LIB:        typ = "Static Library"
            Case VFT_UNKNOWN:           typ = "Unbekannt"
            Case Else:                  typ = "Unbekannt"
        End Select
    End If

    With FileVerInfo
        .StructVer = StructVer
        .FileVer = FileVer
        .ProductVer = ProductVer
        .VerFlags = VerFlags
        .OS = OS
        .typ = typ
        .SubTyp = SubTyp
    End With
    fc_GetVerInfo = FileVerInfo
End Function

Public Function fc_RegisterApllicationFiles(ByVal AppName As String, ByVal FileExtension As String, ByVal FileDescryption As String, ByVal Icon As Byte)
Dim res As Long, hKey As Long

    '### Generiert den neuen Eintrag
    res = RegCreateKey&(HKEY_CLASSES_ROOT, FileDescryption, hKey&)
    res = RegSetValue&(hKey&, "", REG_SZ, AppName, 0&)
  
    '### Generiert die Assoziation mit der Extension
    res = RegCreateKey&(HKEY_CLASSES_ROOT, "." & FileExtension, hKey&)
    res = RegSetValue&(hKey&, "", REG_SZ, FileDescryption, 0&)

    '### Setzt den ausführenden Pfad für die Anwendung
    res = RegCreateKey&(HKEY_CLASSES_ROOT, FileDescryption, hKey&)
    res = RegSetValue&(hKey&, "shell\open\command", REG_SZ, _
            fc_AppPath & App.EXEName & ".exe %0 /open/", MAX_PATH)
    
    '### Setzt das Symbol in Assoziation mit der Extension
    res = RegSetValue&(hKey, "DefaultIcon", REG_SZ, _
            fc_AppPath & App.EXEName & ".exe ," & Icon, MAX_PATH)
End Function

Public Function fc_GetFiles(ByRef Files() As String, ByVal SeekPath As String, _
                                Optional Extensions As String, _
                                Optional ByVal ExcludeExtensions As Boolean = False) As Boolean
Dim strExt() As String
Dim i As Integer
Dim j As Integer
Dim CurFile As String
Dim Count As Long
On Error GoTo ErrOut

    Count = -1
    If Len(Extensions) > 0 Then
        strExt = Split(Extensions, ",")
    Else
        ReDim strExt(0)
        strExt(0) = "*"
    End If
    SeekPath = fc_Korrect_Path(SeekPath)
    If Not IsArray(strExt()) Then GoTo ErrOut
    If Not ExcludeExtensions Then
        For i = LBound(strExt) To UBound(strExt)
Next_i:
            CurFile = Dir(SeekPath & "*." & strExt(i), vbHidden Or vbSystem Or vbNormal Or vbArchive): DoEvents: Sleep 1
            If Len(CurFile) = 0 Then
                i = i + 1
                If i <= UBound(strExt) Then
                    GoTo Next_i
                Else
                    Exit For
                End If
            End If
            Do While Len(CurFile) > 1
                On Error Resume Next
                DoEvents: Sleep 1
                If (GetAttr(SeekPath & CurFile) And vbDirectory) <> vbDirectory Then
                    On Error GoTo ErrOut
                    Count = Count + 1
                    ReDim Preserve Files(Count)
                    Files(Count) = CurFile
                End If
                CurFile = Dir
            Loop
        Next i
    Else
        CurFile = Dir(SeekPath & "*.*", vbHidden Or vbSystem Or vbNormal Or vbArchive): DoEvents: Sleep 1
        Do While CurFile <> ""
            For j = LBound(strExt) To UBound(strExt)
                If Trim(LCase(Mid(fc_FileExtension(CurFile), 2))) = Trim(LCase(strExt(j))) Then
                    CurFile = "_"
                    Exit For
                End If
            Next j
            If Len(CurFile) > 1 Then
                On Error Resume Next
                DoEvents: Sleep 1
                If (GetAttr(SeekPath & CurFile) And vbDirectory) <> vbDirectory Then
                    If Err.Number > 0 Then
                        Err.Clear
                        On Error GoTo ErrOut
                        Count = Count + 1
                        ReDim Preserve Files(Count)
                        Files(Count) = CurFile
                    Else
                        On Error GoTo ErrOut
                        Count = Count + 1
                        ReDim Preserve Files(Count)
                        Files(Count) = CurFile
                    End If
                End If
            End If
            CurFile = Dir
        Loop
    End If
    If Count < 0 Then GoTo ErrOut
    fc_GetFiles = True
Exit Function
ErrOut:
    ReDim Files(0)
    fc_GetFiles = False
End Function


Public Function ReplacePathVariables(ByVal VariablePath As String) As String
Dim tmpPath As String                                                                                   'WindowsXP  Windows98
    tmpPath = VariablePath
    tmpPath = Replace(tmpPath, "%AppPath%", fc_AppPath)                                                 '   Ja          Ja
    tmpPath = Replace(tmpPath, "%AdminTools%", GetSpecialFolder(CSIDL_ADMINTOOLS))                      '   Ja          Nein
    tmpPath = Replace(tmpPath, "%AppDataDir%", GetSpecialFolder(CSIDL_APPDATA))                         '   Ja          Ja
    tmpPath = Replace(tmpPath, "%CDBurnArea%", GetSpecialFolder(CSIDL_CDBURN_AREA))                     '   Ja          Nein
'    tmpPath = Replace(tmpPath, "%CommonAltStartUp%", GetSpecialFolder(CSIDL_COMMON_ALTSTARTUP))        '   Nein        Nein
    tmpPath = Replace(tmpPath, "%CommonAppData%", GetSpecialFolder(CSIDL_COMMON_APPDATA))               '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonAutostart%", GetSpecialFolder(CSIDL_COMMON_STARTUP))             '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonDesktop%", GetSpecialFolder(CSIDL_COMMON_DESKTOPDIRECTORY))      '   Ja          Ja
    tmpPath = Replace(tmpPath, "%CommonDocuments%", GetSpecialFolder(CSIDL_COMMON_DOCUMENTS))           '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonFavorites%", GetSpecialFolder(CSIDL_COMMON_FAVORITES))           '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonMusic%", GetSpecialFolder(CSIDL_COMMON_MUSIC))                   '   Ja          Nein
'    tmpPath = Replace(tmpPath, "%CommonOemLinks%", GetSpecialFolder(CSIDL_COMMON_OEM_LINKS))           '   Nein        Nein
    tmpPath = Replace(tmpPath, "%CommonPictures%", GetSpecialFolder(CSIDL_COMMON_PictureS))             '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonPrograms%", GetSpecialFolder(CSIDL_COMMON_PROGRAMS))             '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonStartMenu%", GetSpecialFolder(CSIDL_COMMON_STARTMENU))           '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonTemplatesDir%", GetSpecialFolder(CSIDL_COMMON_TEMPLATES))        '   Ja          Nein
    tmpPath = Replace(tmpPath, "%CommonVideos%", GetSpecialFolder(CSIDL_COMMON_VIDEO))                  '   Ja          Nein
'    tmpPath = Replace(tmpPath, "%ControlsDir%", GetSpecialFolder(CSIDL_CONTROLS))                      '   Nein        Nein
    tmpPath = Replace(tmpPath, "%CookiesDir%", GetSpecialFolder(CSIDL_COOKIES))                         '   Ja          Ja
    tmpPath = Replace(tmpPath, "%DesktopDir%", GetSpecialFolder(CSIDL_DESKTOP))                         '   Ja          Ja
'    tmpPath = Replace(tmpPath, "%DriversDir%", GetSpecialFolder(CSIDL_DRIVERS))                        '   Nein        Nein
    tmpPath = Replace(tmpPath, "%FontsDir%", GetSpecialFolder(CSIDL_FONTS))                             '   Ja          Ja
    tmpPath = Replace(tmpPath, "%History%", GetSpecialFolder(CSIDL_HISTORY))                            '   Ja          Ja
    tmpPath = Replace(tmpPath, "%InternetFilesDir%", GetSpecialFolder(CSIDL_INTERNET_CACHE))            '   Ja          Ja
    tmpPath = Replace(tmpPath, "%LocalAppData%", GetSpecialFolder(CSIDL_LOCAL_APPDATA))                 '   Ja          Nein
'    tmpPath = Replace(tmpPath, "%MyDocuments%", GetSpecialFolder(CSIDL_MYDOCUMENTS))                   '   Nein        Nein
'    tmpPath = Replace(tmpPath, "%UserAutostart%", GetSpecialFolder(CSIDL_ALTSTARTUP))                  '   Nein        Nein
    tmpPath = Replace(tmpPath, "%UserDesktop%", GetSpecialFolder(CSIDL_DESKTOPDIRECTORY))               '   Ja          Ja
    tmpPath = Replace(tmpPath, "%UserFavorites%", GetSpecialFolder(CSIDL_FAVORITES))                    '   Ja          Ja
    tmpPath = Replace(tmpPath, "%UserMyDocuments%", GetSpecialFolder(CSIDL_PERSONAL))                   '   Ja          Ja
    tmpPath = Replace(tmpPath, "%UserMyMusic%", GetSpecialFolder(CSIDL_MYMUSIC))                        '   Ja          Nein
    tmpPath = Replace(tmpPath, "%UserMyPictures%", GetSpecialFolder(CSIDL_MYPictureS))                  '   Ja          Nein
    tmpPath = Replace(tmpPath, "%UserMyVideos%", GetSpecialFolder(CSIDL_MYVIDEO))                       '   Ja          Nein
    tmpPath = Replace(tmpPath, "%UserProfile%", GetSpecialFolder(CSIDL_PROFILE))                        '   Ja          Nein
    tmpPath = Replace(tmpPath, "%UserPrograms%", GetSpecialFolder(CSIDL_PROGRAMS))                      '   Ja          Ja
    tmpPath = Replace(tmpPath, "%UserStartMenu%", GetSpecialFolder(CSIDL_STARTMENU))                    '   Ja          Ja
    tmpPath = Replace(tmpPath, "%UserTemplatesDir%", GetSpecialFolder(CSIDL_TEMPLATES))                 '   Ja          Ja
    tmpPath = Replace(tmpPath, "%ProgramsCommonDir%", GetSpecialFolder(CSIDL_PROGRAM_FILES_COMMON))     '   Ja          Nein
    tmpPath = Replace(tmpPath, "%ProgramsDir%", GetProgramsPath)                                        '   Ja          Nein
    tmpPath = Replace(tmpPath, "%RescentFilesDir%", GetSpecialFolder(CSIDL_RECENT))                     '   Ja          Ja
    tmpPath = Replace(tmpPath, "%Resources%", GetSpecialFolder(CSIDL_RESOURCES))                        '   Ja          Nein
'    tmpPath = Replace(tmpPath, "%ResourcesLocalized%", GetSpecialFolder(CSIDL_RESOURCES_LOCALIZED))    '   Nein        Nein
    tmpPath = Replace(tmpPath, "%SendTo%", GetSpecialFolder(CSIDL_SENDTO))                              '   Ja          Ja
'    tmpPath = Replace(tmpPath, "%WindowsSystemDir%", GetSpecialFolder(CSIDL_SYSTEM))                   '   Ja          Nein
    tmpPath = Replace(tmpPath, "%WindowsSystemDir%", GetWinSysPath)
'    tmpPath = Replace(tmpPath, "%WindowsDir%", GetSpecialFolder(CSIDL_WINDOWS))                        '   Ja          Nein
    tmpPath = Replace(tmpPath, "%WindowsDir%", GetWinPath)
    
    If Len(tmpPath) = 0 Or Left(tmpPath, 1) = "\" Then
        Dim p1 As Integer
        Dim p2 As Integer
        p1 = InStr(1, VariablePath, "%")
        p2 = InStr(p1 + 1, VariablePath, "%")
        tmpPath = Replace(VariablePath, Mid(VariablePath, p1, p2 - p1 + 1), "")
        If Left(tmpPath, 1) = "\" Then tmpPath = Mid(tmpPath, 2)
        tmpPath = fc_AppPath & tmpPath
    End If

    ReplacePathVariables = tmpPath
End Function

Public Function GetPathVariables(ByVal CSIDL As SpecialFolderIDs) As String
Dim tmpPath As String
    Select Case CSIDL
        Case CSIDL_ADMINTOOLS
            tmpPath = "%AdminTools%"
        Case CSIDL_APPDATA
            tmpPath = "%AppDataDir%"
        Case CSIDL_CDBURN_AREA
            tmpPath = "%CDBurnArea%"
'        Case CSIDL_COMMON_ALTSTARTUP
'            tmpPath = "%CommonAltStartUp%"
        Case CSIDL_COMMON_APPDATA
            tmpPath = "%CommonAppData%"
        Case CSIDL_COMMON_STARTUP
            tmpPath = "%CommonAutostart%"
        Case CSIDL_COMMON_DESKTOPDIRECTORY
            tmpPath = "%CommonDesktop%"
        Case CSIDL_COMMON_DOCUMENTS
            tmpPath = "%CommonDocuments%"
        Case CSIDL_COMMON_FAVORITES
            tmpPath = "%CommonFavorites%"
        Case CSIDL_COMMON_MUSIC
            tmpPath = "%CommonMusic%"
'        Case CSIDL_COMMON_OEM_LINKS
'            tmpPath = "%CommonOemLinks%"
        Case CSIDL_COMMON_PictureS
            tmpPath = "%CommonPictures%"
        Case CSIDL_COMMON_PROGRAMS
            tmpPath = "%CommonPrograms%"
        Case CSIDL_COMMON_STARTMENU
            tmpPath = "%CommonStartMenu%"
        Case CSIDL_COMMON_TEMPLATES
            tmpPath = "%CommonTemplatesDir%"
        Case CSIDL_COMMON_VIDEO
            tmpPath = "%CommonVideos%"
'        Case CSIDL_CONTROLS
'            tmpPath = "%ControlsDir%"
        Case CSIDL_COOKIES
            tmpPath = "%CookiesDir%"
        Case CSIDL_DESKTOP
            tmpPath = "%DesktopDir%"
'        Case CSIDL_DRIVERS
'            tmpPath = "%DriversDir%"
        Case CSIDL_FONTS
            tmpPath = "%FontsDir%"
        Case CSIDL_HISTORY
            tmpPath = "%History%"
        Case CSIDL_INTERNET_CACHE
            tmpPath = "%InternetFilesDir%"
        Case CSIDL_LOCAL_APPDATA
            tmpPath = "%LocalAppData%"
'        Case CSIDL_MYDOCUMENTS
'            tmpPath = "%MyDocuments%"
'        Case CSIDL_ALTSTARTUP
'            tmpPath = "%UserAutostart%"
        Case CSIDL_DESKTOPDIRECTORY
            tmpPath = "%UserDesktop%"
        Case CSIDL_FAVORITES
            tmpPath = "%UserFavorites%"
        Case CSIDL_PERSONAL
            tmpPath = "%UserMyDocuments%"
        Case CSIDL_MYMUSIC
            tmpPath = "%UserMyMusic%"
        Case CSIDL_MYPictureS
            tmpPath = "%UserMyPictures%"
        Case CSIDL_MYVIDEO
            tmpPath = "%UserMyVideos%"
        Case CSIDL_PROFILE
            tmpPath = "%UserProfile%"
        Case CSIDL_PROGRAMS
            tmpPath = "%UserPrograms%"
        Case CSIDL_STARTMENU
            tmpPath = "%UserStartMenu%"
        Case CSIDL_TEMPLATES
            tmpPath = "%UserTemplatesDir%"
        Case CSIDL_PROGRAM_FILES_COMMON
            tmpPath = "%ProgramsCommonDir%"
        Case CSIDL_PROGRAM_FILES
            tmpPath = "%ProgramsDir%"
        Case CSIDL_RECENT
            tmpPath = "%RescentFilesDir%"
        Case CSIDL_RESOURCES
            tmpPath = "%Resources%"
'        Case CSIDL_RESOURCES_LOCALIZED
'            tmpPath = "%ResourcesLocalized%"
        Case CSIDL_SENDTO
            tmpPath = "%SendTo%"
        Case CSIDL_SYSTEM
            tmpPath = "%WindowsSystemDir%"
            'tmpPath = "%WindowsSystemDir%", GetWinSysPath)
        Case CSIDL_WINDOWS
            tmpPath = "%WindowsDir%"
            'tmpPath = "%WindowsDir%", GetWinPath)
    End Select
    GetPathVariables = tmpPath
End Function

Public Sub FillPathVarsCmb(ByRef cmb As ComboBox)
    cmb.Clear
    cmb.AddItem "kein":                 cmb.ItemData(cmb.ListCount - 1) = -1
    cmb.AddItem "AppPath":              cmb.ItemData(cmb.ListCount - 1) = -2
    cmb.AddItem "AdminTools":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_ADMINTOOLS
    cmb.AddItem "AppDataDir":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_APPDATA
    cmb.AddItem "CDBurnArea":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_CDBURN_AREA
'    cmb.AddItem "CommonAltStartUp":     cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_ALTSTARTUP
    cmb.AddItem "CommonAppData":        cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_APPDATA
    cmb.AddItem "CommonAutostart":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_STARTUP
    cmb.AddItem "CommonDesktop":        cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_DESKTOPDIRECTORY
    cmb.AddItem "CommonDocuments":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_DOCUMENTS
    cmb.AddItem "CommonFavorites":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_FAVORITES
    cmb.AddItem "CommonMusic":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_MUSIC
'    cmb.AddItem "CommonOemLinks":       cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_OEM_LINKS
    cmb.AddItem "CommonPictures":       cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_PictureS
    cmb.AddItem "CommonPrograms":       cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_PROGRAMS
    cmb.AddItem "CommonStartMenu":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_STARTMENU
    cmb.AddItem "CommonTemplatesDir":   cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_TEMPLATES
    cmb.AddItem "CommonVideos":         cmb.ItemData(cmb.ListCount - 1) = CSIDL_COMMON_VIDEO
'    cmb.AddItem "ControlsDir":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_CONTROLS
    cmb.AddItem "CookiesDir":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_COOKIES
    cmb.AddItem "DesktopDir":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_DESKTOP
'    cmb.AddItem "DriversDir":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_DRIVERS
    cmb.AddItem "FontsDir":             cmb.ItemData(cmb.ListCount - 1) = CSIDL_FONTS
    cmb.AddItem "History":              cmb.ItemData(cmb.ListCount - 1) = CSIDL_HISTORY
    cmb.AddItem "InternetFilesDir":     cmb.ItemData(cmb.ListCount - 1) = CSIDL_INTERNET_CACHE
    cmb.AddItem "LocalAppData":         cmb.ItemData(cmb.ListCount - 1) = CSIDL_LOCAL_APPDATA
'    cmb.AddItem "MyDocuments":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_MYDOCUMENTS
    cmb.AddItem "ProgramsCommonDir":    cmb.ItemData(cmb.ListCount - 1) = CSIDL_PROGRAM_FILES_COMMON
    cmb.AddItem "ProgramsDir":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_PROGRAM_FILES
    cmb.AddItem "RescentFilesDir":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_RECENT
    cmb.AddItem "Resources":            cmb.ItemData(cmb.ListCount - 1) = CSIDL_RESOURCES
'    cmb.AddItem "ResourcesLocalized":  cmb.ItemData(cmb.ListCount - 1) = CSIDL_RESOURCES_LOCALIZED
    cmb.AddItem "SendTo":               cmb.ItemData(cmb.ListCount - 1) = CSIDL_SENDTO
'    cmb.AddItem "UserAutostart":        cmb.ItemData(cmb.ListCount - 1) = CSIDL_ALTSTARTUP
    cmb.AddItem "UserDesktop":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_DESKTOPDIRECTORY
    cmb.AddItem "UserFavorites":        cmb.ItemData(cmb.ListCount - 1) = CSIDL_FAVORITES
    cmb.AddItem "UserMyDocuments":      cmb.ItemData(cmb.ListCount - 1) = CSIDL_PERSONAL
    cmb.AddItem "UserMyMusic":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_MYMUSIC
    cmb.AddItem "UserMyPictures":       cmb.ItemData(cmb.ListCount - 1) = CSIDL_MYPictureS
    cmb.AddItem "UserMyVideos":         cmb.ItemData(cmb.ListCount - 1) = CSIDL_MYVIDEO
    cmb.AddItem "UserProfile":          cmb.ItemData(cmb.ListCount - 1) = CSIDL_PROFILE
    cmb.AddItem "UserPrograms":         cmb.ItemData(cmb.ListCount - 1) = CSIDL_PROGRAMS
    cmb.AddItem "UserStartMenu":        cmb.ItemData(cmb.ListCount - 1) = CSIDL_STARTMENU
    cmb.AddItem "UserTemplatesDir":     cmb.ItemData(cmb.ListCount - 1) = CSIDL_TEMPLATES
    cmb.AddItem "WindowsDir":           cmb.ItemData(cmb.ListCount - 1) = CSIDL_WINDOWS
    cmb.AddItem "WindowsSystemDir":     cmb.ItemData(cmb.ListCount - 1) = CSIDL_SYSTEM
End Sub

'#############################################################################################################################

Public Function GetFileAttributesName(ByVal vbAttribute As VbFileAttribute) As String
Dim Att As String
    
    If (vbAttribute And vbAlias) = vbAlias Then Att = "Alias"
    If (vbAttribute And vbArchive) = vbArchive Then Att = Att & " Archive"
    If (vbAttribute And vbDirectory) = vbDirectory Then Att = Att & " Directory"
    If (vbAttribute And vbHidden) = vbHidden Then Att = Att & " Hidden"
    If (vbAttribute And vbReadOnly) = vbReadOnly Then Att = Att & " ReadOnly"
    If (vbAttribute And vbSystem) = vbSystem Then Att = Att & " System"
    If (vbAttribute And vbVolume) = vbVolume Then Att = Att & " Volume"
    If vbAttribute = vbNormal Then Att = "Normal"
    Att = Trim(Att)
    Att = Replace(Att, " ", ", ")
    GetFileAttributesName = Att
End Function
