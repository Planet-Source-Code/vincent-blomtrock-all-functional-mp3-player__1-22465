VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vincent's MP3-player"
   ClientHeight    =   8250
   ClientLeft      =   990
   ClientTop       =   570
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   7455
   Begin VB.TextBox Text2 
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   42
      Top             =   4800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browser"
      Height          =   375
      Left            =   4080
      TabIndex        =   41
      Top             =   3480
      Width           =   877
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2895
      Left            =   120
      TabIndex        =   40
      Top             =   5040
      Visible         =   0   'False
      Width           =   7215
      ExtentX         =   12726
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00000000&
      Caption         =   "Green"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6360
      TabIndex        =   39
      Top             =   4440
      Width           =   855
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00000000&
      Caption         =   "Black"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6360
      TabIndex        =   38
      Top             =   3720
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00000000&
      Caption         =   "Blue"
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   6360
      TabIndex        =   37
      Top             =   4080
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00000000&
      Caption         =   "Green"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Blue"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   2640
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Yellow"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   25
      Top             =   2280
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   3960
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton OptNON 
      BackColor       =   &H00000000&
      Caption         =   "None"
      ForeColor       =   &H0000FF00&
      Height          =   200
      Left            =   2640
      TabIndex        =   23
      Top             =   3960
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton OptRes 
      BackColor       =   &H00000000&
      Caption         =   "Resume"
      ForeColor       =   &H0000FF00&
      Height          =   200
      Left            =   1320
      TabIndex        =   22
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton OptRND 
      BackColor       =   &H00000000&
      Caption         =   "Random"
      ForeColor       =   &H0000FF00&
      Height          =   200
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdOnTop 
      Caption         =   "On Top"
      Height          =   315
      Left            =   5040
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton CmdSysinfo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "System Info"
      Height          =   315
      Left            =   5040
      TabIndex        =   18
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Soundsystem"
      ForeColor       =   &H0000FF00&
      Height          =   1875
      Left            =   5160
      TabIndex        =   11
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton CmdMute 
         Caption         =   "&Mute"
         Height          =   255
         Left            =   420
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   1080
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Max             =   2500
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Min             =   -5000
         Max             =   5000
         TickStyle       =   3
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Balance"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label 
         BackColor       =   &H00000000&
         Caption         =   "Volume    "
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "0%"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   780
         Width           =   435
      End
   End
   Begin VB.TextBox TxtTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   675
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "A&bout"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Cmdvol 
      Caption         =   "&Volume"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Minimi&ze"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   3480
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   1860
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4935
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Keep On Top"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3600
      TabIndex        =   19
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "Terminate Browser"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   7920
      Width           =   7215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   7320
      X2              =   7320
      Y1              =   5040
      Y2              =   7800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   7320
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   7320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   120
      X2              =   120
      Y1              =   5040
      Y2              =   7800
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Backcolor"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   36
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
      Height          =   1215
      Left            =   6120
      TabIndex        =   35
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Fonts"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   34
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4080
      TabIndex        =   33
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Save List"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   32
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Remove All"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   31
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Remove"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Load Playlist"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6120
      TabIndex        =   28
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Add MP3"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   1920
      Width           =   975
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu addmenu 
      Caption         =   "Addmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuaddfile 
         Caption         =   "File"
         Index           =   1
      End
      Begin VB.Menu mnuAdddir 
         Caption         =   "Directory"
         Index           =   2
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuplayah 
      Caption         =   "MP3-player"
      Begin VB.Menu mnuplay 
         Caption         =   "Play"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnustop 
         Caption         =   "Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuaddfile2 
         Caption         =   "Add File"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAdddir2 
         Caption         =   "Add Dir"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuRemall 
         Caption         =   "Remove All"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuVolume 
         Caption         =   "Volume"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnumini 
         Caption         =   "Minimize"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuOnTop 
         Caption         =   "Keep on Top"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuLoadList 
         Caption         =   "Load List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSavelist 
         Caption         =   "Save List"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCurrentIndex As Integer
Dim tinseconden As Integer
Dim minuten As Integer
Dim seconden As Integer

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Check1_Click()
If Check1.Value = 0 Then
AlwaysOnTop Form1, True
Else
AlwaysOnTop Form1, False
End If
End Sub



Private Sub CmdLoadList_Click()
End Sub


Private Sub CmdOnTop_Click()
    If Check3.Value = 0 Then
    AlwaysOnTop Form1, True
    Check3.Value = 1
    Else
    If Check3.Value = 1 Then
    AlwaysOnTop Form1, False
    Check3.Value = 0
    End If
    End If
End Sub

Private Sub CmdSaveList_Click()

End Sub

Private Sub CmdSysinfo_Click()
Call StartSysInfo
End Sub

Private Sub CmdAbout_Click()
Form2.Show vbModal
End Sub

Private Sub CmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub

Private Sub CmdClear_Click()

End Sub

Private Sub CmdMute_Click()
If MediaPlayer1.Mute = False Then
MediaPlayer1.Mute = True
Else
MediaPlayer1.Mute = False
End If
End Sub

Private Sub CmdPause_Click()
If List1.ListCount = 0 Then Exit Sub
If Text1.Text = "" Then Exit Sub
If CmdPause.Caption = "Pause" Then
MediaPlayer1.Pause
CmdPause.Caption = "Resume"
Else
MediaPlayer1.Play
CmdPause.Caption = "Pause"
End If
End Sub

Private Sub CmdPlay_Click()
Text1 = List1.Text
On Error Resume Next
MediaPlayer1.FileName = Text1.Text
If Text1.Text <> "" Then
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration
CmdPause.Enabled = True
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub CmdRem_Click()

End Sub
Private Sub CmdStop_Click()
MediaPlayer1.Stop
Slider1.Value = 0
Text1.Text = ""
CmdPause.Enabled = False
End Sub

Private Sub Cmdvol_Click()
On Error GoTo errorhandler
       Dim lngresult As Long
       lngresult = Shell("c:\windows\Sndvol32.exe", vbNormalFocus)
       Exit Sub
errorhandler:
    lngresult = Shell("c:\winnt\system32\Sndvol32.exe", vbNormalFocus)
End Sub

Private Sub Command1_Click()
Form1.WindowState = 1
End Sub



Private Sub Command2_Click()
Form1.Height = 8865
WebBrowser1.Visible = True
startingaddress = "http://home.planetinternet.be/~jansse54"
WebBrowser1.Navigate startingaddress
Text2.Visible = True
Text2.Text = "connecting to startpage..."
Label13.Visible = True
Label13.Caption = "connecting to http://home.planetinternet.be/~jansse54"
Label14.Visible = True

End Sub

Private Sub Form_Load()
Form1.Height = 5640
CmdPause.Caption = "Pause"
Slider3.Value = MediaPlayer1.Volume
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()
  PopupMenu addmenu, , X + 4000, Y + 2000
End Sub



Private Sub Label14_Click()

Form1.Height = 5640
WebBrowser1.Visible = False


Text2.Visible = False

Label13.Visible = False
Label14.Visible = False
End Sub

Private Sub Label5_Click()
Dim file As String
CommonDialog2.DialogTitle = "Load your list."
   CommonDialog2.MaxFileSize = 16384
   CommonDialog2.FileName = ""
   CommonDialog2.Filter = "list Files|*.mlt"
   CommonDialog2.ShowOpen     ' = 1
If CommonDialog2.FileName = "" Then Exit Sub
file = CommonDialog2.FileName
Dim a As String
Dim X As String
On Error GoTo Error
Open file For Input As #1
Do Until EOF(1)
Input #1, a$
List1.AddItem a$
Loop
Close 1
Exit Sub
Error:
X = MsgBox("File Not Found", vbOKOnly, "Error")

End Sub

Private Sub Label6_Click()
If List1.ListIndex = -1 Then
MsgBox "No file selected", vbExclamation, "Error"
Else
List1.RemoveItem List1.ListIndex
Text1.Text = ""
End If

End Sub

Private Sub Label7_Click()
List1.Clear
Text1.Text = ""
End Sub

Private Sub Label8_Click()
Dim naampje As String
naampje = InputBox("Name of the list?", "ListName")
naampje = naampje & ".mlt"
Open (App.Path & "\" & naampje) For Output As #1
       Dim i%
       For i = 0 To List1.ListCount - 1
       Print #1, List1.List(i)
       Next
       Close #1
    'CommonDialog3.DialogTitle = "Save your list."
    'CommonDialog3.MaxFileSize = 16384
    'CommonDialog3.filename = ""
    'CommonDialog3.Filter = "list Files|*.mlt"
    'CommonDialog3.InitDir = App.Path
    'CommonDialog3.DefaultExt = ".mlt"
    'CommonDialog3.ShowSave

End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
End Sub

Private Sub List1_DblClick()
Text1 = List1.Text
On Error Resume Next
MediaPlayer1.FileName = Text1.Text
If Text1.Text <> "" Then
MediaPlayer1.Play
Slider1.Max = MediaPlayer1.Duration
CmdPause.Enabled = True
Else
MsgBox "No file to play", vbOKOnly, "Error"
End If
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)

If OptRND.Value = True Then
Randomize Timer
 MyValue = Int((List1.ListCount * Rnd))
    List1.ListIndex = MyValue
    MediaPlayer1.FileName = Text1.Text
    If Text1.Text <> "" Then
        MediaPlayer1.Play
        Slider1.Max = MediaPlayer1.Duration
        CmdPause.Enabled = True
        Exit Sub
    Else
        MsgBox "No file to play", vbOKOnly, "Error"
    End If
Else
    If OptRes.Value = True Then
        X = 0
        End If
        If List1.ListCount = 0 Then
Else
    If List1.ListIndex + 1 = List1.ListCount Then
        List1.ListIndex = 0
    Else
        List1.ListIndex = List1.ListIndex + 1
    End If
    If List1.Text <> "" Then
        MediaPlayer1.FileName = List1.Text
        MediaPlayer1.Play
        Slider1.Max = MediaPlayer1.Duration
        Text1.Text = LCase$(Left$(List1.Text, Len(List1.Text) - 4))
    End If
End If
If OptNON.Value = True Then
    MediaPlayer1.Stop
End If
End If
End Sub

Private Sub mnuAbout_Click()
CmdAbout_Click
End Sub

Private Sub mnuadddir_Click(Index As Integer)
Form3.Show vbModal
End Sub

Private Sub mnuAdddir2_Click()
Form3.Show vbModal
End Sub

Private Sub mnuaddfile_Click(Index As Integer)
CommonDialog1.DialogTitle = "Load your MP3."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "MP3 Files|*.MP3"
   CommonDialog1.ShowOpen     ' = 1
If CommonDialog1.FileTitle <> "" Then
List1.AddItem CommonDialog1.FileName
Text1.Text = CommonDialog1.FileName
Exit Sub
Else
Exit Sub
End If
End Sub


Private Sub mnuaddfile2_Click()
CommonDialog1.DialogTitle = "Load your MP3."
   CommonDialog1.MaxFileSize = 16384
   CommonDialog1.FileName = ""
   CommonDialog1.Filter = "MP3 Files|*.MP3"
   CommonDialog1.ShowOpen     ' = 1
If CommonDialog1.FileTitle <> "" Then
List1.AddItem CommonDialog1.FileName
Text1.Text = CommonDialog1.FileName
Exit Sub
Else
Exit Sub
End If
End Sub

Private Sub mnuExit_Click()
End
Unload Me
End Sub

Private Sub mnuLoadList_Click()
Label5_Click
End Sub

Private Sub mnumini_Click()
Command1_Click
End Sub

Private Sub mnuMute_Click()
CmdMute_Click
End Sub

Private Sub mnuOnTop_Click()
CmdOnTop_Click
End Sub

Private Sub mnuplay_Click()
CmdPlay_Click
End Sub

Private Sub mnuRemall_Click()
CmdClear_Click
End Sub

Private Sub mnuRemove_Click()
CmdRem_Click
End Sub

Private Sub mnuSavelist_Click()
Label8_Click
End Sub

Private Sub mnustop_Click()
CmdStop_Click
End Sub

Private Sub mnuVolume_Click()
Cmdvol_Click
End Sub

Private Sub Option1_Click()
Label1.ForeColor = &HFFFF&
Text1.ForeColor = &HFFFF&
Label.ForeColor = &HFFFF&
Label2.ForeColor = &HFFFF&
List1.ForeColor = &HFFFF&
Label4.ForeColor = &HFFFF&
Label5.ForeColor = &HFFFF&
Label6.ForeColor = &HFFFF&
Label7.ForeColor = &HFFFF&
Label8.ForeColor = &HFFFF&
Label10.ForeColor = &HFFFF&
Label12.ForeColor = &HFFFF&
Label14.ForeColor = &HFFFF&
Frame1.ForeColor = &HFFFF&
TxtTime.ForeColor = &HFFFF&
Check3.ForeColor = &HFFFF&
OptRND.ForeColor = &HFFFF&
OptRes.ForeColor = &HFFFF&
OptNON.ForeColor = &HFFFF&
End Sub

Private Sub Option2_Click()
Label1.ForeColor = &HF4F000
Text1.ForeColor = &HF4F000
Label.ForeColor = &HF4F000
Label2.ForeColor = &HF4F000
List1.ForeColor = &HF4F000
Label4.ForeColor = &HF4F000
Label5.ForeColor = &HF4F000
Label6.ForeColor = &HF4F000
Label7.ForeColor = &HF4F000
Label8.ForeColor = &HF4F000
Label10.ForeColor = &HF4F000
Label12.ForeColor = &HF4F000
Label14.ForeColor = &HF4F000
Frame1.ForeColor = &HF4F000

TxtTime.ForeColor = &HF4F000
Check3.ForeColor = &HF4F000
OptRND.ForeColor = &HF4F000
OptRes.ForeColor = &HF4F000
OptNON.ForeColor = &HF4F000
End Sub

Private Sub Option3_Click()
Label1.ForeColor = &HFF00&
Text1.ForeColor = &HFF00&
Label.ForeColor = &HFF00&
Label2.ForeColor = &HFF00&
List1.ForeColor = &HFF00&
Label4.ForeColor = &HFF00&
Label5.ForeColor = &HFF00&
Label6.ForeColor = &HFF00&
Label7.ForeColor = &HFF00&
Label8.ForeColor = &HFF00&
Label10.ForeColor = &HFF00&
Label12.ForeColor = &HFF00&
Label14.ForeColor = &HFF00&
Frame1.ForeColor = &HFF00&

TxtTime.ForeColor = &HFF00&
Check3.ForeColor = &HFF00&
OptRND.ForeColor = &HFF00&
OptRes.ForeColor = &HFF00&
OptNON.ForeColor = &HFF00&
End Sub

Private Sub Option4_Click()
Form1.BackColor = &HFF8040
List1.BackColor = &HFF8040
Text1.BackColor = &HFF8040
Frame1.BackColor = &HFF8040

End Sub

Private Sub Option5_Click()
Form1.BackColor = &H0&
List1.BackColor = &H0&

Text1.BackColor = &H0&
Frame1.BackColor = &H0&
End Sub

Private Sub Option6_Click()
Form1.BackColor = &H4F6531
List1.BackColor = &H4F6531

Text1.BackColor = &H4F6531
Frame1.BackColor = &H4F6531
End Sub





Private Sub Slider1_Scroll()
MediaPlayer1.CurrentPosition = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
On Error GoTo DamnYou
If Slider2.Value > -500 And Slider2.Value < 500 Then
Label4.Caption = "Center"
End If
If Slider2.Value < -500 Then
Label4.Caption = "Left"
End If
If Slider2.Value > 500 Then
Label4.Caption = "Right"
End If
MediaPlayer1.Balance = Slider2.Value
Exit Sub
DamnYou:
MsgBox "Err"
Exit Sub
End Sub


Private Sub Slider3_Scroll()
Dim pim, sha
Dim foo As Integer, poo As Integer
Label2.ForeColor = RGB(0 + Slider3.Value / 10, 0, 0)
sha = Slider3.Value - 2500
MediaPlayer1.Volume = sha
On Error GoTo hell
poo = Slider3.min
foo = Slider3.Value
Label2.Caption = foo \ 25 & " %"
hell:
Exit Sub
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next


    If KeyCode = 13 Then
        Url$ = Text2.Text
        WebBrowser1.Navigate Url$
    End If
If KeyCode = 13 Then
Label13.Caption = "connecting to..." + Text2.Text + "..." + "just a minute please"
End If
End Sub

Private Sub Timer1_Timer()
Slider1.Value = MediaPlayer1.CurrentPosition
tinseconden = MediaPlayer1.CurrentPosition
Dim min As Integer
Dim sec As Integer
min = tinseconden \ 60
sec = tinseconden - (min * 60)
If sec = "-1" Then sec = "0"
TxtTime.Text = min & ":" & sec
Slider1.Value = MediaPlayer1.CurrentPosition
twinsec = MediaPlayer1.Duration
Dim a As Integer
Dim b As Integer
a = twinsec \ 60
b = twinsec - (a * 60)
If b = "-1" Then sec = "0"
 Label9.Caption = a & ":" & b
 End Sub

Public Sub AlwaysOnTop(Form1 As Form, SetOnTop As Boolean)
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If

    SetWindowPos Form1.hwnd, lFlag, Form1.Left / Screen.TwipsPerPixelX, _
    Form1.Top / Screen.TwipsPerPixelY, Form1.Width / Screen.TwipsPerPixelX, _
    Form1.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub






Private Sub WebBrowser1_DownloadComplete()
Text2.Text = WebBrowser1.LocationURL
Label13.Caption = "Download complete -   " + WebBrowser1.LocationName
End Sub
