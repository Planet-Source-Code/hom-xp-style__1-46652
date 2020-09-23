VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "XP Style"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5895
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   210
      ScaleHeight     =   3255
      ScaleWidth      =   5415
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   5410
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   50
         Picture         =   "MainFrm.frx":1CFA
         ScaleHeight     =   1215
         ScaleWidth      =   5295
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   5295
         Begin VB.TextBox Txt1 
            Height          =   2.45745e5
            Left            =   2.45745e5
            MultiLine       =   -1  'True
            TabIndex        =   15
            Text            =   "MainFrm.frx":19234
            Top             =   2.45745e5
            Visible         =   0   'False
            Width           =   2.45745e5
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   1
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox tbxApp 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   1920
         Width           =   4095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Change the Settings"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Make XP Style"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Makes the selected Application look like Windows XP Applications"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "Exit XP Style"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "Application:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Restore"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   19
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton CreatManifest 
      Caption         =   "&Creat Manifest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2700
      Left            =   225
      TabIndex        =   17
      Top             =   555
      Visible         =   0   'False
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   4763
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Application's Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Application Path"
         Object.Width           =   5733
      EndProperty
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2.40125e5
      Top             =   1.80001e5
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2.45745e5
      Top             =   2.45745e5
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2.45745e5
      Top             =   2.45745e5
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   6120
      ScaleHeight     =   3255
      ScaleWidth      =   5415
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   30
         Picture         =   "MainFrm.frx":194AF
         ScaleHeight     =   2535
         ScaleWidth      =   5295
         TabIndex        =   20
         Top             =   120
         Width           =   5295
         Begin VB.CheckBox Check1 
            Caption         =   "Creat manifest file as a &Hidden file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox Check3 
            Caption         =   "&Run application after creating manifest file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3420
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Add file to R&estore List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox Check5 
            Caption         =   "E&xit XP Style after creating manifest file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1800
            Width           =   3135
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Creat manifest file as a &System file"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Value           =   1  'Checked
            Width           =   2895
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Apply Changes"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Default"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Restores the settings to default type"
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.TextBox tbxPath 
      Height          =   375
      Left            =   2.45745e5
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2.45745e5
      Width           =   4815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2.45745e5
      Top             =   2.45745e5
   End
   Begin VB.TextBox tbx2 
      Height          =   1335
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "MainFrm.frx":458FD
      Top             =   8400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tbx1 
      Height          =   1005
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "MainFrm.frx":45AA9
      Top             =   8400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3365
      Left            =   130
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   450
      Width           =   5590
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6588
      TabWidthStyle   =   2
      TabFixedWidth   =   2999
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Creat Manifet File"
            Key             =   "C"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Restore"
            Key             =   "R"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   "S"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Newer 
      Caption         =   "Newer"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu Refresh1 
         Caption         =   "Re&fresh"
      End
   End
   Begin VB.Menu popup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu Open1 
         Caption         =   "&Open"
      End
      Begin VB.Menu Folder 
         Caption         =   "Open &Containing Folder"
      End
      Begin VB.Menu Properties1 
         Caption         =   "&Properties"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu restore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu Delete1 
         Caption         =   "&Delete from the list"
      End
      Begin VB.Menu Rewrite1 
         Caption         =   "R&ewrite the manifest"
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Open2 
         Caption         =   "&Open...          "
         Shortcut        =   ^O
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About        "
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TargetName As String
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional
    lpClass As String 'Optional
    hkeyClass As Long 'Optional
    dwHotKey As Long 'Optional
    hIcon As Long 'Optional
    hProcess As Long 'Optional
    End Type
    Private Const SEE_MASK_INVOKEIDLIST = &HC
    Private Const SEE_MASK_NOCLOSEPROCESS = &H40
    Private Const SEE_MASK_FLAG_NO_UI = &H400


Private Declare Function ShellExecuteEx Lib "shell32" _
    Alias "ShellExecuteExA" _
    (SEI As SHELLEXECUTEINFO) As Long

Dim Tabs As Integer
Dim TabFocus As Boolean
Dim AddList As Boolean
Dim SystemFile As Boolean
Dim HiddenFile, XP As Boolean
Dim Nim As Integer
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private FSO As New FileSystemObject
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long

    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32s = 0


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    dwRevision As Long
    End Type
Private Sub ShowProperties(sFilename As String, hWndOwner As Long)
    
    'open a file properties property page fo
    '     r
    'specified file if return value
    Dim SEI As SHELLEXECUTEINFO
    
    'Fill in the SHELLEXECUTEINFO structure
    '


    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
        SEE_MASK_INVOKEIDLIST Or _
        SEE_MASK_FLAG_NO_UI
        .hwnd = hWndOwner
        .lpVerb = "properties"
        .lpFile = sFilename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    
    'call the API to display the property sh
    '     eet
    Call ShellExecuteEx(SEI)
    
End Sub
Public Function GetTarget(strPath As String) As String
    'Gets target path from a shortcut file
    On Error GoTo Error_Loading
    Dim wshShell As Object
    Dim wshLink As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set wshLink = wshShell.CreateShortcut(strPath)
    TargetName = wshLink.TargetPath
    Set wshLink = Nothing
    Set wshShell = Nothing
    Exit Function
Error_Loading:
    GetTarget = "Error occured."
End Function
Public Sub OS()
    Dim MsgEnd As String
    Dim junk
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    junk = GetVersionEx(osvi)
    If junk <> 0 And osvi.dwPlatformId = VER_PLATFORM_WIN32_NT And osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 1 Then XP = True
End Sub

Private Function TabS1()
'Clicking on tabstrip 1
        Tabs = 1
        Timer1.Enabled = True
        CreatManifest.Visible = False
        Picture1.Visible = False
        Picture2.Visible = True
        ListView1.Visible = False
        Command4.Visible = False
        ListView1.Enabled = False
        Command1.TabStop = True
        Command3.TabStop = True
        Command2.TabStop = True
        Command7.TabStop = True
        tbxApp.TabStop = True
        ListView1.TabStop = False
        Command4.TabStop = False
        Command5.TabStop = False
        Command6.TabStop = False
        Check1.TabStop = False
        Check2.TabStop = False
        Check3.TabStop = False
        Check4.TabStop = False
        Check5.TabStop = False
End Function
Private Function TabS2()
'Clicking on tabstrip 2
        Tabs = 2
        Picture1.Visible = False
        Picture2.Visible = False
        ListView1.Visible = True
        CreatManifest.Visible = True
        Command4.Visible = True
        ListView1.Enabled = True
        Check1.TabStop = False
        Check2.TabStop = False
        Check3.TabStop = False
        Check4.TabStop = False
        Check5.TabStop = False
        Command1.TabStop = False
        Command2.TabStop = False
        Command7.TabStop = False
        tbxApp.TabStop = False
        Command5.TabStop = False
        Command6.TabStop = False
        Command4.TabStop = True
        ListView1.TabStop = True
End Function
Private Function TabS3()
'Clicking on tabstrip3
        Tabs = 3
        TabStrip1.SelectedItem = TabStrip1.Tabs.Item(3)
        Picture1.Visible = True
        Picture2.Visible = False
        Command6.TabStop = True
        Command5.TabStop = True
        Check1.TabStop = True
        Check2.TabStop = True
        Check3.TabStop = True
        Check4.TabStop = True
        Check5.TabStop = True
        Command1.TabStop = False
        Command3.TabStop = False
        Command2.TabStop = False
        Command7.TabStop = False
        Command4.TabStop = False
        tbxApp.TabStop = False
        ListView1.TabStop = False
        ListView1.Visible = False
        Command4.Visible = False
        CreatManifest.Visible = False
        ListView1.Enabled = False
End Function
Private Sub about_Click()
'Shows About Form
    frmAbout.Show
End Sub

Private Sub Check1_Click()
'Enables "Apply" and "Default" Buttons
    Command6.Enabled = True
    Command5.Enabled = True
    If Check1.Value = 0 Then
        Check2.Value = 0
        Check2.Enabled = False
    Else
        Check2.Enabled = True
    End If
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Check2_Click()
'Enables "Apply" and "Default" Buttons
    Command6.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub
Private Sub Check3_Click()
'Enables "Apply" and "Default" Buttons and enables Systemfile CheckBox
    Command6.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub Check3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Check4_Click()
'Enables "Apply" and "Default" Buttons
    Command6.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub Check4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub


Private Sub Check5_Click()
'Enables "Apply" and "Default" Buttons
    Command6.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub Check5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub


Private Sub Command1_Click()
'Opens file select
On Error Resume Next
Desktop1 = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
    With CommonDialog1
        .DialogTitle = "Application:"
        .CancelError = False
        .FileName = ""
        .InitDir = Desktop1
        .Filter = "Application|*.exe"
        .MaxFileSize = 32000
        .ShowOpen
    End With
    If CommonDialog1.FileTitle <> "" And CommonDialog1.FileName <> "" Then
        tbxApp.Text = CommonDialog1.FileTitle
        tbxPath.Text = CommonDialog1.FileName
    End If
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Then
        Command1.Value = True
    ElseIf KeyCode = 88 Then
        Command3.Value = True
    ElseIf KeyCode = 83 Then
        Command7.Value = True
    ElseIf KeyCode = 77 And Command2.Enabled = True Then
        Command2.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Command2_Click()
    'Finds file name

    For a = Len(tbxApp.Text) To 1 Step -1
        If Mid(tbxApp.Text, a, 1) = "\" Then
            tbxPath.Text = tbxApp.Text
            tbxApp.Text = Mid(tbxApp.Text, a + 1, Len(tbxApp.Text))
        End If
    Next
    'Cause to accepts applications only
    If Right(tbxApp.Text, 4) <> ".exe" And Right(tbxApp.Text, 4) <> ".EXE" And Right(tbxApp.Text, 4) <> ".EXe" And Right(tbxApp.Text, 4) <> ".ExE" And Right(tbxApp.Text, 4) <> ".Exe" And Right(tbxApp.Text, 4) <> ".eXE" And Right(tbxApp.Text, 4) <> ".exE" And Right(tbxApp.Text, 4) <> ".eXe" And Right(tbxApp.Text, 4) <> ".lnk" And Right(tbxApp.Text, 4) <> ".LNK" And Right(tbxApp.Text, 4) <> ".LNk" And Right(tbxApp.Text, 4) <> ".LnK" And Right(tbxApp.Text, 4) <> ".Lnk" And Right(tbxApp.Text, 4) <> ".lNK" And Right(tbxApp.Text, 4) <> ".lnK" And Right(tbxApp.Text, 4) <> ".lNk" Then
        tbxApp.Text = ""
        Call MsgBox("You must enter an application name with EXE or LNK extension", vbCritical, "Application")
        Timer1.Enabled = True
        Exit Sub
    End If

    'Search to find the file
    If FSO.FileExists(tbxPath.Text) = False Then
        Call MsgBox("There is no file with such name!", vbCritical, "File missing...")
        Exit Sub
    End If

    If Right(tbxApp.Text, 4) = ".lnk" Or Right(tbxApp.Text, 4) = ".LNK" Or Right(tbxApp.Text, 4) = ".LNk" Or Right(tbxApp.Text, 4) = ".LnK" Or Right(tbxApp.Text, 4) = ".Lnk" Or Right(tbxApp.Text, 4) = ".lNK" Or Right(tbxApp.Text, 4) = ".lnK" Or Right(tbxApp.Text, 4) = ".lNk" Then
        Call GetTarget(tbxPath.Text)
        tbxPath.Text = TargetName
        For a = Len(tbxPath.Text) To 1 Step -1
            If Mid(tbxPath.Text, a, 1) = "\" Then
                tbxApp.Text = Mid(tbxPath.Text, a + 1, Len(tbxPath.Text))
            End If
        Next
         
    End If

    'Confirm Msgbox
    Nim = MsgBox("Are you sure you want to change " & """" & Mid(tbxApp.Text, 1, Len(tbxApp.Text) - 4) & """" & " to Windows XP Style?", vbYesNo, "Confirm")
    If Nim = 6 Then

        'Search to find the manifest file
        If FSO.FileExists(tbxPath.Text & ".manifest") = True Then
             On Error GoTo Error2
             FSO.DeleteFile (tbxPath.Text & ".manifest")
        End If
        
        'Creats manifest file
        Open tbxPath.Text & ".manifest" For Append As 1
        Print #1, Txt1.Text
        Close 1

        
        If Check3.Value = 1 Then Call Shell(tbxPath.Text, vbNormalFocus)

        
        SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", tbxPath.Text, "WIN2000"
        'Adds to list
        If AddList = True Then
            On Error Resume Next
            Dim cmd, cmFileName As String, cmCommand
            'Save to manifest settings
            SaveRegString HKEY_LOCAL_MACHINE, "Software\ManifestMaker", tbxPath.Text, ""
            'save to win reg
            cmd = Trim$(Command$)
            cmd = Right(cmd, Len(cmd) - 1)
            cmd = Left(cmd, Len(cmd) - 1)
            cmCommand = Left(cmd, 1)
            cmFileName = Right(cmd, Len(cmd) - 1)
            cmFileName = Left(cmFileName, Len(cmFileName) - 1)
        
            If cmd > "" Then
                Select Case cmCommand
                Case "E"

                    If FSO.FileExists(cmFileName & ".manifest") = True Then
                        FSO.DeleteFile (cmFileName & ".manifest")
                    End If

                    Open cmFileName & ".manifest" For Append As 1
                    'fix this
                    Print #1, tbx1.Text & ShortFileName(cmFileName) & """"
                    Print #1, tbx2.Text
                    Close 1
                    'save to manifest settings
                    SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FileName", ShortFileName(cmFileName)
                    SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FilePath", cmFileName
                    'save to win reg
                    SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName, "WIN2000"
    
                Case "D"

                    'Delete  from registry
                    DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName)
                    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName
                    'Deletes manifest
                    FSO.DeleteFile (cmFileName & ".manifest")
                End Select
            End If

            Dim fApp
            Dim rApp, fName, fPath
            For Random = 1 To ListView1.ListItems.Count
                Call ListView1.ListItems.Remove(1)
            Next
    
            'load saved alerts to listview
            Call GetValues(HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1)
        End If
    Else
        Exit Sub
    End If
    'Set file as Hidden or Hidden and System
    If SystemFile = True And HiddenFile = True Then
         SetAttr tbxPath.Text & ".manifest", vbSystem Or vbHidden
    ElseIf HiddenFile = True Then
         SetAttr tbxPath.Text & ".manifest", vbHidden
    End If
    If Check5.Value = 1 Then End
    If ListView1.ListItems.Count = 0 Then Command3.Enabled = False
    'Runs application
    'Clears the form
    tbxApp.Text = ""
    tbxPath.Text = ""
    Command2.Enabled = False
    If Check2.Value = 5 Then
        End
    Else
        Command1.SetFocus
        Exit Sub
    End If
    Exit Sub

1:
    'Error Found:
    Call MsgBox("File not found. Maybe the file name is not one part.", vbOKOnly, "No File")
    tbxApp.Text = ""
    tbxPath.Text = ""
    Command2.Enabled = False
    Exit Sub
Error2:
    Call MsgBox("The manifest already exists and it is Read-only. Cannot change it.", vbCritical, "Error in writing manifest")
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Then
        Command1.Value = True
    ElseIf KeyCode = 88 Then
        Command3.Value = True
    ElseIf KeyCode = 83 Then
        Command7.Value = True
    ElseIf KeyCode = 77 And Command2.Enabled = True Then
        Command2.Value = True
    Else
        Beep
    End If
End Sub
Private Sub Command3_Click()
    'Exits
    End
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Then
        Command1.Value = True
    ElseIf KeyCode = 88 Then
        Command3.Value = True
    ElseIf KeyCode = 83 Then
        Command7.Value = True
    ElseIf KeyCode = 77 And Command2.Enabled = True Then
        Command2.Value = True
    Else
        Beep
    End If
End Sub


Private Sub Command4_Click()
    On Error Resume Next
    If ListView1.SelectedItem.Selected = False Then
        MsgBox "Please Select Application To Restore"
        Exit Sub
    End If
    'Confirm Msgbox
If FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = True Then
    h = MsgBox("Are you sure you want to restore " & """" & Mid(ListView1.SelectedItem.Text, 1, Len(ListView1.SelectedItem.Text)) & """" & "?", vbYesNo, "Confirm to restore")
    If h = 6 Then
        Command4.Enabled = False
        'Delete  from registry
        DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
        DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
        'delete manifest
        FSO.DeleteFile (ListView1.SelectedItem.SubItems(1) & ".manifest")
        'Delete from listview
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
        Exit Sub
    Else
        Exit Sub
    End If
ElseIf FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = False And FSO.FileExists(ListView1.SelectedItem.SubItems(1) & ".manifest") = False Then
    Message = MsgBox("The file and its manifest file have been deleted or moved. Do you want to delete it from the list?", vbYesNo, "File Missing...")
    If Message = 6 Then
        Command4.Enabled = False
        DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
        DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
        Exit Sub
    End If
ElseIf FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = False And FSO.FileExists(ListView1.SelectedItem.SubItems(1) & ".manifest") = True Then
    Message = MsgBox("The file has been deleted or moved. Do you want to delete the manifest file?", vbYesNo, "File Missing...")
    If Message = 6 Then
        Command4.Enabled = False
        FSO.DeleteFile (ListView1.SelectedItem.SubItems(1) & ".manifest")
        DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
        DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
        Exit Sub
    End If
ElseIf FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = True And FSO.FileExists(ListView1.SelectedItem.SubItems(1) & ".manifest") = False Then
    Message = MsgBox("The manifest file has been deleted or moved. Do you want to delete it from the list?", vbYesNo, "File Missing...")
    If Message = 6 Then
        Command4.Enabled = False
        DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
        DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
        Exit Sub
    End If
End If
ListView1.SetFocus
End Sub
Private Sub Command4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 114 Then
        Command4.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Command5_Click()
    'Sets to default
    Check1.Value = 1
    Check2.Value = 1
    Check3.Value = 1
    Check4.Value = 1
    Check5.Value = 0
    Command6.Enabled = True
    Command5.Enabled = False
End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Command6_Click()
    'Apply changes
    Command6.Enabled = False
    If Check1.Value = 1 Then
        HiddenFile = True
    Else
        HiddenFile = False
    End If
    If Check2.Value = 1 Then
        SystemFile = True
    Else
        SystemFile = False
    End If
    If Check4.Value = 1 Then
        AddList = True
    Else
        AddList = False
    End If
End Sub

Private Sub Command6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 72 And Check1.Value = 0 Then
        Check1.Value = 1
        Check1.SetFocus
    ElseIf KeyCode = 72 And Check1.Value = 1 Then
        Check1.Value = 0
        Check1.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 0 And Check2.Enabled = True Then
        Check2.Value = 1
        Check2.SetFocus
    ElseIf KeyCode = 83 And Check2.Value = 1 And Check2.Enabled = True Then
        Check2.Value = 0
        Check2.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 0 Then
        Check3.Value = 1
        Check3.SetFocus
    ElseIf KeyCode = 82 And Check3.Value = 1 Then
        Check3.Value = 0
        Check3.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 0 Then
        Check4.Value = 1
        Check4.SetFocus
    ElseIf KeyCode = 69 And Check4.Value = 1 Then
        Check4.Value = 0
        Check4.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 0 Then
        Check5.Value = 1
        Check5.SetFocus
    ElseIf KeyCode = 88 And Check5.Value = 1 Then
        Check5.Value = 0
        Check5.SetFocus
    ElseIf KeyCode = 68 And Command5.Enabled = True Then
        Command5.Value = True
    ElseIf KeyCode = 65 And Command6.Enabled = True Then
        Command6.Value = True
    Else
        Beep
    End If
End Sub

Private Sub Command7_Click()
    'Shows settings
    TabStrip1.SelectedItem = TabStrip1.Tabs.Item(3)
    TabS3
    Tabs = 3
    Check1.SetFocus
End Sub

Private Sub Command7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 66 Then
        Command1.Value = True
    ElseIf KeyCode = 88 Then
        Command3.Value = True
    ElseIf KeyCode = 83 Then
        Command7.Value = True
    ElseIf KeyCode = 77 And Command2.Enabled = True Then
        Command2.Value = True
    Else
        Beep
    End If
End Sub

Private Sub CreatManifest_Click()
    TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1)
    TabS1
    Tabs = 1
    tbxApp.SetFocus
End Sub

Private Sub Delete1_Click()
    Message = MsgBox("Are you sure you want to delete " & """" & ListView1.SelectedItem.Text & """" & " from the list?", vbYesNo, "Delete Confirm")
    If Message = 6 Then
        DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        Command4.Enabled = False
    End If
End Sub

Private Sub exit_Click()
    'Exits
    End
End Sub

Private Sub Folder_Click()
Dim a As Integer
Dim OpenFolder As String
   For a = Len(ListView1.SelectedItem.SubItems(1)) To 1 Step -1
        If Mid(ListView1.SelectedItem.SubItems(1), a, 1) = "\" Then
            OpenFolder = Mid(ListView1.SelectedItem.SubItems(1), 1, a)
            Shell "Explorer " + OpenFolder, vbNormalFocus
            Exit Sub
        End If
    Next
End Sub

Private Sub Form_Initialize()
    'Uses comctl32.dll
    Call InitCommonControls
End Sub
Private Sub Form_Load()
    Call OS
    If XP = False Then
        OSMessage = MsgBox("You must use this application in Windows XP. Are you sure you want to continue?", vbYesNo, "Operating System error")
        If OSMessage = 7 Then End
    End If

    Dim cmd, cmFileName As String, cmCommand
    If FSO.FileExists(App.Path & "\" & App.EXEName & ".exe" & ".manifest") = False Then
        On Error Resume Next
        MainFrm.Hide
        Open cmFileName & App.Path & "\" & App.EXEName & ".exe" & ".manifest" For Append As 1
        'fix this
        Print #1, Txt1.Text
        Close 1
        SetAttr App.Path & "\" & App.EXEName & ".exe" & ".manifest", vbSystem Or vbHidden
        Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
        FSO.DeleteFile (App.Path & "\" & App.EXEName & ".exe" & ".manifest")
        End
    End If


    'sets Variables and Properties
    Command4.Enabled = False
    TabFocus = False
    Tabs = 1
    AddList = True
    HiddenFile = True
    SystemFile = True

    Call GetValues(HKEY_LOCAL_MACHINE, "Software\ManifestMaker", Me.ListView1)

    '''this code is here for the context menus
    On Error Resume Next
    cmd = Trim$(Command$)
    cmd = Right(cmd, Len(cmd) - 1)
    cmd = Left(cmd, Len(cmd) - 1)
    cmCommand = Left(cmd, 1)
    cmFileName = Right(cmd, Len(cmd) - 1)
    cmFileName = Left(cmFileName, Len(cmFileName) - 1)
    If cmd > "" Then
        Select Case cmCommand
        Case "E"
            If FSO.FileExists(cmFileName & ".manifest") = True Then
                FSO.DeleteFile (cmFileName & ".manifest")
            End If
            Open cmFileName & ".manifest" For Append As 1
            'fix this
            Print #1, tbx1.Text & ShortFileName(cmFileName) & """"
            Print #1, tbx2.Text
            Close 1
            'save to manifest settings
            SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FileName", ShortFileName(cmFileName)
            SaveRegString HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName), "FilePath", cmFileName
            'save to win reg
            SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName, "WIN2000"
        Case "D"
            'Delete  from registry
            DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ShortFileName(cmFileName)
            DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", cmFileName
            'delete manifest
            FSO.DeleteFile (cmFileName & ".manifest")
        End Select
        End
    End If
    Dim fApp
    Dim rApp, fName, fPath

    'load saved alerts to listview

    If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
    Picture1.Left = 240
    Picture2.Left = 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAbout
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    'Disables changing
    Cancel = 1
End Sub

Private Sub ListView1_Click()
On Error Resume Next
    If ListView1.SelectedItem.Selected = False Then
        Command4.Enabled = False
    Else
        Command4.Enabled = True
    End If
End Sub
Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count <> 0 Then
        If ListView1.SelectedItem.Selected = True Then
            If FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = True Then
                Call Shell(ListView1.SelectedItem.SubItems(1), vbNormalFocus)
            Else
                Message = MsgBox("The file might have been deleted or moved. Do you want to delete it from the list?", vbYesNo, "File Missing...")
                If Message = 6 Then
                    DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
                    DeleteValue HnKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
                    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                    If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
                End If
        End If
    End If
End If
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Message = MsgBox("Are you sure you want to delete " & """" & ListView1.SelectedItem.Text & """" & " from the list?", vbYesNo, "Delete Confirm")
        If Message = 6 Then
            DeleteValue HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1)
            ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
            Command4.Enabled = False
         End If
    End If
    Timer3.Enabled = True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 114 And Command4.Enabled = True Then
        Command4.Value = True
        KeyAscii = 0
    ElseIf KeyAscii = 99 Then
        CreatManifest.Value = True
    Else
        Beep
    End If
    KeyAscii = 0
End Sub
Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If ListView1.SelectedItem.Selected = False Then
        Command4.Enabled = False
    Else
        Command4.Enabled = True
    End If
 
    If ListView1.SelectedItem.Selected = True And KeyCode = 93 Then
        PopupMenu popup, , 330, (ListView1.SelectedItem.Index - 1) * 225 + 1120
    ElseIf ListView1.SelectedItem.Selected = False And KeyCode = 93 Then
        PopupMenu Newer, , 270, 900
    End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer3.Enabled = True
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If ListView1.ListItems.Count <> 0 Then
        If Button = 2 And ListView1.SelectedItem.Selected = True Then
            PopupMenu popup
        ElseIf Button = 2 And ListView1.SelectedItem.Selected = False Then
            PopupMenu Newer
        End If
    ElseIf ListView1.ListItems.Count = 0 And Button = 2 Then
        PopupMenu Newer
    End If
End Sub
Private Sub Open1_Click()
    If ListView1.SelectedItem.Selected = True Then
        If FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = True Then
            Call Shell(ListView1.SelectedItem.SubItems(1), vbNormalFocus)
        Else
            Message = MsgBox("The file might have been deleted or moved. Do you want to delete it from the list?", vbYesNo, "File Missing...")
                If Message = 6 Then
                    DeleteKey HKEY_LOCAL_MACHINE, "Software\OutersoftInc\ManifestMaker\Manifests\" & ListView1.SelectedItem.Text & ".exe"
                    DeleteValue HnKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1)
                    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
                    If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
                End If
        End If
    End If
End Sub

Private Sub Open2_Click()
    If TabStrip1.SelectedItem <> TabStrip1.Tabs.Item(1) Then
        TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1)
        TabS1
        tbxApp.SetFocus
    End If
    Timer2.Enabled = True
End Sub

Private Sub Properties1_Click()
    If FSO.FileExists(ListView1.SelectedItem.SubItems(1)) = False Then
        Call MsgBox("The Application has been deleted or moved. Cannot open the properties page.", vbCritical, "Missing File...")
    Else
        Call ShowProperties(ListView1.SelectedItem.SubItems(1), Me.hwnd)
    End If
End Sub

Private Sub Refresh1_Click()
    ListView1.ListItems.Clear
    Call GetValues(HKEY_LOCAL_MACHINE, "Software\ManifestMaker", Me.ListView1)
    If ListView1.ListItems.Count = 0 Then Command4.Enabled = False
End Sub
Private Sub restore_Click()
    Command4.Value = True
End Sub

Private Sub Rewrite1_Click()
    On Error GoTo Error1
    Dim FileAttr As Integer
    FileAttr = 0
    If FSO.FileExists(ListView1.SelectedItem.SubItems(1) & ".manifest") = True Then
        Dim SSSS As file
        Set SSSS = FSO.GetFile(ListView1.SelectedItem.SubItems(1) & ".manifest")
        FileAttr = SSSS.Attributes
        FSO.DeleteFile (ListView1.SelectedItem.SubItems(1) & ".manifest")
    End If
    SaveRegString HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", ListView1.SelectedItem.SubItems(1), "WIN2000"
    SaveRegString HKEY_LOCAL_MACHINE, "Software\ManifestMaker", ListView1.SelectedItem.SubItems(1), ""
    Open ListView1.SelectedItem.SubItems(1) & ".manifest" For Append As 1
    Print #1, Txt1.Text
    Close 1
    If FileAttr = 0 Or FileAttr = 6 Then
        SetAttr (ListView1.SelectedItem.SubItems(1) & ".manifest"), vbHidden Or vbSystem
    ElseIf FileAttr = 2 Then
        SetAttr (ListView1.SelectedItem.SubItems(1) & ".manifest"), vbHidden
    ElseIf FileAttr = 34 Then
        SetAttr (ListView1.SelectedItem.SubItems(1) & ".manifest"), vbHidden Or vbArchive
    ElseIf FileAttr = 38 Then
        SetAttr (ListView1.SelectedItem.SubItems(1) & ".manifest"), vbHidden Or vbArchive Or vbSystem
    End If
    Exit Sub
Error1:
    Call MsgBox("Manifest file is Read-only. Cannot change it.", vbCritical, "Read-only manifest file")
End Sub

Private Sub TabStrip1_GotFocus()
    TabFocus = True
End Sub
Private Sub TabStrip1_KeyDown(KeyCode As Integer, Shift As Integer)
    'Shows the Tabs
    If KeyCode = 37 And TabStrip1.SelectedItem.Index = 2 Then
        TabS1
    ElseIf KeyCode = 37 And TabStrip1.SelectedItem.Index = 3 Then
        TabS2
    ElseIf KeyCode = 39 And TabStrip1.SelectedItem.Index = 1 Then
        TabS2
    ElseIf KeyCode = 39 And TabStrip1.SelectedItem.Index = 2 Then
        TabS3
    End If
End Sub

Private Sub TabStrip1_LostFocus()
    TabFocus = False
End Sub

Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1) Then Tabs = 1
If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(2) Then Tabs = 2
If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(3) Then Tabs = 3
Timer4.Enabled = True
End Sub

Private Sub TabStrip1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Shows the Tabs
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        TabS1
        If Tabs <> 1 And TabFocus = False Then tbxApp.SetFocus
        Tabs = 1
    Case 2
        TabS2
        If Tabs <> 2 And TabFocus = False Then ListView1.SetFocus
        Tabs = 2
    Case 3
        TabS3
        If Tabs <> 3 And TabFocus = False Then Check1.SetFocus
        Tabs = 3
    End Select
End Sub
Private Sub tbxApp_KeyPress(KeyAscii As Integer)
    'Does not accpet " in the TextBox
    If KeyAscii = 34 Then KeyAscii = 0
End Sub
Private Sub tbxApp_Change()
    'Does not accpet " in the TextBox
    tbxApp.Text = Replace(tbxApp.Text, """", "")
    If tbxApp.Text = "" Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
End Sub

Private Sub tbxApp_KeyUp(KeyCode As Integer, Shift As Integer)
Timer1.Enabled = False
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
    'Shows BalloonTooltip
If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1) Then
    On Error Resume Next
    If Right(tbxApp.Text, 4) <> ".exe" And Right(tbxApp.Text, 4) <> ".EXE" And Right(tbxApp.Text, 4) <> ".EXe" And Right(tbxApp.Text, 4) <> ".ExE" And Right(tbxApp.Text, 4) <> ".Exe" And Right(tbxApp.Text, 4) <> ".eXE" And Right(tbxApp.Text, 4) <> ".exE" And Right(tbxApp.Text, 4) <> ".eXe" And Right(tbxApp.Text, 4) <> ".lnk" And Right(tbxApp.Text, 4) <> ".LNK" And Right(tbxApp.Text, 4) <> ".LNk" And Right(tbxApp.Text, 4) <> ".LnK" And Right(tbxApp.Text, 4) <> ".Lnk" And Right(tbxApp.Text, 4) <> ".lNK" And Right(tbxApp.Text, 4) <> ".lnK" And Right(tbxApp.Text, 4) <> ".lNk" And tbxApp.Text <> "" Then
        ShowBalloonTip tbxApp.hwnd, "Application", "You must enter an application name here", etiInfo
        Timer1.Enabled = False
    End If
End If
End Sub

Private Sub Timer2_Timer()
    Command1.Value = True
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
    If ListView1.SelectedItem.Selected = False Then
        Command4.Enabled = False
    Else
        Command4.Enabled = True
    End If
    Timer3.Enabled = False
End Sub
Private Sub Timer4_Timer()
    If TabStrip1.SelectedItem = TabStrip1.Tabs.Item(1) And Tabs <> 1 Then
        Tabs = 1
        Call TabS1
        If TabFocus = False Then tbxApp.SetFocus
    ElseIf TabStrip1.SelectedItem = TabStrip1.Tabs.Item(2) And Tabs <> 2 Then
        Tabs = 2
        Call TabS2
        If TabFocus = False Then ListView1.SetFocus
    ElseIf TabStrip1.SelectedItem = TabStrip1.Tabs.Item(3) And Tabs <> 3 Then
        Tabs = 3
        Call TabS3
        If TabFocus = False Then Check1.SetFocus
    End If
    Timer4.Enabled = False
End Sub
