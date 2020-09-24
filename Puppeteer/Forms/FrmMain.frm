VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " Funny Jo3  /  By Soldier007"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   6840
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp 
      Height          =   735
      Index           =   1
      Left            =   5640
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picTemp 
      Height          =   735
      Index           =   0
      Left            =   4680
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox QuestionAskBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      MouseIcon       =   "FrmMain.frx":271ED
      MousePointer    =   99  'Custom
      ScaleHeight     =   855
      ScaleWidth      =   15375
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   10680
      Visible         =   0   'False
      Width           =   15375
      Begin VB.TextBox QuestionAskText 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4200
         TabIndex        =   36
         ToolTipText     =   "Please press <Enter> after typing a question.."
         Top             =   270
         Width           =   10815
      End
      Begin VB.Image QuestionAskBarExit 
         Height          =   615
         Left            =   120
         MouseIcon       =   "FrmMain.frx":274F7
         MousePointer    =   99  'Custom
         ToolTipText     =   "Click to closed Question Ask Bar.."
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.PictureBox Source 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00CCCCCC&
      BorderStyle     =   0  'None
      Height          =   13080
      Left            =   0
      Picture         =   "FrmMain.frx":27801
      ScaleHeight     =   12738.77
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   9
      Top             =   -30000
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox DisplayMenu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Left            =   0
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.TextBox NewBackground 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Top             =   2520
         Width           =   4695
      End
      Begin VB.PictureBox AboutButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5520
         MouseIcon       =   "FrmMain.frx":30444
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
      End
      Begin VB.PictureBox TellStoryButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   10920
         MouseIcon       =   "FrmMain.frx":3074E
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1155
      End
      Begin VB.PictureBox HelpButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3840
         MouseIcon       =   "FrmMain.frx":30A58
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
      End
      Begin VB.PictureBox SaveButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9600
         MouseIcon       =   "FrmMain.frx":30D62
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1155
      End
      Begin VB.PictureBox EditButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         MouseIcon       =   "FrmMain.frx":3106C
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1155
      End
      Begin VB.TextBox DisplayLeftPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox DisplayTopPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   3
         Left            =   480
         MouseIcon       =   "FrmMain.frx":31376
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   2160
         Width           =   135
      End
      Begin VB.PictureBox BackgroundButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   720
         MouseIcon       =   "FrmMain.frx":31680
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1155
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   480
         MouseIcon       =   "FrmMain.frx":3198A
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   1560
         Width           =   135
      End
      Begin VB.PictureBox DisplayAllButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5610
         MouseIcon       =   "FrmMain.frx":31C94
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1155
      End
      Begin VB.PictureBox PuppetPeedyButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   10920
         MouseIcon       =   "FrmMain.frx":31F9E
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1155
      End
      Begin VB.PictureBox PuppetMerlinButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8280
         MouseIcon       =   "FrmMain.frx":322A8
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1155
      End
      Begin VB.PictureBox PuppetGenieButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9600
         MouseIcon       =   "FrmMain.frx":325B2
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1155
      End
      Begin VB.PictureBox PuppetRobbyButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   12240
         MouseIcon       =   "FrmMain.frx":328BC
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1155
      End
      Begin VB.PictureBox MinimizeButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         MouseIcon       =   "FrmMain.frx":32BC6
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
      End
      Begin VB.PictureBox ShutdownButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         MouseIcon       =   "FrmMain.frx":32ED0
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1575
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
      End
      Begin VB.PictureBox DisplayRobbyButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4320
         MouseIcon       =   "FrmMain.frx":331DA
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1155
      End
      Begin VB.PictureBox DisplayGenieButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1755
         MouseIcon       =   "FrmMain.frx":334E4
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1155
      End
      Begin VB.PictureBox DisplayMerlinButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         MouseIcon       =   "FrmMain.frx":337EE
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1155
      End
      Begin VB.PictureBox DisplayPeedyButton 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3030
         MouseIcon       =   "FrmMain.frx":33AF8
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   1155
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1155
      End
      Begin VB.TextBox StoryBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   1935
         Left            =   8280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1080
         Width           =   6615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   480
         MouseIcon       =   "FrmMain.frx":33E02
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1320
         Width           =   135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   0
         Left            =   480
         MouseIcon       =   "FrmMain.frx":3410C
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1095
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Background"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   2040
         TabIndex        =   45
         Top             =   2535
         Width           =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   32
         X2              =   536
         Y1              =   257
         Y2              =   257
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Control Settings:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   480
         TabIndex        =   32
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   2400
         TabIndex        =   31
         Top             =   1830
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   720
         TabIndex        =   30
         Top             =   1830
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display Animations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   720
         TabIndex        =   26
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Current Background"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   720
         TabIndex        =   24
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign a  Puppet for story telling:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   8280
         TabIndex        =   21
         Top             =   3720
         Width           =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   32
         X2              =   536
         Y1              =   256
         Y2              =   256
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Puppet to Activate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Top             =   3000
         Width           =   2235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Story Teller Text: (Please enter the stories to be tell by Puppeteer)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   8280
         TabIndex        =   8
         Top             =   840
         Width           =   5865
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display Puppet Command Moves List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   1320
         Width           =   2640
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display Question Ask Bar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   1845
      End
   End
   Begin VB.ListBox Used 
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox PuppetCommandsGesturesBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5970
      Left            =   0
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   228
      TabIndex        =   37
      Top             =   4680
      Visible         =   0   'False
      Width           =   3420
      Begin VB.ListBox PuppetCommandMovesLists 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   4905
         Left            =   360
         MouseIcon       =   "FrmMain.frx":34416
         MousePointer    =   99  'Custom
         TabIndex        =   38
         ToolTipText     =   "Please select command moves by <Clicking> from this list.."
         Top             =   825
         Width           =   2655
      End
      Begin VB.Image PuppetCommandsGesturesBarExit 
         Height          =   615
         Left            =   120
         MouseIcon       =   "FrmMain.frx":34720
         MousePointer    =   99  'Custom
         ToolTipText     =   "Click to closed Puppet Command Move List.."
         Top             =   120
         Width           =   3135
      End
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   4680
      Top             =   6000
      _cx             =   847
      _cy             =   847
   End
   Begin AgentObjectsCtl.Agent Agent 
      Index           =   3
      Left            =   7080
      Top             =   6000
      _cx             =   847
      _cy             =   847
   End
   Begin AgentObjectsCtl.Agent Agent 
      Index           =   2
      Left            =   6480
      Top             =   6000
      _cx             =   847
      _cy             =   847
   End
   Begin AgentObjectsCtl.Agent Agent 
      Index           =   1
      Left            =   5880
      Top             =   6000
      _cx             =   847
      _cy             =   847
   End
   Begin AgentObjectsCtl.Agent Agent 
      Index           =   0
      Left            =   5280
      Top             =   6000
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Puppeteer Version 1.0
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Puppeteer Version 1.0. Contact me for
' additional help/suggestions
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

Option Explicit

Dim MerlinActivate As Boolean
Dim GenieActivate As Boolean
Dim PeedyActivate As Boolean
Dim RobbyActivate As Boolean
Dim Character As IAgentCtlCharacterEx
Dim Merlin As IAgentCtlCharacterEx
Dim Genie As IAgentCtlCharacterEx
Dim Peedy As IAgentCtlCharacterEx
Dim Robby As IAgentCtlCharacterEx
Dim Request As IAgentCtlRequest
Dim Shown As Boolean
Dim ctr As Integer
Dim Autohide As Boolean
Dim CommandMovesLists
Dim blit, bbit
Const DATAPATH1 = "C:\WINDOWS\MSAGENT\CHARS\MERLIN.ACS"
Const DATAPATH2 = "C:\WINDOWS\MSAGENT\CHARS\GENIE.ACS"
Const DATAPATH3 = "C:\WINDOWS\MSAGENT\CHARS\PEEDY.ACS"
Const DATAPATH4 = "C:\WINDOWS\MSAGENT\CHARS\ROBBY.ACS"

Private Sub AboutButton_Click()
    Dim xMsg
    xMsg = MsgBox("Puppeteer Version 1.0" & vbCrLf & _
                "This is developed by Walter A. Narvasa" & _
                "Copyright (c), 2001" & _
                "All rights reserved", vbOKOnly + vbInformation, "About")
End Sub

Private Sub Form_Load()
    ' Hide the menu bar
    HideBar
    Autohide = True
    Moderator = "Merlin"
    Call GUIfx("Merlin", DisplayMerlinButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Genie", DisplayGenieButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Peedy", DisplayPeedyButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Robby", DisplayRobbyButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("All", DisplayAllButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Browse", BackgroundButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Merlin", PuppetMerlinButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Genie", PuppetGenieButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Peedy", PuppetPeedyButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Robby", PuppetRobbyButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Edit", EditButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Save", SaveButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Tell Story", TellStoryButton, 0, 0, 1155, 495, Source, 0, 0, 120, 285, "Arial", 9, False)
    Call GUIfx("Shut Down", ShutdownButton, 0, 0, 1155, 495, Source, 0, 32, 120, 285, "Arial", 9, False)
    Call GUIfx("Minimize", MinimizeButton, 0, 0, 1155, 495, Source, 0, 32, 120, 285, "Arial", 9, False)
    Call GUIfx("Help Guide", HelpButton, 0, 0, 1155, 495, Source, 0, 32, 120, 285, "Arial", 9, False)
    Call GUIfx("About", AboutButton, 0, 0, 1155, 495, Source, 0, 32, 120, 285, "Arial", 9, False)
    Call GUIfx("Puppeteer Version 1.0 - by Walter A. Narvasa", DisplayMenu, 0, 0, 1050, 350, Source, 0, 95, 17, 375, "Arial", 12, True)
    Call GUIfx("Puppet Command Moves:", PuppetCommandsGesturesBar, 0, 0, 1050, 550, Source, 0, 415, 26, 38, "Arial", 9, True)
    Call GUIfx("Please ask some questions: (English Only)", QuestionAskBar, 0, 0, 1050, 350, Source, 0, 810, 280, 550, "Arial", 9, True)
    MyAgent.Characters.Load "CharacterID", DATAPATH1
    Set Character = MyAgent.Characters("CharacterID")
    Character.LanguageID = &H409
    PuppetCommandMovesLists.Clear
    For Each CommandMovesLists In Character.AnimationNames
            PuppetCommandMovesLists.AddItem CommandMovesLists
    Next
    Call DisplayMerlinButton_Click
    GenieActivate = False
    PeedyActivate = False
    RobbyActivate = False
End Sub

Private Sub Form_Activate()
    Me.Picture = picTemp(0).Picture
    Call InitializeControls
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ' Size the menu to fit the screen
    DisplayMenu.Width = ScaleWidth
    ' Make sure it stays at the left edge
    DisplayMenu.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Program wouldn't close right away while the menu _
        was folding
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Shown And Autohide Then HideBar
End Sub

Private Sub DisplayMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the menu's hidden then show it
    If Not Shown Then ShowBar
End Sub

Sub HideBar()
    ' Make sure the border is hidden
    ' Set menubar flag
    Shown = False
    ' Make the menu "fold up" but leave a bit visible _
        so we can access the menu, looks ugly with _
        a border around the menu
    For ctr = DisplayMenu.Top To ((-1 * DisplayMenu.Height) + 60) Step -5
        DisplayMenu.Top = ctr
        ' This ensures that it shows the menu "folding"
        DoEvents
    Next ctr
End Sub

Sub ShowBar()
    ' Make sure the border is hidden
    Shown = True
    ' Make the menu "fold down"
    For ctr = DisplayMenu.Top To 0 Step 5
        DisplayMenu.Top = ctr
        ' This ensures that it shows the menu "folding"
        DoEvents
    Next ctr
End Sub

Private Sub BackgroundButton_Click()
    On Error Resume Next
    Dialog.Filter = "Puppeteer Version New Background Picture (*.bmp)|*.bmp"
    Dialog.FilterIndex = 1
    Dialog.ShowOpen
    Call WriteINI("MAIN", "NewBackground", Dialog.filename, App.Path + "\Settings\Puppeteer.INI")
    NewBackground.Text = Dialog.filename
    Me.Picture = LoadPicture(Dialog.filename)
End Sub

Private Sub DisplayTopPosition_LostFocus()
    Call WriteINI("MAIN", "DisplayTopPosition", DisplayTopPosition.Text, App.Path + "\Settings\Puppeteer.INI")
End Sub

Private Sub DisplayLeftPosition_LostFocus()
    Call WriteINI("MAIN", "DisplayLeftPosition", DisplayLeftPosition.Text, App.Path + "\Settings\Puppeteer.INI")
End Sub

Private Sub DisplayMerlinButton_Click()
    If MerlinActivate = True Then
        MsgBox "Merlin already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        MerlinActivate = True
        Agent(0).Characters.Load "Merlin", DATAPATH1
        Set Merlin = Agent(0).Characters("Merlin")
        Merlin.LanguageID = &H409
        Merlin.Show
        Call AddMerlinCommands
        Merlin.MoveTo 34, 109
        Merlin.Speak "Welcome to Puppeteer Version 1.0 !"
        Merlin.Speak "I am a Puppet and can be ask, commanded, and even tell stories."
        Merlin.Speak "You can ask me any question you want even my love life."
        Merlin.Speak "But ask me in a nice manner and in English please because i can't undestand any language but English only, Ok!"
        Merlin.Speak "Have fun with me because i can do magic tricks and can disappear forever."
        Merlin.MoveTo 350, 522
        Merlin.Play "Surprised"
    End If
End Sub

Private Sub DisplayGenieButton_Click()
    If GenieActivate = True Then
        MsgBox "Genie already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        GenieActivate = True
        Agent(1).Characters.Load "Genie", DATAPATH2
        Set Genie = Agent(1).Characters("Genie")
        Genie.LanguageID = &H409
        Genie.Show
        Call AddGenieCommands
        Genie.MoveTo 34, 109
        Genie.Speak "Hi! I am Genie, one of the buddy of Merlin."
        Genie.Speak "I am also a Puppet and can be ask, commanded, and even tell stories."
        Genie.Speak "You can also ask me any question you want even my love life."
        Genie.Speak "But ask me in a nice manner and in English please because i can't undestand any language but English only, Ok!"
        Genie.Speak "Have fun with me because your wish is my command and grant you three wishes."
        Genie.MoveTo 400, 522
        Genie.Play "Surprised"
    End If
End Sub

Private Sub DisplayPeedyButton_Click()
    If PeedyActivate = True Then
        MsgBox "Peedy already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        PeedyActivate = True
        Agent(2).Characters.Load "Peedy", DATAPATH3
        Set Peedy = Agent(2).Characters("Peedy")
        Peedy.LanguageID = &H409
        Peedy.Show
        Call AddPeedyCommands
        Peedy.MoveTo 34, 109
        Peedy.Speak "Hi! I am Peedy, one of the buddy of Merlin."
        Peedy.Speak "I am also a Puppet and can be ask, commanded, and even tell stories."
        Peedy.Speak "You can also ask me any question you want even my love life."
        Peedy.Speak "But ask me in a nice manner and in English please because i can't undestand any language but English only, Ok!"
        Peedy.Speak "Have fun with me because i can fly, talk, laugh, eat, and sing."
        Peedy.MoveTo 450, 522
        Peedy.Play "Surprised"
    End If
End Sub

Private Sub DisplayRobbyButton_Click()
    If RobbyActivate = True Then
        MsgBox "Robby already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        RobbyActivate = True
        Agent(3).Characters.Load "Robby", DATAPATH4
        Set Peedy = Agent(3).Characters("Robby")
        Robby.LanguageID = &H409
        Robby.Show
        Call AddRobbyCommands
        Robby.MoveTo 34, 109
        Robby.Speak "Hi! I am Robby, one of the buddy of Merlin."
        Robby.Speak "I am also a Puppet and can be ask, commanded, and even tell stories."
        Robby.Speak "You can also ask me any question you want even my love life."
        Robby.Speak "But ask me in a nice manner and in English please because i can't undestand any language but English only, Ok!"
        Robby.Speak "Have fun with me because i am strong as steel and can do no man cannot do."
        Robby.MoveTo 500, 522
        Robby.Play "Surprised"
    End If
End Sub

Private Sub DisplayAllButton_Click()
    If MerlinActivate = True Then
        MsgBox "Merlin already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        Call DisplayMerlinButton_Click
    End If
    If GenieActivate = True Then
        MsgBox "Genie already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        Call DisplayGenieButton_Click
    End If
    If PeedyActivate = True Then
        MsgBox "Peedy already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        Call DisplayPeedyButton_Click
    End If
    If RobbyActivate = True Then
        MsgBox "Robby already activated!", vbOKOnly + vbCritical, "Puppeteer - Warning:"
    Else
        Call DisplayRobbyButton_Click
    End If
End Sub

Private Sub PuppetCommandMovesLists_Click()
    If Moderator = "Merlin" Then
        Merlin.Play PuppetCommandMovesLists.List(PuppetCommandMovesLists.ListIndex)
    ElseIf Moderator = "Genie" Then
        Genie.Play PuppetCommandMovesLists.List(PuppetCommandMovesLists.ListIndex)
    ElseIf Moderator = "Peedy" Then
        Peedy.Play PuppetCommandMovesLists.List(PuppetCommandMovesLists.ListIndex)
    ElseIf Moderator = "Robby" Then
        Robby.Play PuppetCommandMovesLists.List(PuppetCommandMovesLists.ListIndex)
    End If
End Sub

Private Sub PuppetCommandsGesturesBarExit_Click()
    PuppetCommandsGesturesBar.Visible = False
    Check1(1).Value = 0
End Sub

Private Sub PuppetGenieButton_Click()
    Moderator = "Genie"
End Sub

Private Sub PuppetMerlinButton_Click()
    Moderator = "Merlin"
End Sub

Private Sub PuppetPeedyButton_Click()
    Moderator = "Peedy"
End Sub

Private Sub PuppetRobbyButton_Click()
    Moderator = "Robby"
End Sub

Private Sub QuestionAskBarExit_Click()
    QuestionAskBar.Visible = False
    Check1(0).Value = 0
End Sub

Private Sub EditButton_Click()
    StoryBox.Enabled = True
    StoryBox.SetFocus
End Sub

Private Sub SaveButton_Click()
    Call WriteINI("MAIN", "StoryBox", StoryBox.Text, App.Path + "\Settings\Puppeteer.INI")
    StoryBox.Enabled = False
End Sub

Private Sub TellStoryButton_Click()
    HideBar
    Autohide = True
    StoryBox.Enabled = False
    If Moderator = "Merlin" Then
        If MerlinActivate = True Then
            Merlin.Play "Announce"
            Merlin.Play "Read"
            Merlin.Speak "I will now start to read a story so all of you be quiet and listen to my wonderful story."
            Merlin.Speak StoryBox.Text
            Merlin.Speak "Thank you for listening to my wonderful story."
            Merlin.Play "Greet"
        End If
    ElseIf Moderator = "Genie" Then
        If GenieActivate = True Then
            Genie.Play "Announce"
            Genie.Play "Read"
            Genie.Speak "I will now start to read a story so all of you be quiet and listen to my wonderful story."
            Genie.Speak StoryBox.Text
            Genie.Speak "Thank you for listening to my wonderful story."
            Genie.Play "Greet"
        End If
    ElseIf Moderator = "Peedy" Then
        If PeedyActivate = True Then
            Peedy.Play "Announce"
            Peedy.Play "Read"
            Peedy.Speak "I will now start to read a story so all of you be quiet and listen to my wonderful story."
            Peedy.Speak StoryBox.Text
            Peedy.Speak "Thank you for listening to my wonderful story."
            Peedy.Play "Greet"
        End If
    ElseIf Moderator = "Robby" Then
        If RobbyActivate = True Then
            Robby.Play "Announce"
            Robby.Play "Read"
            Robby.Speak "I will now start to read a story so all of you be quiet and listen to my wonderful story."
            Robby.Speak StoryBox.Text
            Robby.Speak "Thank you for listening to my wonderful story."
            Robby.Play "Greet"
        End If
    End If
End Sub

Private Sub Agent_Click(Index As Integer, ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    If Button = vbLeftButton Then
        If Index = 0 Then
                Moderator = "Merlin"
                Merlin.Play "Surprised"
                Merlin.Speak "I am the current Puppet that will answer to your question.|Be careful with that pointer!|You can also ask questions to other Puppets by clicking at them.|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to what you are doing!|Are you insane?|Are you crazy?|What is some matter with you?"
                Merlin.Play "RestPose"
        ElseIf Index = 1 Then
                Moderator = "Genie"
                Genie.Play "Surprised"
                Genie.Speak "I am the current Puppet that will answer to your question.|Be careful with that pointer!|You can also ask questions to other Puppets by clicking at them.|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to what you are doing!|Are you insane?|Are you crazy?|What is some matter with you?"
                Genie.Play "RestPose"
        ElseIf Index = 2 Then
                Moderator = "Peedy"
                Peedy.Play "Surprised"
                Peedy.Speak "I am the current Puppet that will answer to your question.|Be careful with that pointer!|You can also ask questions to other Puppets by clicking at them.|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to what you are doing!|Are you insane?|Are you crazy?|What is some matter with you?"
                Peedy.Play "RestPose"
        ElseIf Index = 3 Then
                Moderator = "Robby"
                Robby.Play "Surprised"
                Robby.Speak "I am the current Puppet that will answer to your question.|Be careful with that pointer!|You can also ask questions to other Puppets by clicking at them.|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to what you are doing!|Are you insane?|Are you crazy?|What is some matter with you?"
                Robby.Play "RestPose"
        End If
    End If
End Sub

Private Sub Agent_Command(Index As Integer, ByVal UserInput As Object)
    If Index = 0 Then
        Select Case UserInput.Name
                Case "what"
                        Merlin.Speak AI("what")
                Case "how"
                        Merlin.Speak AI("how")
                Case "where"
                        Merlin.Speak AI("where")
                Case "why"
                        Merlin.Speak AI("why")
                Case "which"
                        Merlin.Speak AI("which")
                Case "what"
                        Merlin.Speak AI("what")
                Case "who"
                        Merlin.Speak AI("who")
                Case "when"
                        Merlin.Speak AI("when")
                Case "doyou"
                        Merlin.Speak AI("do you")
                Case "goodbye"
                        Merlin.Speak "You said Good bye to Merlin"
                Case "exit"
                        Merlin.Speak "Merlin is ready to exit, Bye!"
        End Select
    ElseIf Index = 1 Then
        Select Case UserInput.Name
                Case "what"
                        Genie.Speak AI("what")
                Case "how"
                        Genie.Speak AI("how")
                Case "where"
                        Genie.Speak AI("where")
                Case "why"
                        Genie.Speak AI("why")
                Case "which"
                        Genie.Speak AI("which")
                Case "what"
                        Genie.Speak AI("what")
                Case "who"
                        Genie.Speak AI("who")
                Case "when"
                        Genie.Speak AI("when")
                Case "doyou"
                        Genie.Speak AI("do you")
                Case "goodbye"
                        Genie.Speak "You said Good bye to Genie"
                Case "exit"
                        Genie.Speak "Genie is ready to exit, Bye!"
        End Select
    ElseIf Index = 2 Then
        Select Case UserInput.Name
                Case "what"
                        Peedy.Speak AI("what")
                Case "how"
                        Peedy.Speak AI("how")
                Case "where"
                        Peedy.Speak AI("where")
                Case "why"
                        Peedy.Speak AI("why")
                Case "which"
                        Peedy.Speak AI("which")
                Case "what"
                        Peedy.Speak AI("what")
                Case "who"
                        Peedy.Speak AI("who")
                Case "when"
                        Peedy.Speak AI("when")
                Case "doyou"
                        Peedy.Speak AI("do you")
                Case "goodbye"
                        Peedy.Speak "You said Good bye to Peedy"
                Case "exit"
                        Peedy.Speak "Peedy is ready to exit, Bye!"
        End Select
    ElseIf Index = 3 Then
        Select Case UserInput.Name
                Case "what"
                        Robby.Speak AI("what")
                Case "how"
                        Robby.Speak AI("how")
                Case "where"
                        Robby.Speak AI("where")
                Case "why"
                        Robby.Speak AI("why")
                Case "which"
                        Robby.Speak AI("which")
                Case "what"
                        Robby.Speak AI("what")
                Case "who"
                        Robby.Speak AI("who")
                Case "when"
                        Robby.Speak AI("when")
                Case "doyou"
                        Robby.Speak AI("do you")
                Case "goodbye"
                        Robby.Speak "You said Good bye to Robby"
                Case "exit"
                        Robby.Speak "Robby is ready to exit, Bye!"
        End Select
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    If Index = 0 Then
        If Check1(0).Value = 1 Then
            QuestionAskBar.Visible = True
            QuestionAskText.Enabled = True
            QuestionAskText.SetFocus
        Else
            QuestionAskBar.Visible = False
            QuestionAskText.Enabled = False
        End If
    ElseIf Index = 1 Then
        If Check1(1).Value = 1 Then
            PuppetCommandsGesturesBar.Visible = True
        Else
            PuppetCommandsGesturesBar.Visible = False
        End If
    ElseIf Index = 2 Then
        If Check1(2).Value = 1 Then
            DisplayTopPosition.Enabled = True
            DisplayLeftPosition.Enabled = True
            DisplayTopPosition.SetFocus
        Else
            DisplayTopPosition.Enabled = False
            DisplayLeftPosition.Enabled = False
        End If
    ElseIf Index = 3 Then
        If Check1(3).Value = 1 Then
            BackgroundButton.Enabled = True
            Call WriteINI("MAIN", "ChangeBackground", "True", App.Path + "\Settings\Puppeteer.INI")
        Else
            BackgroundButton.Enabled = False
            Call WriteINI("MAIN", "ChangeBackground", "False", App.Path + "\Settings\Puppeteer.INI")
        End If
    End If
End Sub

Function OldQuestion(UseText As String) As Boolean
    Dim i
    For i = 0 To Used.ListCount - 1
        If UCase(Used.List(i)) = UCase(UseText) Then
            OldQuestion = True
            Exit Function
        Else
            OldQuestion = False
        End If
    Next i
End Function


Function OkQuestion(TheText As String)
    Dim TempText As String
    Dim Ekstra As String
    Dim Text(0 To 15) As String
    Dim Number As Integer
    If InStr(1, TheText, "elvis", vbTextCompare) Then GoTo Theking
        TempText = Replace(TheText, " ", "")
    If TheText = TempText Then
        Ekstra = ""
    Else
        Ekstra = " "
    End If
Start:
    If InStr(1, TheText, "What" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "How" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "Where" & Ekstra, vbTextCompare) Then GoTo Where
    If InStr(1, TheText, "Why" & Ekstra, vbTextCompare) Then GoTo Why
    If InStr(1, TheText, "Which" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "Who" & Ekstra, vbTextCompare) Then GoTo Who
    If InStr(1, TheText, "When" & Ekstra, vbTextCompare) Then GoTo When
    Text(0) = "I think so"
    Text(1) = "What the hell do you think ?"
    Text(2) = "Yeah"
    Text(3) = "Nope"
    Text(4) = "Comeone .. evryone knows the answer is NO"
    Text(5) = "Yepper"
    Text(6) = "Not as far as i know"
    Text(7) = "Yes"
    Text(8) = "No"
    Text(9) = "Comeone .. evryone knows the answer is Yes"
    Text(10) = "Yeah Right?"
    Text(11) = "Are you kiddin' me?"
    Text(12) = "Don't try to make me laugh!"
    Text(13) = "Ok then, Will you further more explain to me?"
    Text(14) = "You have delayed reaction on what you are saying!"
    Text(15) = "Don't ask me!, Ask your Dog"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
WhichWhatHow:
    Text(0) = "How the hell should i know ?"
    Text(1) = "It's getting late ... ask me tomorrow"
    Text(2) = "Ehhhhhh"
    Text(3) = "Weeelllll"
    Text(4) = "Thats funny ..... someone just asked me that lately"
    Text(5) = "Damn thats's a good question"
    Text(6) = "You would like to know that i think ;)"
    Text(7) = "I won't tell you that"
    Text(8) = "Damn i forgot"
    Text(9) = "I wish you knew how much i hate curious people"
    Text(10) = "Are you sure with your question?"
    Text(11) = "Until further notice my friend!"
    Text(12) = "Are you making me feel guilty?"
    Text(13) = "Just as what everybody does"
    Text(14) = "I will answer that in the near future OK!"
    Text(15) = "That is pretty personal question, I may say"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Where:
    Text(0) = "Wasn't that in Europe ?"
    Text(1) = "I think it's somehwere around Africa"
    Text(2) = "Just a second .. i'll go get the atlas"
    Text(3) = "In Germany"
    Text(4) = "Right up in heaven"
    Text(5) = "In your grandmoms backyard"
    Text(6) = "I think it was in the tomato-soup we got last friday"
    Text(7) = "It's in the whitehouse"
    Text(8) = "Last time i cheked it was in my underwear"
    Text(9) = "Maybe in my new washmachine ?"
    Text(10) = "In my bed"
    Text(11) = "Right down in Hell!!"
    Text(12) = "At you place"
    Text(13) = "Your Dog House!"
    Text(14) = "Right in front of you Bath Room"
    Text(15) = "At the lake"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
Exit Function
Why:
    Text(0) = "Because you ask all those stupid questions"
    Text(1) = "Check in the encyclopedia"
    Text(2) = "It was my destiny"
    Text(3) = "Ehmmmm"
    Text(4) = "because that's the way i wanted it"
    Text(5) = "Ask your president"
    Text(6) = "Weellllll"
    Text(7) = "I don't remember"
    Text(8) = "Come on ask me a relevant question"
    Text(9) = "Why don't you ask your mom ?"
    Text(10) = "I have amnesia right now can't remember anything!!"
    Text(11) = "because you try to kiss my ass"
    Text(12) = "I don't care!! The hell with you!!"
    Text(13) = "because we are meant for each other"
    Text(14) = "because you always makes me horny??!!!"
    Text(15) = "I beg your pardon, Please come again?"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Who:
    Text(0) = "Don't you think it was Donald Duck"
    Text(1) = "Michael Jordan"
    Text(2) = "Jim Carrey"
    Text(3) = "A strange man in the middle of the 60's"
    Text(4) = "Mrs. Beth Logan "
    Text(5) = "It was your dentist"
    Text(6) = "Brad Pitt"
    Text(7) = "Your Tatay"
    Text(8) = "Erap"
    Text(9) = "Your Lola"
    Text(10) = "Dennis Rodman"
    Text(11) = "Rosanna Roces"
    Text(12) = "Ara Mina"
    Text(13) = "Your Mother, Ok!"
    Text(14) = "Your Senator Tessie Oreta"
    Text(15) = "Cindy Crawford"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
When:
    Text(0) = "Tomorrow"
    Text(1) = "Yesterday"
    Text(2) = "It was in year 1900"
    Text(3) = "When you are old enough to hear it"
    Text(4) = "Damn it's a long time ago"
    Text(5) = "In the Edsa Revolution"
    Text(6) = "When the day you were born"
    Text(7) = "The first day your father changed his underwear"
    Text(8) = "I't was in the same period when Erap was thrown away from his seat ........"
    Text(9) = "Now"
    Text(10) = "The day after tomorrow"
    Text(11) = "Last Saturday"
    Text(12) = "Tomorrow Evening!"
    Text(13) = "Bloody Monday!!, Ok!"
    Text(14) = "Since the first day I ate your dog."
    Text(15) = "Last Friday, when I fell in the manhole!"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Theking:
    If InStr(1, TheText, "alive", vbTextCompare) Or InStr(1, TheText, "living", vbTextCompare) Or InStr(1, TheText, "dead", vbTextCompare) Then
        OkQuestion = "What the hell do you think ? Offcourse he's alive, remember he's the king ;)"
    Else
        GoTo Start
    End If
End Function

Function NoQuestion()
    Dim Text(0 To 15) As String
    Dim Number As Integer
    Text(0) = "Ehh... why just try with a question ;)"
    Text(1) = "I'm better to answer questions"
    Text(2) = "Wee OK ... well what about asking me a question"
    Text(3) = "I love questions with sense only!"
    Text(4) = "Pllzz ask me a question"
    Text(5) = "I'm only a Puppet .... ask me a question"
    Text(6) = "Try me with a question ;)"
    Text(7) = "OK, but ask me a question now"
    Text(8) = "Let's try a with an intelligent question, ok!"
    Text(9) = "I'm only good at questions, understand!"
    Text(10) = "Are you running out of questions?"
    Text(11) = "Maybe you are sick of me thats why you don't want to ask me some questions"
    Text(12) = "I think you are dumb thats why you can't ask me a question!"
    Text(13) = "I'm sick and tired of what you are saying, just ask me questions please!"
    Text(14) = "Questions only or else i will kick you out off my Puppet World!!"
    Text(15) = "I'm sure you are not only dumb but silly also, ask me a good question!"
    Randomize
    Number = Int((Rnd * 16) + 1) - 1
    NoQuestion = Text(Number)
End Function

Function AI(Text As String)
    Dim TempText As String
    Dim Ekstra As String
    TempText = Replace(Text, " ", "")
    If OldQuestion(TempText) = False Then
        Used.AddItem TempText
        If Text = TempText Then
            Ekstra = ""
        Else
            Ekstra = " "
        End If
        If InStr(1, Text, "What" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "How" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Where" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Why" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Which" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Who" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "When" & Ekstra, vbTextCompare) Then GoTo Question
        If Right(TempText, 1) = "?" Then GoTo Question
        AI = NoQuestion
        Exit Function
Question:
        AI = OkQuestion(Text)
    Else
        AI = "I already answered you on that.."
    End If
End Function
 
Private Sub ShutdownButton_Click()
    End
End Sub

Private Sub QuestionAskText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Moderator = "Merlin" Then
            If MerlinActivate = True Then
                Merlin.Play "Think"
                Merlin.Speak AI(QuestionAskText.Text)
            End If
        ElseIf Moderator = "Genie" Then
            If GenieActivate = True Then
                Genie.Play "Think"
                Genie.Speak AI(QuestionAskText.Text)
            End If
        ElseIf Moderator = "Peedy" Then
            If PeedyActivate = True Then
                Peedy.Play "Think"
                Peedy.Speak AI(QuestionAskText.Text)
            End If
        ElseIf Moderator = "Robby" Then
            If RobbyActivate = True Then
                Robby.Play "Think"
                Robby.Speak AI(QuestionAskText.Text)
            End If
        End If
        QuestionAskText.Text = ""
    End If
End Sub

Private Sub AddMerlinCommands()
    'Add voice commands to Merlin
    Merlin.Commands.Add "what", "what", "what", True, True
    Merlin.Commands.Add "how", "how", "how", True, True
    Merlin.Commands.Add "where", "where", "where", True, True
    Merlin.Commands.Add "why", "why", "why", True, True
    Merlin.Commands.Add "which", "which", "which", True, True
    Merlin.Commands.Add "who", "who", "who", True, True
    Merlin.Commands.Add "when", "when", "when", True, True
    Merlin.Commands.Add "doyou", "do you", "do you", True, True
    Merlin.Commands.Add "goodbye", "Good By", "Good By", True, True
    Merlin.Commands.Add "exit", "exit", "exit", True, True
End Sub

Private Sub AddGenieCommands()
    'Add voice commands to Genie
    Genie.Commands.Add "what", "what", "what", True, True
    Genie.Commands.Add "how", "how", "how", True, True
    Genie.Commands.Add "where", "where", "where", True, True
    Genie.Commands.Add "why", "why", "why", True, True
    Genie.Commands.Add "which", "which", "which", True, True
    Genie.Commands.Add "who", "who", "who", True, True
    Genie.Commands.Add "when", "when", "when", True, True
    Genie.Commands.Add "doyou", "do you", "do you", True, True
    Genie.Commands.Add "goodbye", "Good By", "Good By", True, True
    Genie.Commands.Add "exit", "exit", "exit", True, True
End Sub

Private Sub AddPeedyCommands()
    'Add voice commands to Peedy
    Peedy.Commands.Add "what", "what", "what", True, True
    Peedy.Commands.Add "how", "how", "how", True, True
    Peedy.Commands.Add "where", "where", "where", True, True
    Peedy.Commands.Add "why", "why", "why", True, True
    Peedy.Commands.Add "which", "which", "which", True, True
    Peedy.Commands.Add "who", "who", "who", True, True
    Peedy.Commands.Add "when", "when", "when", True, True
    Peedy.Commands.Add "doyou", "do you", "do you", True, True
    Peedy.Commands.Add "goodbye", "Good By", "Good By", True, True
    Peedy.Commands.Add "exit", "exit", "exit", True, True
End Sub

Private Sub AddRobbyCommands()
    'Add voice commands to Robby
    Robby.Commands.Add "what", "what", "what", True, True
    Robby.Commands.Add "how", "how", "how", True, True
    Robby.Commands.Add "where", "where", "where", True, True
    Robby.Commands.Add "why", "why", "why", True, True
    Robby.Commands.Add "which", "which", "which", True, True
    Robby.Commands.Add "who", "who", "who", True, True
    Robby.Commands.Add "when", "when", "when", True, True
    Robby.Commands.Add "doyou", "do you", "do you", True, True
    Robby.Commands.Add "goodbye", "Good By", "Good By", True, True
    Robby.Commands.Add "exit", "exit", "exit", True, True
End Sub

Private Sub InitializeControls()
    On Error Resume Next
    Dim sngStartTime As Single, BackgroundLogic As String
    sngStartTime = Timer
    Do Until (Timer - sngStartTime) > 32
          DoEvents
    Loop
    StoryBox.Text = ReadINI("MAIN", "StoryBox", App.Path + "\Settings\Puppeteer.INI")
    DisplayTopPosition.Text = ReadINI("MAIN", "DisplayTopPosition", App.Path + "\Settings\Puppeteer.INI")
    DisplayLeftPosition.Text = ReadINI("MAIN", "DisplayLeftPosition", App.Path + "\Settings\Puppeteer.INI")
    BackgroundLogic = ReadINI("MAIN", "ChangeBackground", App.Path + "\Settings\Puppeteer.INI")
    If BackgroundLogic = "True" Then
        Check1(3).Value = 1
        NewBackground.Text = ReadINI("MAIN", "NewBackground", App.Path + "\Settings\Puppeteer.INI")
        picTemp(1).Picture = LoadPicture(Trim(NewBackground.Text))
    Else
        Check1(3).Value = 0
    End If
    DisplayMenu.Visible = True
    Me.Picture = picTemp(1).Picture
End Sub
