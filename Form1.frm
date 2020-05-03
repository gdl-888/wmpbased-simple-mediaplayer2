VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "심플 미디어 2 플레이어"
   ClientHeight    =   9000
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   15360
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   9000
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows 기본값
   Begin ComctlLib.Slider sldVolume 
      Height          =   5415
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "소리 크기 조절 막대"
      Top             =   1800
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   9551
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   25
      SmallChange     =   10
      Max             =   100
      SelStart        =   50
      TickStyle       =   3
      Value           =   50
   End
   Begin ComctlLib.Slider sldPosision 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   7200
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   873
      _Version        =   327682
      BorderStyle     =   1
      Max             =   32767
      TickFrequency   =   10
   End
   Begin VB.Timer timRotPlay 
      Interval        =   125
      Left            =   13320
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8822&
      Caption         =   "↗"
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      Top             =   7200
      Width           =   615
   End
   Begin VB.Timer timPosAndStatusChanger 
      Interval        =   1000
      Left            =   12960
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9340
      _Version        =   393216
      TabOrientation  =   3
      Tab             =   1
      TabHeight       =   2284
      ShowFocusRect   =   0   'False
      BackColor       =   14855538
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":1C6224
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "WMP"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":1C6240
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "comDrive"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstDirList"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lstFileList"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtPattern"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":1C625C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "Label6"
      Tab(2).Control(5)=   "Label5"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "보안"
         Height          =   855
         Left            =   -74760
         TabIndex        =   15
         Top             =   2400
         Width           =   11295
         Begin VB.Label Label12 
            Caption         =   "재생한 음악 기록하지 않기: 아니오"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   4815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "재생"
         Height          =   855
         Left            =   -74760
         TabIndex        =   14
         Top             =   1320
         Width           =   11295
         Begin VB.Label Label1 
            Caption         =   "반복: 아니오"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "화면표시"
         Height          =   855
         Left            =   -74760
         TabIndex        =   13
         Top             =   240
         Width           =   11295
         Begin VB.Label Label2 
            Caption         =   "애니메이션 효과: 예"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "열기"
         Height          =   255
         Left            =   9840
         TabIndex        =   10
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txtPattern 
         Height          =   270
         Left            =   3600
         TabIndex        =   9
         Text            =   "*.flv;*.flac;*.mov;*.wma;*.wmv;*.avi;*.mp4;*.mp3;*.mid;*.midi;*.rmi;*.wav;*.mp3;*.mp2;*.mp1;*.mpe;*.mpg;*.mpeg;*.snd"
         Top             =   4800
         Width           =   6135
      End
      Begin VB.FileListBox lstFileList 
         Height          =   4050
         Left            =   3600
         Pattern         =   "*.flv;*.flac;*.mov;*.wma;*.wmv;*.avi;*.mp4;*.mp3;*.mid;*.midi;*.rmi;*.wav;*.mp3;*.mp2;*.mp1;*.mpe;*.mpg;*.mpeg;*.snd"
         TabIndex        =   8
         Top             =   600
         Width           =   7815
      End
      Begin VB.DirListBox lstDirList 
         Height          =   4500
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.DriveListBox comDrive 
         Height          =   300
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   11175
      End
      Begin WMPLibCtl.WindowsMediaPlayer WMP 
         Height          =   5040
         Left            =   -74880
         TabIndex        =   29
         Top             =   120
         Width           =   11460
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "none"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   20214
         _cy             =   8890
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "설정"
         Height          =   255
         Left            =   12240
         TabIndex        =   27
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '투명
         Caption         =   "열기"
         Height          =   255
         Left            =   -62760
         TabIndex        =   26
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '투명
         Caption         =   "열기"
         Height          =   255
         Left            =   -62760
         TabIndex        =   25
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '투명
         Caption         =   "열기"
         Height          =   255
         Left            =   12240
         TabIndex        =   24
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "설정"
         Height          =   255
         Left            =   -62760
         TabIndex        =   23
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "설정"
         Height          =   255
         Left            =   -62760
         TabIndex        =   22
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "지금 재생"
         Height          =   255
         Left            =   -63000
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "지금 재생"
         Height          =   255
         Left            =   12000
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "지금 재생"
         Height          =   255
         Left            =   -63000
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Image imgPause 
      Height          =   855
      Left            =   1680
      ToolTipText     =   "일시 중지"
      Top             =   7800
      Width           =   855
   End
   Begin VB.Image imgPlay 
      Height          =   1095
      Left            =   240
      ToolTipText     =   "재생"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Image imgNext 
      Height          =   855
      Left            =   4800
      ToolTipText     =   "다음 항목"
      Top             =   7950
      Width           =   975
   End
   Begin VB.Image imgPrev 
      Height          =   855
      Left            =   3840
      ToolTipText     =   "이전 항목"
      Top             =   7920
      Width           =   975
   End
   Begin VB.Image imgStop 
      Height          =   855
      Left            =   2760
      ToolTipText     =   "중지"
      Top             =   7920
      Width           =   855
   End
   Begin VB.Image imgClearHis 
      Height          =   495
      Left            =   480
      ToolTipText     =   "기록 지우기"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   13800
      Top             =   480
      Width           =   495
   End
   Begin VB.Image imgGuard 
      Height          =   495
      Index           =   5
      Left            =   14640
      Top             =   7320
      Width           =   735
   End
   Begin VB.Image imgGuard 
      Height          =   735
      Index           =   4
      Left            =   13920
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Image imgGuard 
      Height          =   495
      Index           =   3
      Left            =   5760
      Top             =   8520
      Width           =   9615
   End
   Begin VB.Image imgGuard 
      Height          =   495
      Index           =   2
      Left            =   14880
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image imgGuard 
      Height          =   495
      Index           =   1
      Left            =   13800
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image imgGuard 
      Height          =   495
      Index           =   0
      Left            =   1320
      Top             =   0
      Width           =   12660
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgPlayR 
      Height          =   1200
      Index           =   3
      Left            =   240
      Picture         =   "Form1.frx":1C6278
      Stretch         =   -1  'True
      Top             =   7800
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgPlayR 
      Height          =   1200
      Index           =   2
      Left            =   240
      Picture         =   "Form1.frx":1CE37A
      Stretch         =   -1  'True
      Top             =   7800
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgPlayR 
      Height          =   1200
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":1D647C
      Stretch         =   -1  'True
      Top             =   7800
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgPlayR 
      Height          =   1200
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":1DE85E
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1485
   End
   Begin VB.Label lblTime 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   11
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label lblHistory 
      BackStyle       =   0  '투명
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   11655
   End
   Begin VB.Label lblCaption1 
      BackStyle       =   0  '투명
      Caption         =   "재생한 음악:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgSuperVol 
      Height          =   975
      Left            =   360
      ToolTipText     =   "최대 크기"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image imgHighVol 
      Height          =   1455
      Left            =   360
      ToolTipText     =   "큰 소리"
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image imgMedVol 
      Height          =   615
      Left            =   360
      ToolTipText     =   "적당한 크기"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Image imgSmallVol 
      Height          =   1815
      Left            =   360
      ToolTipText     =   "작은 소리"
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image imgMute 
      Height          =   615
      Left            =   360
      ToolTipText     =   "조용히"
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   7920
      Width           =   6015
   End
   Begin VB.Image imgIconView 
      Height          =   615
      Left            =   14400
      ToolTipText     =   "아이콘 표시"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgClose 
      Height          =   615
      Left            =   14880
      ToolTipText     =   "닫기"
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu filem 
      Caption         =   "파일(&F)"
      Begin VB.Menu openm 
         Caption         =   "열기(&O)"
      End
      Begin VB.Menu closem 
         Caption         =   "닫기(&C)"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu exitm 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu configm 
      Caption         =   "설정(&S)"
      Begin VB.Menu loopm 
         Caption         =   "반복(&L)"
      End
      Begin VB.Menu anim 
         Caption         =   "애니메이션 효과(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu dishism 
         Caption         =   "기록 안 함(&H)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
 ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Public cnt As Integer
Public ini As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Sub anim_Click()
    On Error Resume Next
    anim.Checked = Not anim.Checked
    If anim.Checked = True Then
        Label2.Caption = "애니메이션 효과: 예"
        timRotPlay.Enabled = True
        imgPlayR(0).Visible = True
        imgPlayR(1).Visible = False
        imgPlayR(2).Visible = False
        imgPlayR(3).Visible = False
    Else
        timRotPlay.Enabled = False
        Check1.Value = 0
        Label2.Caption = "애니메이션 효과: 아니오"
    End If
End Sub

Private Sub closem_Click()
    On Error Resume Next
    ini = True
    Me.Caption = "심플 미디어 2 플레이어"
    WMP.URL = "c:\windows\media\onestop.mid"
    WMP.Controls.Stop
End Sub

Private Sub comDrive_Change()
    On Error Resume Next
    lstDirList.Path = comDrive.Drive
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    ini = False
    WMP.URL = lstFileList.Path & "\" & lstFileList.FileName
    SSTab1.Tab = 0
    If Len(lblHistory.Caption) > 350 Then lblHistory.Caption = ""
    
    If dishism.Checked = False Then
        If lblHistory.Caption <> "" Then
            lblHistory.Caption = lblHistory.Caption & "           " & lstFileList.FileName
        Else
            lblHistory.Caption = lstFileList.FileName
        End If
    End If
    
    cnt = cnt + 1
    If cnt >= 5 Then
        'MsgBox "알 수 없는 오류로 5번 이상 불러올 경우 오류가 발생할 수 있습니다.", 48, "경고"
        'Unload Me
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    WMP.fullScreen = True
End Sub

Private Sub dishism_Click()
    On Error Resume Next
    dishism.Checked = Not dishism.Checked
    If dishism.Checked = True Then
        Label12.Caption = "재생한 음악 기록하지 않기: 예"
        lblHistory.Caption = ""
    Else
        Label12.Caption = "재생한 음악 기록하지 않기: 아니오"
    End If
End Sub

Private Sub exitm_Click()
    On Error Resume Next
    imgClose_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    cnt = 0
    sldVolume = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Settings", "Volume", 50)
    WMP.settings.volume = 100 - GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Settings", "Volume", 50)
    comDrive.Drive = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Logs", "FileBrowse\Drive", comDrive.Drive)
    lstDirList.Path = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Logs", "FileBrowse\Path", lstDirList.Path)
    WMP.settings.setMode "loop", GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Config", "Loop", False)
    loopm.Checked = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Config", "Loop", False)
    anim.Checked = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Config", "Animation", True)
    dishism.Checked = GetSetting("PowerPoint Based Designing Media Player (PPTMP)", "Config", "DisHistory", False)
    If anim.Checked = True Then
        timRotPlay.Enabled = True
        imgPlayR(0).Visible = True
        imgPlayR(1).Visible = False
        imgPlayR(2).Visible = False
        imgPlayR(3).Visible = False
        Label2.Caption = "애니메이션 효과: 예"
    Else
        timRotPlay.Enabled = False
        Label2.Caption = "애니메이션 효과: 아니오"
    End If
    
    If loopm.Checked = True Then
        Label1.Caption = "반복: 예"
    Else
        Label1.Caption = "반복: 아니오"
    End If
    
    If dishism.Checked = True Then
        Label12.Caption = "재생한 음악 기록하지 않기: 예"
        lblHistory.Caption = ""
    Else
        Label12.Caption = "재생한 음악 기록하지 않기: 아니오"
    End If
    
    If 100 - sldVolume.Value >= 75 Then
        timRotPlay.Interval = 75
    ElseIf 100 - sldVolume.Value >= 50 Then
        timRotPlay.Interval = 125
    ElseIf 100 - sldVolume.Value >= 25 Then
        timRotPlay.Interval = 160
    ElseIf 100 - sldVolume.Value >= 1 Then
        timRotPlay.Interval = 250
    ElseIf 100 - sldVolume.Value = 0 Then
        timRotPlay.Interval = 0
    End If
    ini = True
    WMP.URL = "c:\windows\media\onestop.mid"
    WMP.Controls.Stop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim ReturnValue As Long
    If Button = 1 Then
    Call ReleaseCapture
    ReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Config", "Loop", loopm.Checked
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Settings", "Volume", sldVolume.Value
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Logs", "FileBrowse\Path", lstDirList.Path
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Logs", "FileBrowse\Drive", comDrive.Drive
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Config", "Animation", anim.Checked
    SaveSetting "PowerPoint Based Designing Media Player (PPTMP)", "Config", "DisHistory", dishism.Checked
End Sub

Private Sub imgClearHis_Click()
    On Error Resume Next
    lblHistory.Caption = ""
End Sub

Private Sub imgClose_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub imgHighVol_Click()
    On Error Resume Next
    WMP.settings.volume = 75
    sldVolume.Value = 100 - WMP.settings.volume
    sldVolume_Scroll
End Sub

Private Sub imgIconView_Click()
    On Error Resume Next
    Me.WindowState = 1
End Sub

Private Sub imgMedVol_Click()
    On Error Resume Next
    WMP.settings.volume = 50
    sldVolume.Value = 100 - WMP.settings.volume
    sldVolume_Scroll
End Sub

Private Sub imgMute_Click()
    On Error Resume Next
    WMP.settings.volume = 0
    sldVolume.Value = 100 - WMP.settings.volume
    sldVolume_Scroll
End Sub

Private Sub imgNext_Click()
    On Error Resume Next
    WMP.Controls.Next
End Sub

Private Sub imgPause_Click()
    On Error Resume Next
    WMP.Controls.pause
End Sub

Private Sub imgPlay_Click()
    On Error Resume Next
    WMP.Controls.Play
End Sub

Private Sub imgPrev_Click()
    On Error Resume Next
    WMP.Controls.Previous
End Sub

Private Sub imgSmallVol_Click()
    On Error Resume Next
    WMP.settings.volume = 25
    sldVolume.Value = 100 - WMP.settings.volume
    sldVolume_Scroll
End Sub

Private Sub imgStop_Click()
    On Error Resume Next
    WMP.Controls.Stop
End Sub

Private Sub imgSuperVol_Click()
    On Error Resume Next
    WMP.settings.volume = 100
    sldVolume.Value = 100 - WMP.settings.volume
    sldVolume_Scroll
End Sub

Private Sub Label1_Click()
    loopm_Click
End Sub

Private Sub Label10_Click()
    SSTab1.Tab = 1
End Sub

Private Sub Label11_Click()
    SSTab1.Tab = 2
End Sub

Private Sub Label12_Click()
    dishism_Click
End Sub

Private Sub Label2_Click()
    anim_Click
End Sub

Private Sub Label4_Click()
    SSTab1.Tab = 0
End Sub

Private Sub Label5_Click()
    SSTab1.Tab = 0
End Sub

Private Sub Label7_Click()
    SSTab1.Tab = 2
End Sub

Private Sub Label9_Click()
    SSTab1.Tab = 1
End Sub

Private Sub loopm_Click()
    On Error Resume Next
    loopm.Checked = Not loopm.Checked
    If loopm.Checked = True Then
        WMP.settings.setMode "loop", True
        Label1.Caption = "반복: 예"
    Else
        WMP.settings.setMode "loop", False
        Label1.Caption = "반복: 아니오"
    End If
End Sub

Private Sub lstDirList_Change()
    On Error Resume Next
    lstFileList.Path = lstDirList.Path
End Sub

Private Sub openm_Click()
    On Error Resume Next
    SSTab1.Tab = 1
End Sub

Private Sub sldPosision_Scroll()
    On Error Resume Next
    WMP.Controls.currentPosition = sldPosision.Value
End Sub

Private Sub sldVolume_Scroll()
    On Error Resume Next
    WMP.settings.volume = 100 - sldVolume.Value
    If 100 - sldVolume.Value >= 75 Then
        timRotPlay.Interval = 75
    ElseIf 100 - sldVolume.Value >= 50 Then
        timRotPlay.Interval = 125
    ElseIf 100 - sldVolume.Value >= 25 Then
        timRotPlay.Interval = 160
    ElseIf 100 - sldVolume.Value >= 1 Then
        timRotPlay.Interval = 250
    ElseIf 100 - sldVolume.Value = 0 Then
        timRotPlay.Interval = 0
    End If
End Sub

Private Sub timPosAndStatusChanger_Timer()
    On Error Resume Next
    sldPosision.Value = WMP.Controls.currentPosition
    lblStatus.Caption = WMP.Status
    lblTime.Caption = WMP.Controls.currentPositionString
    If ini = False Then
        Me.Caption = WMP.Status & " - " & WMP.currentMedia.Name & " - " & "심플 미디어 2 플레이어 " & "(" & WMP.Controls.currentPositionString & ")"
    End If
End Sub

Private Sub timRotPlay_Timer()
    On Error Resume Next
    If WMP.playState = wmppsPlaying Then
        If imgPlayR(0).Visible = True Then
            imgPlayR(0).Visible = False
            imgPlayR(1).Visible = True
            imgPlayR(2).Visible = False
            imgPlayR(3).Visible = False
        ElseIf imgPlayR(1).Visible = True Then
            imgPlayR(0).Visible = False
            imgPlayR(1).Visible = False
            imgPlayR(2).Visible = True
            imgPlayR(3).Visible = False
        ElseIf imgPlayR(2).Visible = True Then
            imgPlayR(0).Visible = False
            imgPlayR(1).Visible = False
            imgPlayR(2).Visible = False
            imgPlayR(3).Visible = True
        Else
            imgPlayR(0).Visible = True
            imgPlayR(1).Visible = False
            imgPlayR(2).Visible = False
            imgPlayR(3).Visible = False
        End If
    End If
End Sub

Private Sub txtPattern_Change()
    On Error Resume Next
    lstFileList.Pattern = txtPattern.Text
End Sub

Private Sub txtPattern_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Def
    If KeyCode = 13 Then
        lstFileList.Pattern = txtPattern.Text
    End If
    Exit Sub
    
Def:
    lstFileList.Pattern = "*.flv;*.flac;*.mov;*.wma;*.wmp;*.avi;*.mp4;*.mp3;*.mid;*.midi;*.rmi;*.wav;*.mp3;*.mp2;*.mp1;*.mpe;*.mpg;*.mpeg;*.snd"
End Sub

Private Sub WMP_MediaChange(ByVal Item As Object)
    On Error Resume Next
    sldPosision.Max = WMP.currentMedia.duration
End Sub

