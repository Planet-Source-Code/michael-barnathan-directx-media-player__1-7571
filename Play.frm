VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectMedia Player"
   ClientHeight    =   3555
   ClientLeft      =   5160
   ClientTop       =   3450
   ClientWidth     =   4905
   Icon            =   "Play.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4905
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1680
      Top             =   720
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      SelectRange     =   -1  'True
      TickStyle       =   3
      TickFrequency   =   0
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1200
      Picture         =   "Play.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton CMD_Play 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "Play.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton CMD_PauseMusic 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Picture         =   "Play.frx":0E4E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton CMD_Stop 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Picture         =   "Play.frx":1304
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   560
      Left            =   3240
      Picture         =   "Play.frx":17BA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      DragIcon        =   "Play.frx":2544
      DragMode        =   1  'Automatic
      Height          =   320
      Left            =   4200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   145
   End
   Begin MSComDlg.CommonDialog CDialog_Open 
      Left            =   -5000
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4
   End
   Begin MSComDlg.CommonDialog CDialog_DLS 
      Left            =   -5000
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   -5000
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   -5000
      Top             =   1200
   End
   Begin VB.Frame Frame_SegmentInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4890
      Begin MSComCtl2.UpDown UpDown_Volume 
         Height          =   315
         Left            =   -5000
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   50
         BuddyControl    =   "EDIT_Volume"
         BuddyDispid     =   196625
         OrigLeft        =   4320
         OrigTop         =   360
         OrigRight       =   4560
         OrigBottom      =   735
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 / 00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3720
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Stopped"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label EDIT_Volume 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -5000
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   900
         Width           =   735
      End
      Begin VB.Label LBL_Tempo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label LBL_Name 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2955
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LBL_Length 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Length: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   735
      End
      Begin VB.Label LBL_TimeSig 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Sig:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   900
         Width           =   735
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   3840
      X2              =   4680
      Y1              =   960
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   4680
      X2              =   4680
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   4680
      X2              =   3840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label ElapsedTime 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Menu MNU_File 
      Caption         =   "&File"
      Begin VB.Menu MNU_Open 
         Caption         =   "&Open..."
      End
      Begin VB.Menu MNU_Play 
         Caption         =   "&Play"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_Pause 
         Caption         =   "P&ause"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_Stop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
      End
      Begin VB.Menu MNU_Options_GMReset 
         Caption         =   "GM &Reset!"
      End
      Begin VB.Menu MNU_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MNU_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnuDLS 
         Caption         =   "Connect to DLS"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal lMilliseconds As Long)

Dim Dragging As Boolean
Dim MySeconds

Dim dx As New DirectX7
Dim perf As DirectMusicPerformance
Dim perf2 As DirectMusicPerformance
Dim seg As DirectMusicSegment
Dim segstate As DirectMusicSegmentState
Dim loader As DirectMusicLoader
Dim col As DirectMusicCollection
Public GetStartTime As Long
Public Offset As Long
Public mtTime As Long
Public mtLength As Double
Public dTempo As Double
Dim timesig As DMUS_TIMESIGNATURE
Dim portcaps As DMUS_PORTCAPS
Dim IsPlayingCheck As Boolean
Dim msg As String
Dim time As Double
Dim Offset2 As Long
Dim ElapsedTime2 As Double
Dim fIsPaused As Boolean
Private Enum CONST_DMUS_SEGF_FLAGS
    DMUS_SEGF_AFTERPREPARETIME = 1024
    DMUS_SEGF_BEAT = 4096
    DMUS_SEGF_CONTROL = 512
    DMUS_SEGF_DEFAULT = 16384
    DMUS_SEGF_GRID = 2048
    DMUS_SEGF_MEASURE = 8192
    DMUS_SEGF_NOINVALIDATE = 32768
    DMUS_SEGF_QUEUE = 256
    DMUS_SEGF_REFTIME = 64
    DMUS_SEGF_SECONDARY = 128
End Enum
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Sub localerror(ErrorNum As Long, ErrorDesc As String)
    msg = ErrorDesc
    msg = "(" & ErrorNum & ") - " & msg
    MsgBox msg
End Sub

Private Sub CMD_PauseMusic_Click()
On Error GoTo LocalErrors

    If seg Is Nothing Then
        CMD_PauseMusic.BackColor = &H8000000F 'gray
        Exit Sub
    End If

    IsPlayingCheck = perf.IsPlaying(seg, segstate)
    If IsPlayingCheck = True Then 'music is playing
        fIsPaused = True
        ' pause music and button down
        mtTime = perf.GetMusicTime()
        GetStartTime = segstate.GetStartTime()
        Call perf.Stop(seg, Nothing, 0, 0)
        CMD_PauseMusic.BackColor = &HFFFFC0 'blue
    Else
        If CMD_PauseMusic.BackColor = &HFFFFC0 Then 'button is blue
            'unpause
            fIsPaused = False
            Offset = mtTime - GetStartTime + Offset + 1
            Call seg.SetStartPoint(Offset)
            Set segstate = perf.PlaySegment(seg, 0, 0)
            CMD_PauseMusic.BackColor = &H8000000F 'gray
            Sleep (90)
        End If
    End If
Exit Sub
LocalErrors:
    Call localerror(Err.Number, Err.Description)
End Sub

Private Sub CMD_PauseMusic_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub CMD_Play_Click()
    If seg Is Nothing Then
        MsgBox ("Please open a segment or MIDI file before playing")
        Exit Sub
    End If
    
    If fIsPaused Then
        Offset = mtTime - GetStartTime + Offset + 1
        Call seg.SetStartPoint(Offset)
        Set segstate = perf.PlaySegment(seg, 0, 0)
        CMD_PauseMusic.BackColor = &H8000000F 'gray
        Sleep (90)
    Else
        Offset = 0
        If perf.IsPlaying(seg, segstate) = True Then
            Call perf.Stop(seg, segstate, 0, 0)
        End If
        seg.SetStartPoint (0)
        Set segstate = perf.PlaySegment(seg, 0, 0)
        CMD_PauseMusic.BackColor = &H8000000F 'gray
        Sleep (90)
        Exit Sub
    End If
    fIsPaused = False
End Sub

Private Sub CMD_Play_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub CMD_Stop_Click()
    If seg Is Nothing Then
        Exit Sub
    End If
    
    fIsPaused = False
    CMD_PauseMusic.BackColor = &H8000000F
    Call perf.Stop(seg, segstate, 0, 0)
    CMD_Play.Enabled = True
    MNU_Play.Enabled = True
    time = 0
    ElapsedTime = vbNullString
    Slider1.Value = 1
End Sub



Private Sub CMD_Stop_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Command1_Click()
    MNU_Open_Click
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Command3.Tag = "PRESSED" Then UpDown_Volume.Value = UpDown_Volume.Tag: Command2.Enabled = True: Command3.Tag = "": Command3.Picture = LoadPicture(App.path & "\sound.bmp") Else: Command3.Picture = LoadPicture(App.path & "\sound2.bmp"): Command3.Tag = "PRESSED": UpDown_Volume.Tag = UpDown_Volume.Value: UpDown_Volume.Value = 0: Command2.Enabled = False
End Sub

Private Sub EDIT_Volume_Change()
    perf.SetMasterVolume (EDIT_Volume.Caption * 42 - 3000)
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If Source = Command2 Then Dragging = True
If Source = Command2 And Y <> 700 Then Source.Top = 700
If Source = Command2 And X < Line2.X2 Then Source.Left = Line2.X2
If Source = Command2 And X > Line2.X1 Then Source.Left = Line2.X1
If Source = Command2 And X >= Line2.X2 And X <= Line2.X1 Then Source.Left = X: UpDown_Volume.Value = (100 / (Line2.X1 - Line2.X2)) * (X - Line2.X2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 37 Then UpDown_Volume.Value = UpDown_Volume.Value - 5: Command2.Left = Line2.X2 + (Line2.X1 - Line2.X2) * (UpDown_Volume.Value / 100)
If KeyCode = 39 Then UpDown_Volume.Value = UpDown_Volume.Value + 5: Command2.Left = Line2.X2 + (Line2.X1 - Line2.X2) * (UpDown_Volume.Value / 100)
End Sub

Private Sub Form_Load()
    On Error GoTo LocalErrors
    
    Set loader = dx.DirectMusicLoaderCreate()
    
        
    'Creating a Perf2 so that we can get all the segment information without having to play the segment
    Set perf2 = dx.DirectMusicPerformanceCreate()
    Call perf2.Init(Nothing, 0)
    perf2.SetPort -1, 80
    Call perf2.GetMasterAutoDownload

    Set perf = dx.DirectMusicPerformanceCreate()
    Call perf.Init(Nothing, 0)
    perf.SetPort -1, 80
    Call perf.SetMasterAutoDownload(True)
    perf.SetMasterVolume (EDIT_Volume.Caption * 42 - 3000)
    EDIT_Volume.Caption = UpDown_Volume.Value
    Timer1.Enabled = True
Exit Sub
LocalErrors:
    Call localerror(Err.Number, Err.Description)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dragging = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (perf Is Nothing) Then perf.CloseDown
End Sub

Private Sub MNU_Exit_Click()    ' Exit the app
    Unload Me
End Sub


Private Sub MNU_Open_Click() ' This is where we load new segments and get their info
Dim name As String
Dim Minutes As Integer
Dim a As Integer
Dim length As Integer
Dim length2 As Integer
Static sCurdir As String
Static lFilter As Long

On Error GoTo LocalErrors
   
If Not seg Is Nothing And Not segstate Is Nothing Then ' There is a Segment and a SegmentState
    If perf.IsPlaying(seg, segstate) = True Then ' Segment currently playing, so exit
        MsgBox "Please Stop currently playing music before selecting a new segment"
        Exit Sub
    ElseIf CMD_PauseMusic.BackColor = &HFFFFC0 Then ' Segment currently paused, so exit
        MsgBox "Please Stop currently playing music before selecting a new segment"
        Exit Sub
    End If
End If
   
    Set loader = Nothing
    Set loader = dx.DirectMusicLoaderCreate
    CDialog_Open.Filter = "Segment Files (*.sgt)|*.sgt|MIDI Files (*.mid)|*.mid"   ' Set filters
    If lFilter = 0 Then
        CDialog_Open.FilterIndex = 2   ' Specify default filter
    Else
        CDialog_Open.FilterIndex = lFilter   ' Specify default filter
    End If
    CDialog_Open.filename = vbNullString
    If sCurdir = vbNullString Then
        'Set the init folder to \windows\media if it exists.  If not, set it to the \windows folder
        Dim sWindir As String
        sWindir = Space$(255)
        If GetWindowsDirectory(sWindir, 255) = 0 Then
            'We couldn't get the windows folder for some reason, use the c:\
            CDialog_Open.InitDir = "C:\"
        Else
            Dim sMedia As String
            sWindir = Left$(sWindir, InStr(sWindir, Chr$(0)) - 1)
            If Right$(sWindir, 1) = "\" Then
                sMedia = sWindir & "Media"
            Else
                sMedia = sWindir & "\Media"
            End If
            If Dir$(sMedia, vbDirectory) <> vbNullString Then
                CDialog_Open.InitDir = sMedia
            Else
                CDialog_Open.InitDir = sWindir
            End If
        End If
    Else
        CDialog_Open.InitDir = sCurdir
    End If
    CDialog_Open.ShowOpen   ' Display the Open dialog box

    If CDialog_Open.filename <> vbNullString Then 'The operation was not Canceled
        Set seg = loader.LoadSegment(CDialog_Open.filename)
        CMD_Play.Enabled = True
        MNU_Play.Enabled = True
    Else
        Exit Sub
    End If
    sCurdir = GetFolder(CDialog_Open.filename)
    If LCase(Right$(CDialog_Open.filename, 4)) = ".sgt" Then
        lFilter = 1
    Else
        lFilter = 2
    End If
    
    ' Set the search directory based on the placement of the .sgt file that was loaded
    length = Len(CDialog_Open.filename)
    length2 = length
    Dim path As String
    Do While path <> "\"
        path = Mid(CDialog_Open.filename, length, 1)
        length = length - 1
    Loop
    Dim SearchDir As String
    SearchDir = Left(CDialog_Open.filename, length)
    loader.SetSearchDirectory (Left(CDialog_Open.filename, length + 1))
    perf2.SetMasterAutoDownload True
    
    'Set all the Captions to empty
    LBL_Tempo.Caption = vbNullString
    LBL_TimeSig.Caption = vbNullString
    LBL_Name.Caption = vbNullString
    LBL_Length.Caption = vbNullString
    
    'Get Name
    length = Len(CDialog_Open.filename)
    length2 = length
    Do While name <> "\"
        name = Mid(CDialog_Open.filename, length, 1)
        length = length - 1
    Loop
    
    LBL_Name.Caption = Right(CDialog_Open.filename, length2 - (length + 1))
    
    'Play the segment just long enough to get the info
    mtTime = perf2.GetMusicTime()
    Call perf2.PlaySegment(seg, 0, mtTime + 2000)
    
    'GetTempo
    dTempo = perf2.GetTempo(mtTime + 2000, 0)
    LBL_Tempo.Caption = Format(dTempo, "00.00")
    
    'GetTimeSig
    Call perf2.GetTimeSig(mtTime + 2000, 0, timesig)
    LBL_TimeSig.Caption = timesig.beatsPerMeasure & "/" & timesig.beat
    'GetLength
    mtLength = (((seg.GetLength() / 768) * 60) / dTempo)
    ' Put the length in a time that we can relate to
    Minutes = 0
    a = mtLength - 60
    Do While a > 0
        Minutes = Minutes + 1
        a = a - 60
    Loop
    LBL_Length.Caption = Format(Minutes, "00") & ":" & Format((mtLength - (Minutes * 60)), "00.0")
    Label6.Tag = Format(Minutes, "00") & ":" & Format((mtLength - (Minutes * 60)), "00")
    Slider1.Max = Minutes * 600
    Slider1.min = 1
    Label6.Caption = "00:00 / " & Format(Minutes, "00") & ":" & Format((mtLength - (Minutes * 60)), "00")
    ' Now that we retreived all the segment info, we'll stop playing the segment
    Call perf2.Stop(seg, Nothing, 0, 0)
        
    If LCase(Right$(CDialog_Open.filename, 4)) = ".mid" Then
        seg.SetStandardMidiFile
    End If
Exit Sub
LocalErrors:
If Not seg Is Nothing Then
    Call perf2.Stop(seg, Nothing, 0, 0)
End If
    If Format(Right(LBL_Name.Caption, 4), "<") = ".sgt" Then
        MsgBox ("There was a problem gathering all information about the .sgt file that you loaded.  It may be because the files that it references are not located in the same directory (" & SearchDir & ").  However, the segment may still play")
    ElseIf Format(Right(LBL_Name.Caption, 4), "<") = ".mid" Then
        MsgBox ("There was a problem loading the requested MIDI file.  Please try a different file")
    Else
        MsgBox ("There was a problem loading the requested file.  No file has been loaded")
        CDialog_Open.filename = vbNullString
    End If
    
End Sub

Private Sub MNU_Options_GMReset_Click()
On Local Error GoTo localerror
    Call perf.Reset(0)
    
Exit Sub
localerror:
    If Err.Number = 445 Then ' This is a known issue...
        MsgBox "Currently unable to Reset to General MIDI"
    End If
End Sub

Private Function GetFolder(ByVal sFile As String) As String
    Dim lCount As Long
    
    For lCount = Len(sFile) To 1 Step -1
        If Mid$(sFile, lCount, 1) = "\" Then
            GetFolder = Left$(sFile, lCount)
            Exit Function
        End If
    Next
    GetFolder = vbNullString
End Function

Private Sub MNU_Pause_Click()
    CMD_PauseMusic_Click
End Sub

Private Sub MNU_Play_Click()
    CMD_Play_Click
End Sub


Private Sub MNU_Stop_Click()
    CMD_Stop_Click
End Sub




Private Sub mnuDLS_Click()
On Error GoTo LocalErrors
    If seg Is Nothing Then
        MsgBox "You first need to load a segment before you can connect a dls collection to it"
        Exit Sub
    End If

   ' Set filters.
   CDialog_DLS.Filter = "DLS Collections (*.dls)|*.dls"
   ' Specify default filter.
   CDialog_DLS.FilterIndex = 2

   ' Display the Open dialog box.
   CDialog_DLS.ShowOpen
   
    If CDialog_DLS.filename <> vbNullString Then
        Set col = loader.LoadCollection(CDialog_DLS.filename)
        If perf.IsPlaying(seg, segstate) = True Then
            Call perf.Stop(seg, segstate, 0, 0)
            CMD_Play.Enabled = True
            MNU_Play.Enabled = True
        End If
        If col Is Nothing Then
            MsgBox "Unable to Load the specified collection"
            Exit Sub
        Else
            Call seg.ConnectToCollection(col)
        End If
    Else
        Exit Sub
    End If
    Exit Sub
LocalErrors:
    Call localerror(Err.Number, Err.Description)

End Sub



Private Sub Slider1_Change()
If Slider1.Tag = "SCROLLING" Then
Slider1.Tag = ""
fIsPaused = False
Call seg.SetStartPoint(mtLength / Slider1.Max * Slider1.Value)
Set segstate = perf.PlaySegment(seg, 0, 0)
Sleep (90)
End If
Rem mtLength = (((seg.GetLength() / 768) * 60) / dTempo)
End Sub

Private Sub Slider1_Scroll()
Exit Sub 'We got a problem here, so don't go any further
Slider1.Tag = "SCROLLING"
End Sub

Private Sub Timer1_Timer()
    Timer2.Enabled = True
    Timer1.Enabled = False
    GetTime
End Sub

Private Sub Timer2_Timer()
    Timer1.Enabled = True
    Timer2.Enabled = False
    GetTime
End Sub

Private Sub GetTime()
    Dim min As Integer
    Dim a As Single

    ' if we don't have a SegmentState (or Performance), we don't want to check the time
    If segstate Is Nothing Or perf Is Nothing Then
        Exit Sub
    End If
    
    '''''''''''PAUSED
    If CMD_PauseMusic.BackColor = &HFFFFC0 Then 'blue
        Label5.Caption = "Paused"
    
    
    '''''''''''PLAYING
    ElseIf perf.IsPlaying(Nothing, segstate) = True Then
        Label5.Caption = "Playing"
    ' Calculate The time
    ' Calculate in Ticks
        ' First, we'll get the time in raw seconds
        ElapsedTime2 = ((((perf.GetMusicTime() - (segstate.GetStartTime() - Offset)) / 768) * 60) / dTempo)
    
        ' Next, we'll calculate minutes
        min = 0
        a = ElapsedTime2 - 60
        Do While a >= 0
            min = min + 1
            a = a - 60
        Loop
        ' Finally, we'll print out the time with the proper format
        ElapsedTime = Format(min, "00") & ":" & Format(Abs((ElapsedTime2 - (min * 60))), "00")
        MySeconds = min * 60 + Abs((ElapsedTime2 - (min * 60)))
    CMD_Stop.Enabled = True
    MNU_Stop.Enabled = True
    CMD_PauseMusic.Enabled = True
    MNU_Pause.Enabled = True

        
    
    '''''''''''STOPPED
    Else
        Label5.Caption = "Stopped"
        
 
            ElapsedTime = "00:00"

        CMD_Stop.Enabled = False
        MNU_Stop.Enabled = False
        CMD_PauseMusic.Enabled = False
        MNU_Pause.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
If Label6.Tag <> "" And ElapsedTime <> "" And Label5.Caption = "Playing" Then Slider1.Enabled = True: Select Case Slider1.Tag: Case Is <> "SCROLLING": Slider1.Value = MySeconds * 10: End Select: Label6.Caption = ElapsedTime & " / " & Label6.Tag
End Sub

Private Sub UpDown_Volume_Change()
    EDIT_Volume.Caption = UpDown_Volume.Value
End Sub
