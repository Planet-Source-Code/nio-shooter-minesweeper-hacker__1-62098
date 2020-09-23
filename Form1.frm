VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Minesweeper Hacker v1.1 - by NiO_ShOoTer"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AutoRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   795
   End
   Begin VB.Timer Loader 
      Interval        =   10
      Left            =   375
      Top             =   795
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FD896C&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1290
      Left            =   45
      ScaleHeight     =   1260
      ScaleWidth      =   3960
      TabIndex        =   1
      Top             =   60
      Width           =   3990
      Begin VB.PictureBox picStatistics 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   645
         ScaleHeight     =   1230
         ScaleWidth      =   3300
         TabIndex        =   3
         Top             =   15
         Visible         =   0   'False
         Width           =   3300
         Begin VB.Timer RefreshStatistics 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2580
            Top             =   345
         End
         Begin VB.Timer tmrFreezeTime 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2820
            Top             =   375
         End
         Begin VB.CommandButton cmdFreezeTime 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Freeze"
            Height          =   255
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   345
            Width           =   900
         End
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            Height          =   255
            Left            =   1740
            MaxLength       =   3
            TabIndex        =   4
            Top             =   345
            Width           =   525
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "count"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2310
            TabIndex        =   20
            Top             =   915
            Width           =   405
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "uncovered"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   930
            TabIndex        =   19
            Top             =   915
            Width           =   765
         End
         Begin VB.Label lblCCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2730
            TabIndex        =   18
            Top             =   900
            Width           =   525
         End
         Begin VB.Label lblUCells 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1740
            TabIndex        =   17
            Top             =   900
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cells:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   435
            TabIndex        =   16
            Top             =   915
            Width           =   435
         End
         Begin VB.Label lblMLeft 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFE1CE&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1740
            TabIndex        =   14
            Top             =   630
            Width           =   525
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "left"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1455
            TabIndex        =   13
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mines:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   330
            TabIndex        =   12
            Top             =   645
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last click:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   11
            Top             =   105
            Width           =   825
         End
         Begin VB.Label lblLCol 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFE1CE&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1740
            TabIndex        =   10
            Top             =   75
            Width           =   525
         End
         Begin VB.Label lblLRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFE1CE&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2730
            TabIndex        =   9
            Top             =   75
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Col"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1455
            TabIndex        =   8
            Top             =   105
            Width           =   225
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Row"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2370
            TabIndex        =   7
            Top             =   90
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   405
            TabIndex        =   6
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Now"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1380
            TabIndex        =   5
            Top             =   375
            Width           =   315
         End
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         FillColor       =   &H8000000F&
         Height          =   435
         Left            =   150
         Picture         =   "Form1.frx":08CA
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   435
         Width           =   435
      End
   End
   Begin VB.Label lblCell 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuStartMineweeper 
         Caption         =   "Star Mineweeper"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoRefresh 
         Caption         =   "Auto Refresh"
      End
      Begin VB.Menu mnuAutoCheckMines 
         Caption         =   "Auto Check Mines"
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   "Show Statistics"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTransparency 
         Caption         =   "Transparency (2000/XP)"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Refresh"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim game_title As String
Dim game_class As String
Dim running_game As Boolean

Dim TimeValue As String

Private Sub Form_Load()
    game_title = "Minesweeper"          ' GameWindow Title
    game_class = "Minesweeper"          ' GameWindow Class
    
    picLogo.CurrentY = (picLogo.Height / 3) - 40
    picLogo.CurrentX = 700
    picLogo.Print "Minesweeper Hacker"
End Sub

Private Sub Loader_Timer()

        ' Wait till the user loads the game '
        '-----------------------------------'
        Window_hWnd = FindWindow(game_class, game_title) ' Get the handle
        
        If (Window_hWnd = 0 And running_game = True) Then
            running_game = False: ResetForm True
            ResetAll
        ElseIf (Window_hWnd <> 0 And running_game = False) Then
            running_game = True: ResetForm False
            Call Attach
            If mnuStatistics.Checked = False Then LoadTable
        End If
        '-----------------------------------'
End Sub

Private Sub ResetForm(showlogo As Boolean)
    If mnuStatistics.Checked = False Then picLogo.Visible = showlogo
    Me.width = 4185
    Me.Height = 2040
End Sub

Private Sub AutoRefresh_Timer()
    LoadTable
End Sub

Private Sub mnuAutoCheckMines_Click()
If mnuAutoCheckMines.Checked = True Then
    mnuAutoCheckMines.Checked = False
Else
    mnuAutoCheckMines.Checked = True
    LoadTable
    RefreshWindow
End If
End Sub

Private Sub mnuAutoRefresh_Click()
If mnuAutoRefresh.Checked = True Then
    mnuAutoRefresh.Checked = False
    AutoRefresh.Enabled = False
Else
    mnuAutoRefresh.Checked = True
    AutoRefresh.Enabled = True
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuRefresh_Click()
If mnuStatistics.Checked = False Then
    LoadTable
    If mnuAutoCheckMines.Checked Then RefreshWindow
End If
End Sub

Private Sub mnuStartMineweeper_Click()
    Shell "winmine.exe", vbNormalFocus
End Sub

Private Sub mnuStatistics_Click()
If mnuStatistics.Checked = True Then
    mnuStatistics.Checked = False
    mnuAutoRefresh.Enabled = True
    mnuAutoCheckMines.Enabled = True
    picStatistics.Visible = False
    picLogo.Visible = False: LoadTable
    RefreshStatistics.Enabled = False
    If mnuAutoRefresh.Checked Then AutoRefresh.Enabled = True
Else
    ResetForm True: ResetAll
    mnuStatistics.Checked = True
    AutoRefresh.Enabled = False
    mnuAutoRefresh.Enabled = False
    mnuAutoCheckMines.Enabled = False
    picStatistics.Visible = True
    RefreshStatistics.Enabled = True
End If
End Sub

Private Sub cmdFreezeTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdFreezeTime.Caption = "Freeze" Then
    cmdFreezeTime.Caption = "Freezed"
    TimeValue = txtTime
    tmrFreezeTime.Enabled = True
Else
    cmdFreezeTime.Caption = "Freeze"
    tmrFreezeTime.Enabled = False
End If
txtTime.SetFocus
End Sub

Private Sub mnuTransparency_Click()
If mnuTransparency.Checked = True Then
    mnuTransparency.Checked = False
    TranslucentForm Me, 255
Else
    mnuTransparency.Checked = True
    TranslucentForm Me, 200
End If
End Sub

Private Sub RefreshStatistics_Timer()
Dim PeekValue As Long, width As Integer, heigth As Integer
    'last clicked col
    Peek &H1005118, PeekValue, 1: lblLCol = PeekValue
    'last clicked row
    Peek &H100511C, PeekValue, 1: lblLRow = PeekValue
    'current time
    Peek &H100579C, PeekValue, 4: txtTime = PeekValue
    'mines left
    Peek &H1005194, PeekValue, 4: lblMLeft = PeekValue
    'uncovered cells
    Peek &H10057A4, PeekValue, 1: lblUCells = PeekValue
    'cells count
    Peek &H10056AC, PeekValue, 4: width = PeekValue
    Peek &H10056A8, PeekValue, 4: heigth = PeekValue
    lblCCount = width * heigth
End Sub

Private Sub tmrFreezeTime_Timer()
    Poke &H100579C, CLng(TimeValue), 4
End Sub

Private Sub txtTime_Change()
    If txtTime <> "" Then _
    Poke &H100579C, CLng(txtTime), 4: TimeValue = txtTime
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If Not IsNumber(KeyAscii) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Function IsNumber(KeyAscii As Integer) As Boolean
    If InStr("1234567890", Chr(KeyAscii)) Then IsNumber = True Else IsNumber = False
End Function
