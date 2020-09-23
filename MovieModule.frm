VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Movie Module Example v1.2"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Slider Rate 
      Height          =   135
      Left            =   3120
      TabIndex        =   29
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   238
      _Version        =   393216
      Max             =   100
      SelStart        =   50
      Value           =   50
   End
   Begin MSComctlLib.Slider Volume 
      Height          =   135
      Left            =   240
      TabIndex        =   26
      Top             =   3360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   238
      _Version        =   393216
      Max             =   100
      SelStart        =   85
      Value           =   85
   End
   Begin MSComctlLib.Slider H 
      Height          =   135
      Left            =   240
      TabIndex        =   25
      Top             =   3840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   238
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Multimedia Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   6015
      Begin VB.Label Label7 
         Caption         =   "Position:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label L11 
         Caption         =   "Play Rate:"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label L12 
         Caption         =   "Volume:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Movie Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   6015
      Begin VB.Timer T 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3960
         Top             =   600
      End
      Begin VB.Frame P 
         Height          =   2655
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   5775
      End
      Begin MSComDlg.CommonDialog C 
         Left            =   3960
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pause"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Get/Set - Size/Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton Command7 
         Caption         =   "Get Size"
         Height          =   975
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Set Size"
         Height          =   975
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtLeft 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtTop 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "0"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtWidth 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtHeight 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "0"
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Multmedia Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Label L7 
         BackStyle       =   0  'Transparent
         Caption         =   "Length: "
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label L6 
         BackStyle       =   0  'Transparent
         Caption         =   "Position: "
         Height          =   255
         Left            =   1680
         TabIndex        =   36
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label L8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label L10 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame Rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label L9 
         BackStyle       =   0  'Transparent
         Caption         =   "Frames Total:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label L3 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   375
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   2760
      End
      Begin VB.Label L4 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Status: "
         Height          =   1095
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label L1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label L2 
         BackStyle       =   0  'Transparent
         Caption         =   "Frame: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label L5 
         BackStyle       =   0  'Transparent
         Caption         =   "FramesPerSec: "
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   2520
      TabIndex        =   33
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MM As New MovieModule

Private Sub Command1_Click()
    On Error Resume Next
    MM.playMovie
    MM.setVolume Volume.Value * 10 ' set the new movie the selected volume
    MM.setSpeed Rate.Value * 20 'set the new movie to the selected speed
    T.Enabled = True 'set our timer on
    H.Max = Val(MM.getLengthInSec) 'load the position bar with the max length
    MM.timeOut 0.5 'Give the mci device enough time to process
    
    L7.Caption = "Length: " & MM.getFormatLength
    L8.Caption = "Total Time: " & MM.getLengthInMS
    L9.Caption = "Total Frames: " & MM.getLengthInFrames
    L3.Caption = "Status: " & MM.getStatus

    L4.Caption = "Error Status: " & MM.checkError 'Check for a error during the last process

End Sub
Private Sub Command2_Click()
    MM.stopMovie
    T.Enabled = False
    MM.timeOut 1 'give our mci device time to update the status
    L3.Caption = "Status: " & MM.getStatus
    L4.Caption = "Error Status: " & MM.checkError
End Sub
Private Sub Command3_Click()
    If Command3.Caption = "Pause" Then
        MM.pauseMovie
        Command3.Caption = "Resume"
        T.Enabled = False
    Else
        MM.resumeMovie
        Command3.Caption = "Pause"
        T.Enabled = True
    End If
    MM.timeOut 1 'give our mci device time to update the status
    L3.Caption = "Status: " & MM.getStatus
    L4.Caption = "Error Status: " & MM.checkError
End Sub
Private Sub Command4_Click()
    Dim a As Long
    Dim b As Long
'open and set the filename
    C.Filter = "Avi Files (*.avi)|*.avi|Mpeg Files (*.mpeg)|*.mpeg|Mpg Files (*.mpg)|*.mpg|Mov Files (*.mov)|*.mov|All Files (*.*)|*.*"
    C.ShowOpen
    MM.Filename = C.Filename
'open the movie
    MM.openMovieWindow P.hWnd, "child" 'this will open our movie in a child window
    'MM.openMovie 'use this function to open in a popup window
'clear the previous filename
    C.Filename = ""
'fill in status and size information
    L3.Caption = "Status: " & MM.getStatus
    L4.Caption = "Error Status: " & MM.checkError
    MM.extractDefaultMovieSize a, b
    txtWidth.Text = CStr(a)
    txtHeight.Text = CStr(b)
End Sub

Private Sub Command5_Click()
    MM.closeMovie 'close the mci device
    L3.Caption = "Status: " & MM.getStatus
    L4.Caption = "Error Status: " & MM.checkError
    T.Enabled = False
End Sub

Private Sub Command6_Click()
    MM.sizeLocateMovie Val(txtLeft.Text), Val(txtTop.Text), Val(txtWidth.Text), Val(txtHeight.Text)
    L4.Caption = "Error Status: " & MM.checkError
End Sub

Private Sub Command7_Click()
    Dim cWidth As Long
    Dim cHeight As Long
    MM.extractCurrentMovieSize 0, 0, cWidth, cHeight
    txtWidth = CStr(cWidth)
    txtHeight = CStr(cHeight)
    L4.Caption = "Error Status: " & MM.checkError
End Sub

Private Sub Form_Load()
    'set defulat volume and rate values
    Volume.Value = 80
    Rate.Value = 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MM.closeMovie
    Unload Me
    End
End Sub

Private Sub H_Click()
    'change the playback position of the movie
    MM.setPositionTo H.Value
End Sub

Private Sub Rate_Click()
    'change the rate the movie is played
    MM.setSpeed Rate.Value * 20
    L11.Caption = "Play Rate: " & Rate.Value * 2 & "%"
End Sub

Private Sub Slider1_Click()

End Sub

Private Sub T_Timer()
    On Error Resume Next
    L1.Caption = "Time: " & MM.getPositionInMS
    
    L2.Caption = "Frame: " & MM.getPositionInFrames 'skips

    L10.Caption = "Frame Rate: " & MM.getNominalFrameRate
    L6.Caption = "Position: " & MM.getFormatPosition
    L5.Caption = "FramesPerSec: " & MM.getFramePerSecRate
    H.Value = MM.getPositionInSec
End Sub

Private Sub Volume_Click()
    MM.setVolume Volume.Value * 10
    L12.Caption = "Volume: " & Volume.Value & "%"
End Sub
