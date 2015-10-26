VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   10920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   3
      Top             =   10920
      Width           =   2775
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   3615
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   8175
      ExtentX         =   14420
      ExtentY         =   6376
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9060
      Left            =   6240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9060
      ScaleWidth      =   14370
      TabIndex        =   0
      Top             =   3600
      Width           =   14370
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   2380
         TabIndex        =   1
         Top             =   1560
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9630
      Left            =   -3720
      Picture         =   "Form1.frx":21921
      ScaleHeight     =   9630
      ScaleWidth      =   14175
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   14175
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   1800
         TabIndex        =   46
         Top             =   7920
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   9000
         TabIndex        =   45
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   8400
         TabIndex        =   44
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   6600
         TabIndex        =   43
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   6120
         TabIndex        =   42
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   5640
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   5160
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   4800
         TabIndex        =   39
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   4440
         TabIndex        =   38
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   11640
         TabIndex        =   37
         Top             =   9000
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   9720
         TabIndex        =   36
         Top             =   9000
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   7320
         TabIndex        =   35
         Top             =   8880
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   6600
         TabIndex        =   34
         Top             =   9120
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   6240
         TabIndex        =   33
         Top             =   8760
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   27
         Left            =   5760
         TabIndex        =   32
         Top             =   9000
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   26
         Left            =   5280
         TabIndex        =   31
         Top             =   8760
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   25
         Left            =   5040
         TabIndex        =   30
         Top             =   8760
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   24
         Left            =   4800
         TabIndex        =   29
         Top             =   8760
         Width           =   255
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   4320
         TabIndex        =   28
         Top             =   8760
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   3960
         TabIndex        =   27
         Top             =   8760
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   3600
         TabIndex        =   26
         Top             =   8760
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   3000
         TabIndex        =   25
         Top             =   9000
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2400
         TabIndex        =   24
         Top             =   8760
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   9000
         TabIndex        =   23
         Top             =   6240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   9000
         TabIndex        =   22
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   10200
         TabIndex        =   21
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   9000
         TabIndex        =   20
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   13
         Left            =   10200
         TabIndex        =   19
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   12
         Left            =   9000
         TabIndex        =   18
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   10200
         TabIndex        =   17
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   9000
         TabIndex        =   16
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   10200
         TabIndex        =   15
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   9000
         TabIndex        =   14
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   7
         Left            =   2520
         TabIndex        =   13
         Top             =   7440
         Width           =   5535
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   6
         Left            =   6840
         TabIndex        =   12
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   6840
         TabIndex        =   11
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   4440
         TabIndex        =   10
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   3
         Left            =   4800
         TabIndex        =   9
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   4800
         TabIndex        =   8
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         Top             =   3120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
If cnt < 7 Then cnt = cnt + 1
Select Case cnt
Case 1
    BrowseMe "ABOUT THE PROJECT.htm"
Case 2
    Form2.Visible = True
    Form1.Visible = False
Case 3
    BrowseMe "INTRODUCTION TO MICROPROCESSOR.htm"
Case 4
    BrowseMe "Evolution.htm"
    Picture1.Visible = False
    Picture2.Visible = False
    Browser.Move 0, 0, ScaleWidth, ScaleHeight - Command1.Height
    Browser.Visible = True
Case 5
    Browser.Width = 6615
    Browser.Height = 3615
    Browser.Visible = False
    Picture1.Move 0, 0
    Picture1.Visible = True
    Picture2.Visible = False
Case 6
Labe2_Reset
    Browser.Width = 6615
    Browser.Height = 3615
    Browser.Visible = False
    Picture2.Move 0, 0
    Picture2.Visible = True
    Picture1.Visible = False
Case 7
    'Command1.Caption = "Loading ..."
    DoEvents
    Form2.Visible = True
    cnt = 2
    'Form2.Label3(3).Visible = False
End Select
End Sub

Private Sub Command2_Click()
If cnt > 1 Then cnt = cnt - 1
Select Case cnt
Case 1
    BrowseMe "ABOUT THE PROJECT.htm"
Case 2
    Form2.Visible = True
    Form1.Visible = False
Case 3
    BrowseMe "INTRODUCTION TO MICROPROCESSOR.htm"
Case 4
    BrowseMe "Evolution.htm"
    Picture1.Visible = False
    Picture2.Visible = False
    Browser.Move 0, 0, ScaleWidth, ScaleHeight - Command1.Height
    Browser.Visible = True
Case 5
Label_Reset
    Browser.Width = 6615
    Browser.Height = 3615
    Browser.Visible = False
    Picture1.Move 0, 0
    Picture1.Visible = True
    Picture2.Visible = False
Case 6
Labe2_Reset
    Browser.Width = 6615
    Browser.Height = 3615
    Browser.Visible = False
    Picture2.Move 0, 0
    Picture2.Visible = True
    Picture1.Visible = False
Case 7
    Command1.Caption = "Loading ..."
    DoEvents
    'Unload Form2
    'Load Form2
    Form2.Visible = True
    cnt = 2
    'Form2.Label3(3).Visible = False
End Select

End Sub

Private Sub Form_Activate()
Me.WindowState = 2
BrowseMe "BRAINWAVES.htm"
Browser.Move 0, 0, ScaleWidth, ScaleHeight
Browser.Visible = True
Picture1.Visible = False
'If loaded1 = True Then Exit Sub
On Error Resume Next
Dim i As Integer
For i = 2 To 20
    Load Label1(i)
    Label1(i).Move Label1(i).Left, Label1(i - 1).Top + Label1(1).Height + 160, Label1(1).Width, Label1(1).Height
    Label1(i).ZOrder 1
    Label1(i).Visible = True
Next

    Load Label1(21)
    Label1(21).Move Label1(1).Left + 1600, Label1(1).Top, Label1(1).Width, Label1(1).Height
    Label1(21).Visible = True
For i = 22 To 40
    Load Label1(i)
    Label1(i).Move Label1(i - 1).Left, Label1(i - 1).Top + Label1(1).Height + 160, Label1(1).Width, Label1(1).Height
    Label1(i).ZOrder 1
    Label1(i).Visible = True
Next
loaded1 = True

End Sub

Private Sub Label1_Click(Index As Integer)
Browser.Move Label1(Index).Left + Label1(Index).Width, Label1(Index).Top
Browser.Visible = True
Select Case Index
    Case 1, 2
         BrowseMe "X1  AND X2.htm"
    Case 3
        BrowseMe "Reset OUT.htm"
        Browser.Height = 2500
    Case 4, 5
        BrowseMe "SERIAL INPUT DATA.htm"
        Browser.Height = 2000
    Case 6
        BrowseMe "TRAP.htm"
    Case 7, 8, 9
        BrowseMe "RESTART INTERRUPTS.htm"
    Case 10
        BrowseMe "INTR.htm"
        Browser.Height = 2500
    Case 11
        BrowseMe "INTA.htm"
        Browser.Width = 8000
    Case 12 To 19
        BrowseMe "DATA BUS.htm"
    Case 20
        BrowseMe "Vss.htm"
        Browser.Height = 2000
    Case 21
        BrowseMe "VCC.htm"
        Browser.Height = 2000
    Case 22
        BrowseMe "HOLD.htm"
    Case 23
        BrowseMe "HLDA.htm"
    Case 24
        BrowseMe "CLK.htm"
    Case 25
        BrowseMe "RESETIN.htm"
    Case 26
        BrowseMe "READY.htm"
    Case 27
        BrowseMe "IO.htm"
    Case 28
        BrowseMe "S1 AND S0.htm"
    Case 29
        BrowseMe "READ.htm"
    Case 30
        BrowseMe "WRITE.htm"
    Case 31
        BrowseMe "ADDRESS  LATCH  ENABLE.htm"
    Case 32
        BrowseMe "S1 AND S0.htm"
    Case 33 To 40
        BrowseMe "ADDRESS BUS.htm"
End Select


End Sub


Private Sub Label_Reset()
Dim i As Integer
For i = 1 To Label1.Count - 1
    Label1(i).Appearance = 0
    Label1(i).BorderStyle = 0
Next
Browser.Width = 8175
Browser.Height = 3615

End Sub

Private Sub Labe2_Reset()
Dim i As Integer
On Error Resume Next
For i = 0 To Label2.Count - 1
    Label2(i).Appearance = 0
    Label2(i).BorderStyle = 0
Next
Browser.Width = 8175
Browser.Height = 3615

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label_Reset
Label1(Index).BorderStyle = 1
Label1(Index).Appearance = 1
DoEvents

End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not Label1(Index).BorderStyle = 1 Then
Label_Reset
Label1(Index).BorderStyle = 1
Label1(Index).Appearance = 1
DoEvents
End If
End Sub


Private Sub Label2_Click(Index As Integer)
Browser.Move Label2(Index).Left + Label2(Index).Width, Label2(Index).Top
Browser.Visible = True
Select Case Index
Case 0
    BrowseMe "ACCUMULATOR.htm"
Case 1
    BrowseMe "TEMPORARY REGISTER.htm"
Case 2
    BrowseMe "FLAGS.htm"
    Browser.Width = 15000
Case 3
    BrowseMe "ALU.htm"
Case 4
    BrowseMe "scan0003.jpg"
    Browser.Width = 15000
Case 5
    BrowseMe "INSTRUCTION REGISTER.htm"
Case 6
    BrowseMe "INSTRUCTION DECODER.htm"
Case 7
    BrowseMe "TIMING AND CONTROL.htm"
Case 8
    BrowseMe "W.htm"
    Browser.Height = 2000
Case 9
    BrowseMe "z.htm"
    Browser.Height = 2000
Case 10
    BrowseMe "b.htm"
Case 11
    BrowseMe "c.htm"
Case 12
    BrowseMe "d.htm"
Case 13
    BrowseMe "e.htm"
Case 14
    BrowseMe "h.htm"
Case 15
    BrowseMe "l.htm"
Case 16
    BrowseMe "stack pointer.htm"
Case 17
    BrowseMe "program counter.htm"
Case 19
    BrowseMe "CLK.htm"
Case 20
    BrowseMe "READY.htm"
Case 21
    BrowseMe "READ.htm"
Case 22
    BrowseMe "WRITE.htm"
Case 23
    BrowseMe "ADDRESS  LATCH  ENABLE.htm"
Case 24, 25
    BrowseMe "S1 AND S0.htm"
Case 26
    BrowseMe "IO.htm"
Case 27
    BrowseMe "HOLD.htm"
Case 28
    BrowseMe "HLDA.htm"
Case 29
    BrowseMe "RESETIN.htm"
Case 30
    BrowseMe "RESET OUT.htm"
Case 31
    BrowseMe "ADDRESS BUS.htm"
Case 32
    BrowseMe "DATA BUS.htm"
Case 33
    BrowseMe "INTR.htm"
Case 34
    BrowseMe "INTa.htm"
Case 35, 36, 37
    BrowseMe "RESTART INTERRUPTS.htm"
Case 38
    BrowseMe "TRAP.htm"
Case 39, 40
    BrowseMe "SERIAL INPUT DATA.htm"
Case 41
    BrowseMe "X1  AND X2.htm"
End Select
If Browser.Left + Browser.Width > ScaleWidth Then
    Browser.Left = ScaleWidth - Browser.Width
End If
If Browser.Top + Browser.Height > ScaleHeight - Command1.Height Then
    Browser.Top = ScaleHeight - (Browser.Height + Command1.Height)
End If
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Labe2_Reset
Label2(Index).BorderStyle = 1
Label2(Index).Appearance = 1
DoEvents

End Sub


Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Label2(Index).BorderStyle = 1 Then
Labe2_Reset

Label2(Index).BorderStyle = 1
Label2(Index).Appearance = 1
DoEvents
End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Browser.Visible = True Then
    Label_Reset
    Browser.Visible = False
    Browser.Navigate "about:blank"
End If
End Sub



Private Sub BrowseMe(FileName As String)
Browser.Navigate App.Path & "\Documentation\html\" & FileName
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Browser.Visible = True Then
    Labe2_Reset
    Browser.Visible = False
    Browser.Navigate "about:blank"
End If
End Sub



