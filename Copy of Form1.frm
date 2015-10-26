VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "project PMPS basic"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   6720
      Width           =   4695
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         Max             =   0
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   6840
      TabIndex        =   5
      Top             =   3480
      Width           =   375
      Begin VB.VScrollBar VScroll1 
         Height          =   11055
         Left            =   120
         Max             =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H0000FF00&
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2340
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2040
         Left            =   0
         Pattern         =   "*.jpg"
         TabIndex        =   1
         Top             =   2640
         Width           =   2535
      End
   End
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      Height          =   2535
      Left            =   2520
      ScaleHeight     =   2475
      ScaleWidth      =   4875
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   2400
         ScaleHeight     =   435
         ScaleWidth      =   675
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Shape1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu link 
         Caption         =   "Link"
         Begin VB.Menu Image 
            Caption         =   "Image"
         End
         Begin VB.Menu Text 
            Caption         =   "Text"
         End
         Begin VB.Menu open 
            Caption         =   "Open"
         End
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40


Dim placing As Boolean

Dim thisfile As File
Dim thisfolder As Folder
Dim path As String

Private Type link_block
    FileName As String * 50
    BlockIndex As Integer
    Type As String * 10
    L As Integer
    T As Integer
    W As Integer
    H As Integer
End Type

Private Type block
    TooltipText As String * 50
    link As link_block
    L As Integer
    T As Integer
    W As Integer
    H As Integer
End Type

Dim O() As block
Dim changed As Boolean
Dim currentIndex As Integer
Dim moved As Boolean
Private Sub delete_Click()
Shape1(currentIndex).Visible = False
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error GoTo se
Dim fsys As New FileSystemObject
Set thisfolder = fsys.GetFolder(Dir1.path)

thisfolder.Move thisfolder.ParentFolder & "\" & InputBox("Enter New Name", , thisfolder.Name)
Dir1.path = thisfolder.path
Exit Sub
se:
MsgBox Error

End If
End Sub


Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim i As Integer

If Me.MousePointer = 2 Then
Image1.Picture = LoadPicture(File1.path & "\" & File1.FileName)
O(currentIndex).link.FileName = File1.FileName
For i = 0 To Shape1.Count - 1
    If Shape1(i).Visible = True Then Shape1(i).Visible = False
Next

Exit Sub
End If

Dim fnum
fnum = FreeFile

If Image1.Visible = True And changed = True Then
    changed = False
    
    
    Kill Dir1.path & "/" & Image1.Tag & ".dat"
    
    Open Dir1.path & "/" & Image1.Tag & ".dat" For Random As #fnum Len = Len(O(0))
    For i = 1 To UBound(O)
        If Shape1(i).Visible = True Then Put #fnum, , O(i)
    Next
    Close #fnum

    'SavePicture Image1.Image, Image1.Tag
End If


While (Shape1.Count > 1)
    Unload Shape1(Shape1.Count - 1)
Wend

ReDim O(0)
Image1.Picture = LoadPicture(File1.path & "\" & File1.FileName)
Image1.Tag = File1.FileName
VScroll1.Value = 0

If Image1.Height > Height Then
    VScroll1.Max = Image1.ScaleHeight - ScaleHeight
    VScroll1.Min = 0
    VScroll1.LargeChange = VScroll1.Max / 2
    VScroll1.SmallChange = VScroll1.Max / 10
    
    VScroll1.Visible = True
Else
    VScroll1.Visible = False
End If

HScroll1.Value = 0
If Image1.Width > Width Then
    HScroll1.Max = Image1.ScaleWidth - ScaleWidth
    HScroll1.Min = 0
    HScroll1.LargeChange = HScroll1.Max / 2
    HScroll1.SmallChange = HScroll1.Max / 10
    HScroll1.Visible = True
Else
    HScroll1.Visible = False
End If
    Image1.Move 0, 0

    Open Dir1.path & "/" & File1.FileName & ".dat" For Random As #fnum Len = Len(O(0))
    
    If LOF(fnum) / Len(O(0)) = 0 Then
        'Add_tickets
        'Command1.Visible = False
        Close #fnum
    Else
    For i = 1 To LOF(fnum) / Len(O(0)) + 1
        ReDim Preserve O(i)
        Get #fnum, i, O(i)
        Load Shape1(i)
        Shape1(i).Move O(i).L, O(i).T, O(i).W, O(i).H
        Shape1(i).Visible = True
    Next
    Close #fnum
    End If

If Frame1.Left < 0 Then Image1.Left = Frame1.Width
Image1.Visible = True
End Sub

Private Sub File1_DblClick()
On Error GoTo se
Dim fsys As New FileSystemObject
Set thisfile = fsys.GetFile(File1.path & "\" & File1.FileName)

thisfile.Move File1.path & "\" & InputBox("Enter New Name") & ".jpg"
File1.Refresh
Exit Sub
se:
MsgBox Error
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
path = File1.path & "\" & File1.FileName
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If path = File1.path & "\" & File1.FileName Then Exit Sub

Dim fsys As New FileSystemObject

Set thisfile = fsys.GetFile(File1.path & "\" & File1.FileName)
thisfile.Move "temp"
Set thisfile = fsys.GetFile(path)
thisfile.Move File1.path & "\" & File1.FileName
Set thisfile = fsys.GetFile("temp")
thisfile.Move path



End Sub


Private Sub Form_Activate()
On Error Resume Next
Me.Visible = False
DoEvents

'Left = GetSetting(App.EXEName, "Config", "left", Left)
'Top = GetSetting(App.EXEName, "Config", "top", Top)
'Width = GetSetting(App.EXEName, "Config", "width", Width)
'Height = GetSetting(App.EXEName, "Config", "height", Height)
'WindowState = GetSetting(App.EXEName, "Config", "wstate", WindowState)

Drive1.Drive = GetSetting(App.EXEName, "Config", "drive", Drive1.Drive)
Dir1.path = GetSetting(App.EXEName, "Config", "dir", Dir1.path)
File1.path = GetSetting(App.EXEName, "Config", "file", File1.path)

Me.Visible = True
End Sub

Private Sub Form_DblClick()
Unload Me
End Sub


Private Sub Form_Load()
Me.WindowState = 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Visible = True Then Image1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_Resize()
On Error Resume Next
'If WindowState = 1 Then Exit Sub

'SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, _
           Me.Width / 15, Me.Height / 15, SWP_SHOWWINDOW
           
Frame1.Left = 0
Frame1.Top = 0
Frame1.Height = Me.ScaleHeight

Image1.Left = Frame1.Width

File1.Height = Frame1.Height - (Dir1.Top + Dir1.Height) + 200

VScroll1.Height = ScaleHeight
Frame2.Left = ScaleWidth - Frame2.Width
Frame2.Top = 0
Frame2.Height = ScaleHeight

HScroll1.Width = ScaleWidth
Frame3.Width = ScaleWidth
Frame3.Left = 0
Frame3.Top = ScaleHeight - Frame3.Height


File1.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
If WindowState = 1 Then Exit Sub

Me.Visible = False
DoEvents

Image1.Picture = LoadPicture("")
SaveSetting App.EXEName, "Config", "drive", Drive1.Drive
SaveSetting App.EXEName, "Config", "dir", Dir1.path
SaveSetting App.EXEName, "Config", "file", File1.path
SaveSetting App.EXEName, "Config", "wstate", WindowState
SaveSetting App.EXEName, "Config", "left", Left
SaveSetting App.EXEName, "Config", "top", Top
SaveSetting App.EXEName, "Config", "width", Width
SaveSetting App.EXEName, "Config", "height", Height
'End
End Sub



Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.Left = 0
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame2.Left = ScaleWidth - Frame2.Width
End Sub


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame3.Top = ScaleHeight - Frame3.Height
End Sub


Private Sub HScroll1_Change()
Image1.Left = -HScroll1.Value

End Sub


Private Sub HScroll1_Scroll()
Image1.Left = -HScroll1.Value

End Sub


Private Sub Image_Click()
O(currentIndex).link.Type = "image"
Me.MousePointer = 2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Load Shape1(Shape1.Count)
Shape1(Shape1.Count - 1).Move X, Y, 0, 0
Shape1(Shape1.Count - 1).ZOrder 0
Shape1(Shape1.Count - 1).Visible = True
placing = True
changed = True
moved = False
End Sub


Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If placing = True Then
    Shape1(Shape1.Count - 1).Width = X - Shape1(Shape1.Count - 1).Left
    Shape1(Shape1.Count - 1).Height = Y - Shape1(Shape1.Count - 1).Top
End If
If Not Image1.Left = Width / 2 - Image1.Width / 2 Or Not Frame1.Left = -Frame1.Width + 100 Then
    Frame1.Left = -Frame1.Width + 100

If VScroll1.Visible = False And HScroll1.Visible = False Then
    Frame2.Left = Width - 100
    VScroll1.TabStop = True
    VScroll1.SetFocus
    Image1.Left = Width / 2 - Image1.Width / 2

    Frame3.Top = ScaleTop + ScaleHeight - 100
End If
End If

DoEvents

Dim i As Integer

For i = 0 To Shape1.Count - 1
    If Shape1(i).Appearance = 1 Then Shape1(i).Appearance = 0
Next

If Picture1.Visible = True Then Picture1.Visible = False

moved = True
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer
placing = False
If moved = True Then

ReDim Preserve O(UBound(O) + 1)
O(UBound(O)).L = Shape1(Shape1.Count - 1).Left
O(UBound(O)).T = Shape1(Shape1.Count - 1).Top
O(UBound(O)).W = Shape1(Shape1.Count - 1).Width
O(UBound(O)).H = Shape1(Shape1.Count - 1).Height

End If
If Me.MousePointer = 2 Then
O(currentIndex).link.L = O(UBound(O)).L
O(currentIndex).link.T = O(UBound(O)).T
O(currentIndex).link.W = O(UBound(O)).W
O(currentIndex).link.H = O(UBound(O)).H
ReDim Preserve O(UBound(O) - 1)
Image1.Picture = LoadPicture(Dir1.path & "/" & Image1.Tag)
For i = 1 To Shape1.Count - 1
    If Shape1(i).Visible = False Then Shape1(i).Visible = True
Next

Unload Shape1(Shape1.Count - 1)
Me.MousePointer = 1
End If

End Sub


Private Sub Shape1_Click(Index As Integer)
If O(Index).link.Type = "image     " Then
    Picture1.Width = O(Index).link.W
    Picture1.Height = O(Index).link.H
    Picture1.Move Shape1(Index).Left, Shape1(Index).Top

    Picture1.PaintPicture LoadPicture(Dir1.path & "\" & O(Index).link.FileName), 0, 0 _
    , O(Index).link.W, O(Index).link.H _
    , O(Index).link.L, O(Index).link.T _
    , O(Index).link.W, O(Index).link.H
    
    Picture1.Visible = True
End If
End Sub

Private Sub Shape1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
currentIndex = Index
PopupMenu menu
changed = True
End If
End Sub

Private Sub Shape1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shape1(Index).Appearance = 0 Then Shape1(Index).Appearance = 1

End Sub



Private Sub VScroll1_Change()
Image1.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 34 Then
    If File1.ListCount - 1 > File1.ListIndex Then File1.ListIndex = File1.ListIndex + 1
    VScroll1.Value = 0
ElseIf KeyCode = 33 Then
    If 0 < File1.ListIndex Then File1.ListIndex = File1.ListIndex - 1
    VScroll1.Value = 0
ElseIf KeyCode = 27 Then
    Unload Me
End If
End Sub


Private Sub VScroll1_Scroll()
Image1.Top = -VScroll1.Value
End Sub


