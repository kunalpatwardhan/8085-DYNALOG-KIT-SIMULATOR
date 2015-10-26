VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Simulator"
   ClientHeight    =   11520
   ClientLeft      =   2025
   ClientTop       =   1245
   ClientWidth     =   15360
   LinkTopic       =   "Form2"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   -360
      TabIndex        =   57
      Top             =   0
      Width           =   11175
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   59
         Top             =   0
         Width           =   3720
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   3375
         Left            =   0
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   3
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   4650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   2295
      Begin VB.TextBox Hex_Flag 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox sptr 
         Height          =   375
         Left            =   2880
         TabIndex        =   56
         Text            =   "sptr"
         ToolTipText     =   "ptr"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   55
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   6
         Left            =   3120
         TabIndex        =   54
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   5
         Left            =   3600
         TabIndex        =   53
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   52
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   51
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   50
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   49
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox db 
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   48
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox bd 
         Height          =   375
         Left            =   2640
         TabIndex        =   47
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox ptr 
         Height          =   375
         Left            =   2880
         TabIndex        =   46
         Text            =   "ptr"
         ToolTipText     =   "ptr"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox M 
         Height          =   375
         Left            =   2640
         TabIndex        =   45
         Tag             =   "M"
         Text            =   "0"
         ToolTipText     =   "M"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   44
         Text            =   "0"
         ToolTipText     =   "CY"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   43
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   2
         Left            =   5040
         TabIndex        =   42
         Text            =   "0"
         ToolTipText     =   "P"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   41
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   40
         Text            =   "0"
         ToolTipText     =   "AC"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   5
         Left            =   3600
         TabIndex        =   39
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   6
         Left            =   3120
         TabIndex        =   38
         Text            =   "0"
         ToolTipText     =   "Z"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox F 
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   37
         Text            =   "0"
         ToolTipText     =   "S"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox tstr 
         Height          =   285
         Left            =   3600
         TabIndex        =   36
         Text            =   "tstr"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox W 
         Height          =   375
         Left            =   4200
         TabIndex        =   35
         Tag             =   "0"
         Text            =   "0"
         ToolTipText     =   "W"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Z 
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Tag             =   "0"
         Text            =   "0"
         ToolTipText     =   "Z"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox C 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "C"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox D 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "D"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox E 
         Alignment       =   2  'Center
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
         Left            =   1080
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "E"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox H 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "H"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox L 
         Alignment       =   2  'Center
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
         Left            =   1080
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "L"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox SP 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "SP"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox PC 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "PC"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox A 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "Accumulator"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Flag 
         Height          =   375
         Left            =   4800
         TabIndex        =   25
         Tag             =   "0"
         Text            =   "00000000"
         ToolTipText     =   "Flag"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox B 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "0"
         Text            =   "00"
         ToolTipText     =   "B"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   71
         Top             =   5040
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SP"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   70
         Top             =   4080
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L"
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   69
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   68
         Top             =   3120
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   67
         Top             =   2160
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   66
         Top             =   2160
         Width           =   120
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   5
         Left            =   120
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   4
         Left            =   120
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   3
         Left            =   120
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   2
         Left            =   120
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   65
         Top             =   1200
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   64
         Top             =   1200
         Width           =   105
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   1
         Left            =   120
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   63
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   61
         Top             =   240
         Width           =   105
      End
   End
   Begin VB.TextBox Disp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   12000
      TabIndex        =   22
      Text            =   "1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Disp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   9600
      TabIndex        =   21
      Text            =   "0"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   480
      TabIndex        =   72
      Top             =   120
      Width           =   2295
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   77
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   75
         Top             =   1800
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   74
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save Changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   73
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   840
      TabIndex        =   76
      Top             =   1560
      Width           =   200
   End
   Begin VB.Image Display1 
      Height          =   780
      Index           =   1
      Left            =   12990
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Display1 
      Height          =   780
      Index           =   0
      Left            =   12480
      Picture         =   "Form2.frx":13C2
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Display0 
      Height          =   780
      Index           =   3
      Left            =   11400
      Picture         =   "Form2.frx":2784
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Display0 
      Height          =   780
      Index           =   2
      Left            =   10890
      Picture         =   "Form2.frx":3B46
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Display0 
      Height          =   780
      Index           =   1
      Left            =   10350
      Picture         =   "Form2.frx":4F08
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Display0 
      Height          =   780
      Index           =   0
      Left            =   9840
      Picture         =   "Form2.frx":62CA
      Stretch         =   -1  'True
      Top             =   5400
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   9600
      Picture         =   "Form2.frx":768C
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F U4"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   20
      Left            =   13440
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E U3"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   19
      Left            =   12360
      TabIndex        =   19
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D U2"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   18
      Left            =   11280
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C U1"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   17
      Left            =   10200
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B SAVE"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   16
      Left            =   13320
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A LOAD"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   15
      Left            =   12360
      TabIndex        =   15
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9 L"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   14
      Left            =   11280
      TabIndex        =   14
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8 GO/ H"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   13
      Left            =   10200
      TabIndex        =   13
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7 PCL"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   12
      Left            =   13320
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6 PCH"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   11
      Left            =   12240
      TabIndex        =   11
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5 SPL"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   10
      Left            =   11160
      TabIndex        =   10
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4 SPH"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   9
      Left            =   10200
      TabIndex        =   9
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3 REG/ I"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   8
      Left            =   13320
      TabIndex        =   8
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 STEP"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   7
      Left            =   12240
      TabIndex        =   7
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1 CODE"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   6
      Left            =   11160
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 SET"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   5
      Left            =   10080
      TabIndex        =   5
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INR"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   4
      Left            =   8400
      TabIndex        =   4
      Top             =   10440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXEC"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   3
      Left            =   8400
      TabIndex        =   3
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DCR"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   2
      Left            =   8400
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VI"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   1
      Left            =   8400
      TabIndex        =   1
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Buttons 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RESET"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   0
      Left            =   7320
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SetT As Boolean
Dim Go As Boolean
Dim cnt As Byte

Dim INS As String
Dim Hcode As String
'Dim ptr As Long

Dim processing As Boolean
Dim stp As Boolean
Dim REG As Boolean
Dim REG_Index As Integer

Dim startpos As Long, endpos As Long

Dim OpenFile As String
Dim AR As Boolean

Dim Dpos As Integer

Dim Interval As Long
Dim startTime As Double
Dim Tstates As Integer
Dim Description As String
Private Sub Delay(ipInterval As Double)
    While (Timer - startTime < ipInterval)
    If Go = False Then Exit Sub
        DoEvents
    Wend
End Sub

Private Sub A_Change()
ChangeFlags A
A = trm(A)
A.Tag = "Changed"
End Sub

Private Sub B_Change()
ChangeFlags B
B = trm(B)
B.Tag = "Changed"
End Sub


Private Sub bd_Change()
Dim i As Integer
For i = 0 To 7
    bd.SelStart = i
    bd.SelLength = 1
    db(7 - i).Text = bd.SelText
Next
End Sub

Private Sub Buttons_Click(Index As Integer)
If Index = 0 Then
        Disp(0) = "FrIE"
        Disp(1) = "nD"
        ptr = 1
'        A = "A"
'        B = "B"
'        C = "C"
'        D = "D"
'        E = "E"
'        H = "H"
'        L = "L"
'        SP = "SP"
'        PC = "PC"
        SetT = False
        Go = False
        REG = False
        stp = False
        Exit Sub
End If
If REG = True Then
    If Index < 21 And Index > 7 Then REG_Index = Index
    Select Case Index
    Case 2
        If REG_Index > 8 Then
            Buttons_Click REG_Index - 1
        Else
            Buttons_Click 20
        End If
        Exit Sub
    Case 4
        If REG_Index < 20 Then
            Buttons_Click REG_Index + 1
        Else
            Buttons_Click 8
        End If
        Exit Sub
    Case 8
        Disp(0) = "I"
        Disp(1) = ""
    Case 9
        Disp(0) = "SPH"
        SP.SelStart = 0
        SP.SelLength = 2
        Disp(1) = SP.SelText
    Case 10
        Disp(0) = "SPL"
        SP.SelStart = 2
        SP.SelLength = 2
        Disp(1) = SP.SelText
    Case 11
        Disp(0) = "PCH"
        PC.SelStart = 0
        PC.SelLength = 2
        Disp(1) = PC.SelText
    Case 12
        Disp(0) = "PCL"
        PC.SelStart = 2
        PC.SelLength = 2
        Disp(1) = PC.SelText
    Case 13
        Disp(0) = "H"
        Disp(1) = H
    Case 14
        Disp(0) = "L"
        Disp(1) = L
    Case 15
        Disp(0) = "A"
        Disp(1) = A
    Case 16
        Disp(0) = "B"
        Disp(1) = B
    Case 17
        Disp(0) = "C"
        Disp(1) = C
    Case 18
        Disp(0) = "D"
        Disp(1) = D
    Case 19
        Disp(0) = "E"
        Disp(1) = E
    Case 20
        Disp(0) = "F"
        Disp(1) = Hex(BtoD(Flag))
    End Select
    Exit Sub
End If

ptr_Change
Dim i As Integer


If SetT Or Go Or stp Then
    Select Case Index
    Dim tstr As String
             Case 2:
          '   If Disp(1) = "" And Not Go And Len(Disp(0)) = 4 Then
          '      If CLng("&H" & Disp(0)) - startpos < Grid.Rows And CLng("&H" & Disp(0)) - startpos > 0 Then
          '          ptr = CLng("&H" & Disp(0)) - startpos
          '  '        GoTo s0
          '       End If
          '  End If
    
            If ptr > 1 Then
                ptr = ptr - 1 'RAM(0).ListIndex = RAM(0).ListIndex - 1
            Else
                ptr = endpos - (startpos)  'endpos - startpos
            End If
's0:
            Disp(0) = G(ptr, 1) ' 'RAM(0).Text
            Disp(1) = G(ptr, 2)   ' RAM(1).List(RAM(0).ListIndex)
            Disp(1).SelStart = 0
            Disp(1).SelLength = 2
            Disp(1).Enabled = True
'            Disp(1).SetFocus
            Exit Sub
        
        Case 3
            If Go Or stp Then
'                Go = False
 '              stp = False
                Execute
                Exit Sub
            End If
        Case 4:
        'If Disp(1) = "" And Not Go Then
        '     If CLng("&H" & Disp(0)) - startpos < Grid.Rows And CLng("&H" & Disp(0)) - startpos > 0 Then
        '        ptr = CLng("&H" & Disp(0)) - startpos
        '    '    GoTo s1
        '     End If
        'End If
            If ptr < endpos - startpos Then
                ptr = ptr + 1 'ptr=ptr+1
            Else
                ptr = 1
            End If
's1:
            Disp(0) = G(ptr, 1)    'RAM(0).Text
            Disp(1) = G(ptr, 2)  ' RAM(1).List(RAM(0).ListIndex)
            Disp(1).SelStart = 0
            Disp(1).SelLength = 2
            Disp(1).Enabled = True
 '           Disp(1).SetFocus
            Exit Sub
            
        Case 5:
            tstr = 0
        Case 6:
            tstr = 1
        Case 7:
            tstr = 2
        Case 8:
            tstr = 3
        Case 9:
            tstr = 4
        Case 10:
            tstr = 5
        Case 11:
            tstr = 6
        Case 12:
            tstr = 7
        Case 13:
            tstr = 8
        Case 14:
            tstr = 9
        Case 15
            tstr = "A"
        Case 16
            tstr = "B"
        Case 17
            tstr = "C"
        Case 18
            tstr = "D"
        Case 19
            tstr = "E"
        Case 20
            tstr = "F"
    End Select
    
    If Go And Len(Disp(0)) = 4 Then
        Disp(0) = ""
        Disp(1) = ""
    End If
    
    If Len(Disp(0)) < 4 Then
        Disp(0) = Disp(0) & tstr
    Else
        Disp(1).SelText = ""
        Disp(1) = Disp(1) & tstr
    End If
    
Else
    Select Case Index
    Case 5
        Disp(0) = ""
        Disp(1) = ""
        Disp(0).Enabled = True
        SetT = True
    Case 7
        Disp(0) = "" 'G(ptr, 1)
        Disp(1) = "" 'G(ptr, 2)
        Go = True
        stp = True
    Case 8
        REG = True
        Disp(0) = ""
        Disp(1) = ""
    Case 13
        Disp(0) = "" 'G(ptr, 1)
        Disp(1) = "" 'G(ptr, 2)
        Go = True
    End Select
    
End If

End Sub




Private Sub Buttons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Buttons(Index).Appearance = 1
     Buttons(Index).BorderStyle = 1
End Sub

Private Sub Buttons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Buttons(Index).Appearance = 0
    Buttons(Index).BorderStyle = 0
End Sub


Private Sub C_Change()
ChangeFlags C
C = trm(C)
C.Tag = "Changed"
End Sub

Private Sub D_Change()
ChangeFlags D
D = trm(D)
D.Tag = "Changed"
End Sub


Private Sub Disp_Change(Index As Integer)

If Disp(0) = "E   " Or Disp(0) = "FrIE" Or Disp(0) = "E   " Then
    Disp(0).Enabled = False
    Disp(1).Enabled = False
    GoTo es
End If

If processing = True Then Exit Sub

On Error Resume Next
Dim i As Integer
Select Case Index
Case 0:
    If Len(Disp(0)) = 4 Then
        If CLng("&H" & Disp(0)) - startpos < Grid.Rows And CLng("&H" & Disp(0)) - startpos > 0 Then
            ptr = CLng("&H" & Disp(0)) - startpos
        Else
            MsgBox "The Memory location " & Disp(0) & " does not exist." & vbNewLine _
                & "Reseting ..."
                Buttons_Click 0
        End If
    End If
Case 1:
    If Len(Disp(1)) = 2 And Len(Disp(0)) = 4 And CLng("&H" & Disp(0)) - startpos < Grid.Rows And CLng("&H" & Disp(0)) - startpos > 0 Then
         G ptr, 2, Disp(1)
    End If
End Select
es:

' Now Glow LED's
Dim iPath As String
iPath = App.Path & "\images\font\"

If Index = 0 Then
    For i = 0 To 3
        Disp(0).SelStart = i
        Disp(0).SelLength = 1
        If Disp(0).SelText = "" Or Disp(0).SelText = " " Then
            Display0(i).Picture = LoadPicture(iPath & "Null.jpg")
        Else
            Display0(i).Picture = LoadPicture(iPath & Disp(0).SelText & ".jpg")
        End If
    Next
Else
    For i = 0 To 1
        Disp(1).SelStart = i
        Disp(1).SelLength = 1
        If Disp(1).SelText = "" Or Disp(1).SelText = " " Then
            Display1(i).Picture = LoadPicture(iPath & "Null.jpg")
        Else
            Display1(i).Picture = LoadPicture(iPath & Disp(1).SelText & ".jpg")
        End If
    Next

End If
End Sub






Public Sub Execute()
'processing = True
Disp(0) = "E   "
Disp(1) = ""
'On Error GoTo eHandler
While (1)
startTime = Timer
If Hcode = "" Then
    MsgBox "No Instruction to Execute." & vbNewLine _
        & "Aborting ..."
    Disp(0) = "Err "
    Disp(1) = ""
        
    Exit Sub
End If

Add_Description (Hcode)

' < Decode and Execute Instruction '''''''''''''''''''''''''''''''''''''''''''''''''

    Select Case Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ACI 8-bit data
' Add Immediate to Accumulator with Carry
    Case "CE":
    Tstates = 7
        ptr = ptr + 1
        A = Hex(CLng("&H" & A) + CLng("&H" & G(ptr, 2)) + CLng("&H" & F(0)))
        G ptr - 1, 3, "ACI " & G(ptr, 2)
        F(0) = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ADC Reg./Mem
' Add Register to Accumulator with Carry
    Case "88":
        A = Hex(CLng("&H" & A) + CLng("&H" & B) + CLng("&H" & F(0)))
        G ptr, 3, "ADC B"
        F(0) = 0
    Case "89":
        A = Hex(CLng("&H" & A) + CLng("&H" & C) + CLng("&H" & F(0)))
        G ptr, 3, "ADC C"
        F(0) = 0
    Case "8A":
        A = Hex(CLng("&H" & A) + CLng("&H" & D) + CLng("&H" & F(0)))
        G ptr, 3, "ADC D"
        F(0) = 0
    Case "8B":
        A = Hex(CLng("&H" & A) + CLng("&H" & E) + CLng("&H" & F(0)))
        G ptr, 3, "ADC E"
        F(0) = 0
    Case "8C":
        A = Hex(CLng("&H" & A) + CLng("&H" & H) + CLng("&H" & F(0)))
        G ptr, 3, "ADC H"
        F(0) = 0
    Case "8D":
        A = Hex(CLng("&H" & A) + CLng("&H" & L) + CLng("&H" & F(0)))
        G ptr, 3, "ADC L"
        F(0) = 0
    Case "8E":
        A = Hex(CLng("&H" & A) + CLng("&H" & M) + CLng("&H" & F(0)))
        G ptr, 3, "ADC M"
        F(0) = 0
    Case "8F":
        A = Hex(CLng("&H" & A) + CLng("&H" & A) + CLng("&H" & F(0)))
        G ptr, 3, "ADC A"
        F(0) = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ADD Reg/Mem
' Add Register to Accumulator
    Case "80": ' ADD B
        A = Hex(CLng("&H" & A) + CLng("&H" & B))
        G ptr, 3, "ADD B" 'G ptr, 3,  "ADD B"
    Case "81":
        A = Hex(CLng("&H" & A) + CLng("&H" & C))
        G ptr, 3, "ADD C"
    Case "82":
        A = Hex(CLng("&H" & A) + CLng("&H" & D))
        G ptr, 3, "ADD D"
    Case "83":
        A = Hex(CLng("&H" & A) + CLng("&H" & E))
        G ptr, 3, "ADD E"
    Case "84":
        A = Hex(CLng("&H" & A) + CLng("&H" & H))
        G ptr, 3, "ADD H"
    Case "85":
        A = Hex(CLng("&H" & A) + CLng("&H" & L))
        G ptr, 3, "ADD L"
    Case "86":
        A = Hex(CLng("&H" & A) + CLng("&H" & M))
        G ptr, 3, "ADD M"
    Case "87":
        A = Hex(CLng("&H" & A) + CLng("&H" & A))
        G ptr, 3, "ADD A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ADI 8-bit data
' ADD Immediate to Accumulator
    Case "C6":
        ptr = ptr + 1
        A = Hex(CLng("&H" & A) + CLng("&H" & Hcode))
        G ptr - 1, 3, "ADI " & G(ptr, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ANA
' Logical AND with Accumulator
    Case "A0"
        A = Hex(CLng("&H" & A) And CLng("&H" & B))
        G ptr, 3, "ANA B"
    Case "A1"
        A = Hex(CLng("&H" & A) And CLng("&H" & C))
        G ptr, 3, "ANA C"
    Case "A2"
        A = Hex(CLng("&H" & A) And CLng("&H" & D))
        G ptr, 3, "ANA D"
    Case "A3"
        A = Hex(CLng("&H" & A) And CLng("&H" & E))
        G ptr, 3, "ANA E"
    Case "A4"
        A = Hex(CLng("&H" & A) And CLng("&H" & H))
        G ptr, 3, "ANA H"
    Case "A5"
        A = Hex(CLng("&H" & A) And CLng("&H" & L))
        G ptr, 3, "ANA L"
    Case "A6"
        A = Hex(CLng("&H" & A) And CLng("&H" & M))
        G ptr, 3, "ANA M"
    Case "A7"
        A = Hex(CLng("&H" & A) And CLng("&H" & A))
        G ptr, 3, "ANA A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ANI 8-bit data
' AND Immediate with Accumulator
    Case "E6"
        ptr = ptr + 1
        A = Hex(CLng("&H" & A) And CLng("&H" & Hcode))
        G ptr - 1, 3, "ANI " & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CALL 16-bit Address
' Unconditional Subroutine Call
    Case "CD"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        If G(ptr, 2) & G(ptr - 1, 2) = "036E" Then
            Disp(1) = A
'            MsgBox "The subroutine of memory location 036E is called." & vbNewLine _
                & "A = " & A
        G ptr - 2, 3, "CALL " & G(ptr, 2) & G(ptr - 1, 2)  ' < Adding Mnemonics
        GoTo s1
        Else
        G ptr - 2, 3, "CALL " & G(ptr, 2) & G(ptr - 1, 2)  ' < Adding Mnemonics
        CALL_UNconditional
        End If
s1:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Conditional Call to Subroutine
    Case "DC"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CC " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(0) = 1 Then CALL_UNconditional
    Case "D4"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CNC " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(0) = 0 Then CALL_UNconditional
    Case "F4"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CP " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(7) = 0 Then CALL_UNconditional
    Case "FC"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CM " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(7) = 1 Then CALL_UNconditional
    Case "EC"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CPE " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(2) = 1 Then CALL_UNconditional
    Case "E4"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CPO " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(2) = 0 Then CALL_UNconditional
    Case "CC"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CZ " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(6) = 1 Then CALL_UNconditional
    Case "C4"
        ptr = ptr + 1           ' < 3 Byte Instruction
        ptr = ptr + 1
        G ptr - 2, 3, "CNZ " & G(ptr + 2, 2) & G(ptr + 1, 2) ' < Adding Mnemonics
        If F(6) = 0 Then CALL_UNconditional

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CMA
' Complement Accumulator
    Case "2F"
        A = Complement(A)
        G ptr, 3, "CMA"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CMC
' Complement Carry
    Case "3F"
        If F(0) = 0 Then
            F(0) = 1
        Else
            F(0) = 0
        End If
        G ptr, 3, "CMC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CMP Reg./Mem
' Compare with Accumulator
    Case "B8"
        CMP B
        G ptr, 3, "CMP B"
    Case "B9"
        CMP C
        G ptr, 3, "CMP C"
    Case "BA"
        CMP D
        G ptr, 3, "CMP D"
    Case "BB"
        CMP E
        G ptr, 3, "CMP E"
    Case "BC"
        CMP H
        G ptr, 3, "CMP H"
    Case "BD"
        CMP L
        G ptr, 3, "CMP L"
    Case "BE"
        CMP M
        G ptr, 3, "CMP M"
        Tstates = 7
    Case "BF"
        CMP A
        G ptr, 3, "CMP A"
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPI 8-bit
' Compare Immediate with Accumulator
    Case "FE"
        ptr = ptr + 1
        CMP G(ptr, 2)
        G ptr - 1, 3, "CPI " & G(ptr, 2)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DAA
' Decimal-Adjust Accumulator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DAD Reg. pair
' Add Register Pair to H and L Registers
    Case "09":
        DAD B, C
        G ptr, 3, "DAD B"
    Case "19":
        DAD D, E
        G ptr, 3, "DAD D"
    Case "29":
        DAD H, L
        G ptr, 3, "DAD H"
    Case "39":
        DAD 0, SP
        G ptr, 3, "DAD SP"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DCR R/M
' Decrement source by 1
    Case "05":
        B = Hex(CLng("&H" & B) - 1)
        G ptr, 3, "DCR B"
    Case "0D":
        C = Hex(CLng("&H" & C) - 1)
        G ptr, 3, "DCR C"
    Case "15":
        D = Hex(CLng("&H" & D) - 1)
        G ptr, 3, "DCR D"
    Case "1D":
        E = Hex(CLng("&H" & E) - 1)
        G ptr, 3, "DCR E"
    Case "25":
        H = Hex(CLng("&H" & H) - 1)
        G ptr, 3, "DCR H"
    Case "2D":
        L = Hex(CLng("&H" & L) - 1)
        G ptr, 3, "DCR L"
    Case "35":
        M = Hex(CLng("&H" & M) - 1)
        Tstates = 10
        G ptr, 3, "DCR M"
    Case "3D":
        A = Hex(CLng("&H" & A) - 1)
        G ptr, 3, "DCR A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DCX Rp
' Decrement Register Pair by 1
    Case "0B"
         tstr = Format(Hex(CLng("&H" & B & C) - 1), "0###")
         If Len(tstr) = 3 Then tstr = "0" & tstr
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            C = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            B = tstr.SelText
        G ptr, 3, "DCX B"
    Case "1B"
         tstr = Format(Hex(CLng("&H" & D & E) - 1), "0###")
         If Len(tstr) = 3 Then tstr = "0" & tstr
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            E = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            D = tstr.SelText
        G ptr, 3, "DCX D"
    Case "2B"
         tstr = Format(Hex(CLng("&H" & H & L) - 1), "0###")
         If Len(tstr) = 3 Then tstr = "0" & tstr
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            L = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            H = tstr.SelText
        G ptr, 3, "DCX H"
    Case "3B"
         SP = Format(Hex(CLng("&H" & SP) - 1), "0###")
         G ptr, 3, "DCX SP"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INR R/M
' Increment Contents of Register/Memory by 1
    Case "04":
        B = Hex(CLng("&H" & B) + 1)
        G ptr, 3, "INR B"
    Case "0C":
        C = Hex(CLng("&H" & C) + 1)
        G ptr, 3, "INR C"
    Case "14":
        D = Hex(CLng("&H" & D) + 1)
        G ptr, 3, "INR D"
    Case "1C":
        E = Hex(CLng("&H" & E) + 1)
        G ptr, 3, "INR E"
    Case "24":
        H = Hex(CLng("&H" & H) + 1)
        G ptr, 3, "INR H"
    Case "2C":
        L = Hex(CLng("&H" & L) + 1)
        G ptr, 3, "INR L"
    Case "34":
        Tstates = 10
        M = Hex(CLng("&H" & M) + 1)
        G ptr, 3, "INR M"
    Case "3C":
        A = Hex(CLng("&H" & A) + 1)
        G ptr, 3, "INR A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' INX Rp
' Increment Register Pair by 1
    Case "03"
    Tstates = 6
         tstr = Format(Hex(CLng("&H" & B & C) + 1), "0###")
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            C = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            B = tstr.SelText
        G ptr, 3, "INX B"
    Case "13"
    Tstates = 6
         tstr = Format(Hex(CLng("&H" & D & E) + 1), "0###")
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            E = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            D = tstr.SelText
        G ptr, 3, "INX D"
    Case "23"
    Tstates = 6
         tstr = Format(Hex(CLng("&H" & H & L) + 1), "0###")
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            L = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            H = tstr.SelText
        G ptr, 3, "INX H"
    Case "33"
    Tstates = 6
         SP = Format(Hex(CLng("&H" & SP) + 1), "0###")
         G ptr, 3, "INX SP"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JMP 16-bit
' Jump Unconditionaly
    Case "C3"
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        PC = G(ptr, 2) & G(ptr - 1, 2)
        G tstr - 2, 3, "JMP " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JC 16-bit
' Jump on Carry
    Case "DA"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(0) = 1 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JC " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JNC 16-bit
' Jump on No Carry
    Case "D2"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(0) = 0 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JNC " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JP 16-bit
' Jump on positive
    Case "F2"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(7) = 0 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JP " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JM 16-bit
' Jump on minus
    Case "FA"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(7) = 1 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JM " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JPE 16-bit
' Jump on Parity Even
    Case "EA"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(2) = 1 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JPE " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JPO 16-bit
' Jump on Parity Odd
    Case "E2"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(2) = 0 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JPO " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JZ 16-bit
' Jump on Zero
    Case "CA"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = ptr
        If F(6) = 1 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G tstr - 2, 3, "JZ " & G(tstr, 2) & G(tstr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' JNZ 16-bit
' Jump on No Zero
    Case "C2"
    Tstates = 7
        ptr = ptr + 1
        ptr = ptr + 1
        tstr = G(ptr, 2) & G(ptr - 1, 2)
        If F(6) = 0 Then
            PC = G(ptr, 2) & G(ptr - 1, 2)
            Tstates = 10
        End If
        G ptr - 2, 3, "JNZ " & tstr
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LDA 16-bit Address
' Load Accumulator Direct
    Case "3A"
        ptr = ptr + 1
        ptr = ptr + 1
        A = getme(G(ptr, 2) & G(ptr - 1, 2))
        G ptr - 2, 3, "LDA " & G(ptr, 2) & G(ptr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LDAX B/D reg.pair
' Load Accumulator Indirect
    Case "0A"
        A = G(CLng("&H" & B & C) - startpos, 2)
        G ptr, 3, "LDAX B"
    Case "1A"
        A = G(CLng("&H" & D & E) - startpos, 2)
        G ptr, 3, "LDAX D"
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LHLD 16 bit address
' Load H and L Registers Direct
    Case "2A":
        ptr = ptr + 1
        ptr = ptr + 1
        L = getme(G(ptr, 2) & G(ptr - 1, 2))
        H = getme(Hex(CLng("&H" & G(ptr, 2) & G(ptr - 1, 2)) + 1))
        G ptr - 2, 3, "LHLD " & G(ptr, 2) & G(ptr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LXI Reg Pair/16 bit data
' Load Register Pair Immediate
    Case "01": ' LXI B
        ptr = ptr + 1
        C = Hcode
        ptr = ptr + 1
        B = Hcode
        G ptr - 2, 3, "LXI B," & G(ptr, 2) & G(ptr - 1, 2)
    Case "11":
        ptr = ptr + 1
        E = Hcode
        ptr = ptr + 1
        D = Hcode
        G ptr - 2, 3, "LXI D," & G(ptr, 2) & G(ptr - 1, 2)
    Case "21":
        ptr = ptr + 1
        L = Hcode
        ptr = ptr + 1
        H = Hcode
        G ptr - 2, 3, "LXI H," & G(ptr, 2) & G(ptr - 1, 2)
    Case "31": ' SP
        ptr = ptr + 1
        ptr = ptr + 1
        SP = Hcode & G(ptr - 1, 2)
        G ptr - 2, 3, "LXI SP," & G(ptr, 2) & G(ptr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MOV Dest,Source
' Copy from Source to Destination
    Case "40": ' MOV B,B
        B = B
        G ptr, 3, "MOV B,B"
    Case "41": ' MOV B,C
        B = C
        G ptr, 3, "MOV B,C"
    Case "42": ' MOV B,D
        B = D
        G ptr, 3, "MOV B,D"
    Case "43": ' MOV B,E
        B = E
        G ptr, 3, "MOV B,E"
    Case "44": ' MOV B,H
        B = H
        G ptr, 3, "MOV B,H"
    Case "45": ' MOV B,L
        B = L
        G ptr, 3, "MOV B,L"
    Case "46": ' MOV B,M
        B = M
        G ptr, 3, "MOV B,M"
    Case "47": ' MOV B,A
        B = A
        G ptr, 3, "MOV B,A"
    Case "48": ' MOV C,B
        C = B
        G ptr, 3, "MOV C,B"
    Case "49": ' MOV C,C
        C = C
        G ptr, 3, "MOV C,C"
    Case "4A": ' MOV C,D
        C = D
        G ptr, 3, "MOV C,D"
    Case "4B": ' MOV C,E
        C = E
        G ptr, 3, "MOV C,E"
    Case "4C": ' MOV C,H
        C = H
        G ptr, 3, "MOV C,H"
    Case "4D": ' MOV C,L
        C = L
        G ptr, 3, "MOV C,L"
    Case "4E": ' MOV C,M
        C = M
        G ptr, 3, "MOV C,M"
    Case "4F": ' MOV C,A
        C = A
        G ptr, 3, "MOV C,A"
    Case "50": ' MOV D,B
        D = B
        G ptr, 3, "MOV D,B"
    Case "51": ' MOV D,C
        D = C
        G ptr, 3, "MOV D,C"
    Case "52": ' MOV D,D
        D = D
        G ptr, 3, "MOV D,D"
    Case "53": ' MOV D,E
        D = E
        G ptr, 3, "MOV D,E"
    Case "54": ' MOV D,H
        D = H
        G ptr, 3, "MOV D,H"
    Case "55": ' MOV D,L
        D = L
        G ptr, 3, "MOV D,L"
    Case "56": ' MOV D,M
        D = M
        G ptr, 3, "MOV D,M"
    Case "57": ' MOV D,A
        D = A
        G ptr, 3, "MOV D,A"
    Case "58": ' MOV E,B
        E = B
        G ptr, 3, "MOV E,B"
    Case "59": ' MOV E,C
        E = C
        G ptr, 3, "MOV E,C"
    Case "5A": ' MOV E,D
        E = D
        G ptr, 3, "MOV E,D"
    Case "5B": ' MOV E,E
        E = E
        G ptr, 3, "MOV E,E"
    Case "5C": ' MOV E,H
        E = H
        G ptr, 3, "MOV E,H"
    Case "5D": ' MOV E,L
        E = L
        G ptr, 3, "MOV E,L"
    Case "5E": ' MOV E,M
        E = M
        G ptr, 3, "MOV E,M"
    Case "5F": ' MOV E,A
        E = A
        G ptr, 3, "MOV E,A"
        
    Case "60": ' MOV H,B
        H = B
        G ptr, 3, "MOV H,B"
    Case "61": ' MOV H,C
        H = C
        G ptr, 3, "MOV H,C"
    Case "62": ' MOV H,D
        H = D
        G ptr, 3, "MOV H,D"
    Case "63": ' MOV H,E
        H = E
        G ptr, 3, "MOV H,E"
    Case "64": ' MOV H,H
        H = H
        G ptr, 3, "MOV H,H"
    Case "65": ' MOV H,L
        H = L
        G ptr, 3, "MOV H,L"
    Case "66": ' MOV H,M
        H = M
        G ptr, 3, "MOV H,M"
    Case "67": ' MOV H,A
        H = A
        G ptr, 3, "MOV H,A"
        
    Case "68": ' MOV L,B
        L = B
        G ptr, 3, "MOV L,B"
    Case "69": ' MOV L,C
        L = C
        G ptr, 3, "MOV L,C"
    Case "6A": ' MOV L,D
        L = D
        G ptr, 3, "MOV L,D"
    Case "6B": ' MOV L,E
        L = E
        G ptr, 3, "MOV L,E"
    Case "6C": ' MOV L,H
        L = H
        G ptr, 3, "MOV L,H"
    Case "6D": ' MOV L,L
        L = L
        G ptr, 3, "MOV L,L"
    Case "6E": ' MOV L,M
        L = M
        G ptr, 3, "MOV L,M"
    Case "6F": ' MOV L,A
        L = A
        G ptr, 3, "MOV L,A"
        
    Case "70": ' MOV M,B
        M = B
        G ptr, 3, "MOV M,B"
    Case "71": ' MOV M,C
        M = C
        G ptr, 3, "MOV M,C"
    Case "72": ' MOV M,D
        M = D
        G ptr, 3, "MOV M,D"
    Case "73": ' MOV M,E
        M = E
        G ptr, 3, "MOV M,E"
    Case "74": ' MOV M,H
        M = H
        G ptr, 3, "MOV M,H"
    Case "75": ' MOV M,L
        M = L
        G ptr, 3, "MOV M,L"
    Case "77": ' MOV M,A
        M = A
        G ptr, 3, "MOV M,A"
        
    Case "78": ' MOV A,B
        A = B
        G ptr, 3, "MOV A,B"
    Case "79": ' MOV A,C
        A = C
        G ptr, 3, "MOV A,C"
    Case "7A": ' MOV A,D
        A = D
        G ptr, 3, "MOV A,D"
    Case "7B": ' MOV A,E
        A = E
        G ptr, 3, "MOV A,E"
    Case "7C": ' MOV A,H
        A = H
        G ptr, 3, "MOV A,H"
    Case "7D": ' MOV A,L
        A = L
        G ptr, 3, "MOV A,L"
    Case "7E": ' MOV A,M
        A = M
        G ptr, 3, "MOV A,M"
    Case "7F": ' MOV A,A
        A = A
        G ptr, 3, "MOV A,A"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MVI R/Data/Mem
' Move Immediate

    Case "06":
        ptr = ptr + 1
        B = Hcode
        G ptr - 1, 3, "MVI B," & Hcode
    Case "0E":
        ptr = ptr + 1
        C = Hcode
        G ptr - 1, 3, "MVI C," & Hcode
    Case "16":
        ptr = ptr + 1
        D = Hcode
        G ptr - 1, 3, "MVI D," & Hcode
    Case "1E":
        ptr = ptr + 1
        E = Hcode
        G ptr - 1, 3, "MVI E," & Hcode
    Case "26":
        ptr = ptr + 1
        H = Hcode
        G ptr - 1, 3, "MVI H," & Hcode
    Case "2E":
        ptr = ptr + 1
        L = Hcode
        G ptr - 1, 3, "MVI L," & Hcode
    Case "36":
        Tstates = 10
        ptr = ptr + 1
        M = Hcode
        G ptr - 1, 3, "MVI M," & Hcode
    Case "3E":
        ptr = ptr + 1
        A = Hcode
        G ptr - 1, 3, "MVI A," & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NOP
' No Operation
    Case "00"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ORA
' Logically OR with Accumulator
    Case "B0"
        ORA B
        G ptr, 3, "ORA B"
    Case "B1"
        ORA C
        G ptr, 3, "ORA C"
    Case "B2"
        ORA D
        G ptr, 3, "ORA D"
    Case "B3"
        ORA E
        G ptr, 3, "ORA E"
    Case "B4"
        ORA H
        G ptr, 3, "ORA H"
    Case "B5"
        ORA L
        G ptr, 3, "ORA L"
    Case "B6"
        Tstates = 7
        ORA M
        G ptr, 3, "ORA M"
    Case "B7"
        ORA A
        G ptr, 3, "ORA A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ORI 8-bit data
' Logically OR Immediate
    Case "F6"
        ptr = ptr + 1
        ORA Hcode
        G ptr, 3, "ORI " & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PCHL
' Load Program Counter with HL Contents
    Case "E9"
        PC = H & L
        G ptr, 3, "PCHL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' POP Rp
' Pop off Stack to Register Pair
    Case "C1"
        C = G(sptr, 2)
        sptr = sptr + 1
        B = G(sptr, 2)
        G ptr, 3, "POP B"
        sptr = sptr + 1
    Case "D1"
        E = G(sptr, 2)
        sptr = sptr + 1
        D = G(sptr, 2)
        G ptr, 3, "POP D"
        sptr = sptr + 1
    Case "E1"
        L = G(sptr, 2)
        sptr = sptr + 1
        H = G(sptr, 2)
        G ptr, 3, "POP H"
        sptr = sptr + 1
    Case "F1"
        Flag = G(sptr, 2)
        sptr = sptr + 1
        A = G(sptr, 2)
        G ptr, 3, "POP PSW"
        sptr = sptr + 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PUSH
' Push Register Pair onto Stack
    Case "C5"
        PUSH B, C
        G ptr, 3, "PUSH B"
     Case "D5"
        PUSH D, E
        G ptr, 3, "PUSH D"
     Case "E5"
        PUSH H, L
        G ptr, 3, "PUSH H"
     Case "F5"
        PUSH A, Hex_Flag
        G ptr, 3, "PUSH PSW"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RAL
' Rotate Accumulator Left through Carry
    Case "17"
        bd = Cbin(CLng("&H" & A))
        A = Hex(BtoD(db(6) & db(5) & db(4) & db(3) & db(2) & db(1) & db(0) & F(0)))
        F(0) = db(7)
        G ptr, 3, "RAL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RAR
' Rotate Accumulator Right through Carry
    Case "1F"
        bd = Cbin(CLng("&H" & A))
        A = Hex(BtoD(F(0) & db(7) & db(6) & db(5) & db(4) & db(3) & db(2) & db(1)))
        F(0) = db(0)
        G ptr, 3, "RAR"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RLC
' Rotate Accumulator Left
    Case "07"
        bd = Cbin(CLng("&H" & A))
        A = Hex(BtoD(db(6) & db(5) & db(4) & db(3) & db(2) & db(1) & db(0) & db(7)))
        F(0) = db(7)
        G ptr, 3, "RLC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RRC
' Rotate Accumulator Right
    Case "0F"
        bd = Cbin(CLng("&H" & A))
        A = Hex(BtoD(db(0) & db(7) & db(6) & db(5) & db(4) & db(3) & db(2) & db(1)))
        F(0) = db(0)
        G ptr, 3, "RRC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RET
' Return from Subroutine Unconditionaly
    Case "C9"
        tstr = ptr
        RET
        G tstr, 3, "RET"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RC
' Return on Carry
    Case "D8"
    Tstates = 6
        tstr = ptr
        If F(0) = 1 Then
            RET
        End If
        G tstr, 3, "RC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RNC
' Return on no Carry
    Case "D0"
    Tstates = 6
        tstr = ptr
        If F(0) = 0 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RNC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RP
' Return on positive
    Case "F0"
    Tstates = 6
        tstr = ptr
        If F(7) = 0 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RP"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RM
' Return on minus
    Case "F8"
    Tstates = 6
        tstr = ptr
        If F(7) = 1 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RM"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RPE
' Return on Parity Even
    Case "E8"
    Tstates = 6
        tstr = ptr
        If F(2) = 1 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RPE"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RPO
' Return on Parity Odd
    Case "E0"
    Tstates = 6
        tstr = ptr
        If F(2) = 0 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RPO"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RZ
' Return on Zero
    Case "C8"
    Tstates = 6
        tstr = ptr
        If F(6) = 1 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RZ"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' RNZ
' Return on No Zero
    Case "D8"
    Tstates = 6
        tstr = ptr
        If F(6) = 0 Then
            RET
            Tstates = 12
        End If
        G tstr, 3, "RNZ"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SBB Reg./Mem.
' Subtract Source and Borrow from Accumulator
    Case "98"
        Subtract_SBB B
        G ptr, 3, "SBB B"
    Case "99"
        Subtract_SBB C
        G ptr, 3, "SBB C"
    Case "9A"
        Subtract_SBB D
        G ptr, 3, "SBB D"
    Case "9B"
        Subtract_SBB E
        G ptr, 3, "SBB E"
    Case "9C"
        Subtract_SBB H
        G ptr, 3, "SBB H"
    Case "9D"
        Subtract_SBB L
        G ptr, 3, "SBB L"
    Case "9E"
        Tstates = 7
        Subtract_SBB M
        G ptr, 3, "SBB M"
    Case "9F"
        Subtract_SBB A
        G ptr, 3, "SBB A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SBI 8-bit data
' Subtract Immediate with Borrow
    Case "DE"
        ptr = ptr + 1
        Subtract_SBB Hcode
        G ptr, 3, "SBI " & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SHLD 16-bit Address
' Stores H and L Registers Direct
    Case "22":
        ptr = ptr + 1
        ptr = ptr + 1
        setme G(ptr, 2) & G(ptr - 1, 2), L
        setme Hex((G(ptr, 2) & G(ptr - 1, 2)) + 1), H
        G ptr, 3, "SHLD " & G(ptr, 2) & G(ptr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SPHL
' Copy H and L Registers to Stack Pointer
    Case "F9"
        SP = H & L
        G ptr, 3, "SPHL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' STA 16-bit Address
' Store Accumulator Direct
    Case "32"
        ptr = ptr + 1
        ptr = ptr + 1
        setme G(ptr, 2) & G(ptr - 1, 2), A
        G ptr, 3, "STA " & G(ptr, 2) & G(ptr - 1, 2)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' STAX B/D Rp
' Store Accumulator Indirect
    Case "02"
        setme B & C, A
        G ptr, 3, "STAX B"
    Case "12"
        setme D & E, A
        G ptr, 3, "STAX D"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' STC
' Set Carry
    Case "37"
        F(0) = 1
        G ptr, 3, "STC"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SUB Reg./Mem.
' Subtract Register or Memory from Accumulator
    Case "90"
        Subtract_SUB (B)
        G ptr, 3, "SUB B"
    Case "91"
        Subtract_SUB (C)
        G ptr, 3, "SUB C"
    Case "92"
        Subtract_SUB (D)
        G ptr, 3, "SUB D"
    Case "93"
        Subtract_SUB (E)
        G ptr, 3, "SUB E"
    Case "94"
        Subtract_SUB (H)
        G ptr, 3, "SUB H"
    Case "95"
        Subtract_SUB (L)
        G ptr, 3, "SUB L"
    Case "96"
        Subtract_SUB (M)
        G ptr, 3, "SUB M"
        Tstates = 7
    Case "97"
        Subtract_SUB (A)
        G ptr, 3, "SUB A"
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SUI 8-bit data
' Subtract Immediate with Accumulator
    Case "D6"
        ptr = ptr + 1
        Subtract_SUB (Hcode)
        G ptr, 3, "SUI " & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' XCHG
' Exchange H and L with D and E
    Case "EB"
        tstr = D
        D = H
        H = tstr
        tstr = E
        E = L
        L = tstr
        G ptr, 3, "XCHG"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' XRA Reg./Mem.
' Exclusive OR with Accumulator
    Case "A8"
        XRA B
        G ptr, 3, "XRA B"
    Case "A9"
        XRA C
        G ptr, 3, "XRA C"
    Case "AA"
        XRA D
        G ptr, 3, "XRA D"
    Case "AB"
        XRA E
        G ptr, 3, "XRA E"
    Case "AC"
        XRA H
        G ptr, 3, "XRA H"
    Case "AD"
        XRA L
        G ptr, 3, "XRA L"
    Case "AE"
    Tstates = 7
        XRA M
        G ptr, 3, "XRA M"
    Case "AF"
        XRA A
        G ptr, 3, "XRA A"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' XRI 8-bit data
' Exclusive OR Immediate with accumulator
    Case "EE"
        ptr = ptr + 1
        XRA Hcode
        G ptr - 1, 3, "XRI " & Hcode
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' XTHL
' Exchange H and L with Top of Stack
    Case "E3"
        tstr = L
        L = G(sptr, 2)
        G sptr, 2, tstr
        tstr = H
        H = G(sptr + 1, 2)
        G sptr + 1, 2, tstr
        G ptr, 3, "XTHL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Case "76"  ' HLT
        Disp(0) = "FrIE"
        Disp(1) = "nD"
        G ptr, 3, "HLT"
        Exit Sub
            
    Case "C7", "CF", "D7", "DF", "E7", "EF", "F7", "FF"
        Disp(0) = "FrIE"
        Disp(1) = "nD"
        G ptr, 3, "RST"
        Exit Sub
    Case Else
        Colourise ptr, 2, vbRed
        MsgBox "The Hex Code " & Hcode & " has no equivalent simulation." & vbNewLine _
            & "Aborting ..."
        Disp(0) = "Err "
        Disp(1) = ""
        
        Exit Sub
    End Select
    Delay Tstates * (1 / (3.255 * (10 ^ 6)))
    DoEvents
ptr = ptr + 1
If stp = True Or Go = False Then Exit Sub
Wend
Exit Sub
eHandler:
Disp(0) = "Err "
Disp(1) = ""
MsgBox Error
End Sub


Private Sub E_Change()
ChangeFlags E
E = trm(E)
E.Tag = "Changed"
End Sub

Private Sub F_Change(Index As Integer)
Grid.row = ptr
F(Index).Tag = "Changed"

Flag.ToolTipText = "Sign = " & F(7) & "," _
                    & "Zero = " & F(6) & "," _
                    & "unused" & "," _
                    & "Auxilary Carry = " & F(4) & "," _
                    & "unused" & "," _
                    & "Parity = " & F(2) & "," _
                    & "unused" & "," _
                    & "Carry = " & F(0)
Flag.Text = F(7) & F(6) & F(5) & F(4) & F(3) & F(2) & F(1) & F(0)

End Sub

Private Sub Flag_Change()
Dim i As Integer
For i = 0 To 7
    Flag.SelStart = i
    Flag.SelLength = 1
    F(7 - i).Text = Flag.SelText
Next
Hex_Flag = Hex(BtoD(Flag))

End Sub

Private Sub Form_Activate()
If loaded2 = True Then Exit Sub
processing = True
WindowState = 2
Me.Picture = LoadPicture("")
Me.PaintPicture LoadPicture(App.Path & "\images\scan0009.jpg"), 0, 0, Width, Height
Auto_Size_Frames
Disp(0) = "FrIE"
Disp(1) = "nD"
ptr = 1
startpos = CLng("&H" & "C000") - 1
endpos = CLng("&H" & "F000")

Grid.Rows = (endpos + 2) - startpos

Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

On Error GoTo NoFileSelected
    OpenFile = App.Path & "\Mem.dat"
    fnum = FreeFile
    Open OpenFile For Input As #fnum
    Input #fnum, allCells
    EditSelect_Click
    Grid.FillStyle = flexFillRepeat
    Grid.CellAlignment = 1
    Grid.FillStyle = flexFillSingle
    Grid.Clip = allCells
    Close #fnum
' Always keep some backup < '''''''''''''''''''''''''''
    allCells = Grid.Clip
    fnum = FreeFile
    Open OpenFile & ".bak" For Output As #fnum
    Write #fnum, allCells
    Close #fnum
''''''''''''''''''''''''''''''''''''''''''''''' >
    Grid.row = 1
    Grid.Col = 1
    Grid.RowSel = Grid.row
    Grid.ColSel = Grid.Col
GoTo s1
NoFileSelected:
Dim i As Long
For i = startpos To endpos
'RAM(0).AddItem Hex(i)
'RAM(1).AddItem "  "
'RAM(2).AddItem "  "
'RAM(3).AddItem "  "
'RAM(4).AddItem "  "
'Grid.Rows = i - startpos + 1
Grid.TextMatrix(i - startpos, 1) = Hex(i)
DoEvents
Next
s1:
Grid.Visible = True

For i = 0 To Buttons.Count - 1
    Buttons(i).Width = Width * Buttons(i).Width / 14790
    Buttons(i).Height = Height * Buttons(i).Height / 11820
    Buttons(i).Left = Width * Buttons(i).Left / 14790
    Buttons(i).Top = Height * Buttons(i).Top / 11820
    Buttons(i).BorderStyle = 0
    Buttons(i).Caption = ""
    Buttons(i).Visible = True
Next

For i = 0 To Display0.Count - 1
    Display0(i).Width = Width * Display0(i).Width / 14790
    Display0(i).Height = Height * Display0(i).Height / 11820
    Display0(i).Left = Width * Display0(i).Left / 14790
    Display0(i).Top = Height * Display0(i).Top / 11820
    Display0(i).Visible = True
Next
For i = 0 To Display1.Count - 1
    Display1(i).Width = Width * Display1(i).Width / 14790
    Display1(i).Height = Height * Display1(i).Height / 11820
    Display1(i).Left = Width * Display1(i).Left / 14790
    Display1(i).Top = Height * Display1(i).Top / 11820
    Display1(i).Visible = True
Next
    
    Image1.Width = Width * Image1.Width / 14790
    Image1.Height = Height * Image1.Height / 11820
    Image1.Left = Width * Image1.Left / 14790
    Image1.Top = Height * Image1.Top / 11820
    Image1.Visible = True
DoEvents
'Grid.Cols = 5
'Grid.Cols = 13
Grid.Cols = 23
NumberCells

SP = Hex(endpos)
Unload Form1
processing = False
loaded2 = True
End Sub
Sub NumberCells()
Dim i As Integer

 '   For i = 1 To Grid.Cols - 1
 '       Grid.TextMatrix(0, i) = i
 '   Next
    'For i = 1 To Grid.Cols - 1
    '    Grid.TextMatrix(i, 0) = " " & Format$(i, "000")
    'Next
    Grid.ColWidth(0) = TextWidth("A")
    Grid.ColWidth(1) = TextWidth("AAAAA")
    Grid.ColWidth(2) = TextWidth("AAA")
    
    Grid.ColWidth(4) = TextWidth("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")
    Grid.ColWidth(5) = TextWidth("AAA")
    Grid.ColWidth(6) = TextWidth("AAA")
    Grid.ColWidth(7) = TextWidth("AAA")
    Grid.ColWidth(8) = TextWidth("AAA")
    Grid.ColWidth(9) = TextWidth("AAA")
    Grid.ColWidth(10) = TextWidth("AAA")
    Grid.ColWidth(11) = TextWidth("AAA")
    Grid.ColWidth(12) = TextWidth("AAA")

    Grid.ColWidth(13) = TextWidth("AAA")
    Grid.ColWidth(14) = TextWidth("AAA")
    Grid.ColWidth(15) = TextWidth("AAA")
    Grid.ColWidth(16) = TextWidth("AAA")
    Grid.ColWidth(17) = TextWidth("AAA")
    Grid.ColWidth(18) = TextWidth("AAA")
    Grid.ColWidth(19) = TextWidth("AAA")
    Grid.ColWidth(20) = TextWidth("AAA")
    Grid.ColWidth(21) = TextWidth("AAAAA")
    Grid.ColWidth(22) = TextWidth("AAAAA")
'    Grid.ColWidth(23) = TextWidth("AAA")
    
    
    
    G 0, 1, "Address"
    G 0, 2, "Data"
    G 0, 3, "Mnemonics"
    G 0, 4, "Description"
    
    G 0, 5, "F7"
    G 0, 6, "F6"
    G 0, 7, "F5"
    G 0, 8, "F4"
    G 0, 9, "F3"
    G 0, 10, "F2"
    G 0, 11, "F1"
    G 0, 12, "F0"
    
    G 0, 13, "A"
    G 0, 14, "B"
    G 0, 15, "C"
    G 0, 16, "D"
    G 0, 17, "E"
    G 0, 18, "H"
    G 0, 19, "L"
    
    G 0, 20, "M"
    G 0, 21, "SP"
    G 0, 22, "PC"
End Sub

Private Sub EditSelect_Click()
    Grid.row = 1
    Grid.Col = 1
    Grid.RowSel = Grid.Rows - 1
    Grid.ColSel = 2 'Grid.Cols - 1
End Sub
Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
reset_back_colour
'Auto_Size_Frames
End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim tr As Integer
tr = MsgBox("Save Changes?", vbYesNoCancel)
If tr = vbNo Then End
If tr = vbCancel Then
    Cancel = 1
Exit Sub
End If

save_grid
    End

End Sub


Private Sub Frame1_Click()
If Not Frame1.Top = 0 Then
    Frame1.Move 0, 0
    Frame1.BackColor = BackColor
Else
    Auto_Size_Frames
End If

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.BackColor = RGB(247, 251, 168)
End Sub


Private Sub Frame2_Click()
If Not Frame2.Left = 0 Then
    Frame2.Left = 0
    Frame2.BackColor = BackColor
Else
    Auto_Size_Frames
End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame2.BackColor = RGB(247, 251, 168)
End Sub


Private Sub Frame3_Click()
If Not Frame3.Left = 0 Then
    Frame3.Left = 0
    Frame3.BackColor = BackColor
Else
    Auto_Size_Frames
End If

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame3.BackColor = RGB(247, 251, 168)
reset_fore_colour
End Sub


Private Sub Grid_Click()
    Label1.Caption = Grid.row & " : " & Grid.Col
    Text1.Text = Grid.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
 
End Sub


Private Sub Grid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Grid.ToolTipText = G(Grid.MouseRow, Grid.MouseCol)
End Sub


Private Sub H_Change()
ChangeFlags H
H = trm(H)
H.Tag = "Changed"
'Dim i As Integer
'        For i = 0 To RAM(0).ListCount - 1
'            If H & L = RAM(0).List(i) Then
'            GoTo s1
'            End If
'        Next
On Error Resume Next
M = G(CLng("&H" & H & L) - startpos, 2)
's1:
'M = RAM(1).List(i)
End Sub


Private Sub Hex_Flag_Change()
Hex_Flag.ToolTipText = Flag.ToolTipText
Hex_Flag.ForeColor = vbBlue
End Sub

Private Sub L_Change()
ChangeFlags L
L = trm(L)
L.Tag = "Changed"
'Dim i As Integer
'        For i = 0 To RAM(0).ListCount - 1
'            If H & L = RAM(0).List(i) Then
'            GoTo s1
'            End If
'        Next
's1:
'M = RAM(1).List(i)
On Error Resume Next
M = G(CLng("&H" & H & L) - startpos, 2)
End Sub


Private Sub Label3_Click(Index As Integer)
Select Case Index
Case 0
    save_grid
    MsgBox "Changes saved"
Case 2
    End
Case 3
    Form1.Visible = True
    Form2.Visible = False
    cnt = 2
    Auto_Size_Frames
End Select

End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Label3(Index).ForeColor = vbBlue Then
    reset_fore_colour
    Label3(Index).ForeColor = vbBlue
End If
End Sub


Private Sub M_Change()
On Error GoTo eh
'Dim i As Integer
'        For i = 0 To RAM(0).ListCount - 1
'            If H & L = RAM(0).List(i) Then
'            GoTo s1
'            End If
'        Next
's1:
'RAM(1).List(i) = M
G CLng("&H" & H & L) - startpos, 2, M
M.Tag = "Changed"
Exit Sub
eh:
Caption = Caption & " - " & Error & " - " & "During M_Change"
End Sub










Private Sub PC_Change()
On Error Resume Next
If Not PC = "0" Then
    ptr = (CLng("&H" & PC) - 1) - startpos
    PC.Tag = "Changed"
End If
End Sub

Private Sub ptr_Change()
If Not processing Then
Hcode = G(ptr, 2)
PC = G(ptr + 1, 1)
End If


End Sub


Private Sub SP_Change()
If Not SP = "0" Or Not SP = "" Then
    sptr = (CLng("&H" & SP)) - startpos
    SP.Tag = "Changed"
End If
End Sub

Private Sub sptr_Change()
SP = G(sptr, 1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim SRow, SCol As Integer

    If KeyAscii = 13 Then
        Grid.Text = Text1.Text
        SRow = Grid.row + 1
        SCol = Grid.ColSel
        If SRow = Grid.Rows Then
            SRow = Grid.FixedCols
            If SCol < Grid.Cols - Grid.FixedCols Then SCol = SCol + 1
        End If
    
        Grid.row = SRow
        Grid.Col = SCol
        Grid.RowSel = SRow
        Grid.ColSel = SCol
        Text1.Text = Grid.Text
        Text1.SetFocus
        KeyAscii = 0
    Else
        'MsgBox KeyAscii
    End If
End Sub


Public Function getme(add As String) As String ' Returns Hex Data from Address specified
getme = G((CLng("&H" & add) - startpos), 2)    'RAM(1).List(clng("&H" & add) - startpos)
End Function

Public Sub setme(add As String, data As String)
 G (CLng("&H" & add) - startpos), 2, data
End Sub

Private Function Cbin(t1 As String) As String ' Convert Decimal to Binary
Dim t2 As String
While (Not (t1) = 1) And (Not t1 = 0)
t2 = t2 & t1 Mod 2
t1 = Fix(t1 / 2)
Wend
t2 = t2 & t1
t2 = StrReverse(t2)
Cbin = Format(t2, "0#######")
End Function

Private Function G(row As Long, Col As Long, Optional data As String = "") As String
If row < 0 Then Exit Function
If Col = 3 And Not data = "" Then Update_Grid row
If data = "" Then
    G = Grid.TextMatrix(row, Col)
Else
    Grid.TextMatrix(row, Col) = data
End If
End Function

Private Sub ChangeFlags(data As String)
On Error Resume Next
If AR = True Then ' If an arithmatic operation

' < Check for Zero Flag '''''''''''''''''''''''''''''''''
    If CLng("&H" & data) = 0 Then
        F(6) = 1
    Else
        F(6) = 0
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>
' < Check for Carry Flag '''''''''''''''''''''''''''''''''
    If Len(data) = 3 Then
        F(0) = 1
        AR = False
        tstr = data
        tstr.SelStart = 0
        tstr.SelLength = 1
        tstr.SelText = ""
        data = tstr
    Else
        F(0) = 0
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>
' < Check for Sign Flag '''''''''''''''''''''''''''''''''
    tstr = Cbin(CLng("&H" & data))
    tstr.SelStart = 0
    tstr.SelLength = 1
    If tstr.SelText = 1 Then
        F(7) = 1
    Else
        F(7) = 0
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>
' < Check for Parity Flag '''''''''''''''''''''''''''''''''
    tstr = Cbin(CLng("&H" & data))
    Dim t As Integer, i As Integer
    For i = 0 To 7
        tstr.SelStart = i
        tstr.SelLength = 1
        t = t + tstr.SelText
    Next
    
    If t Mod 2 = 0 Then
        F(2) = 1
    Else
        F(2) = 0
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>
' < Check for Auxilary Carry Flag '''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''>

AR = False
End If
End Sub

Private Function trm(data As String)
If Len(data) > 2 Then
    While Len(data) > 2
        tstr = data
        tstr.SelStart = 0
        tstr.SelLength = 1
        tstr.SelText = ""
        trm = tstr.Text
    Wend
        trm = data
ElseIf Len(data) = 1 Then
    trm = "0" & data
ElseIf Len(data) = 0 Then
    trm = "00"
Else
    trm = data
End If



End Function

Private Sub CMP(TMP As String)
If CLng("&H" & A) < CLng("&H" & TMP) Then
            F(0) = 1
            F(6) = 0
        ElseIf CLng("&H" & A) > CLng("&H" & TMP) Then
            F(0) = 0
            F(6) = 0
        Else
            F(0) = 0
            F(6) = 1
        End If
End Sub

Private Function BtoD(t1 As String) As String
Dim i As Integer
Dim t2 As Long
tstr = t1
For i = 0 To 7
    tstr.SelStart = i
    tstr.SelLength = 1
    t2 = t2 + tstr.SelText * 2 ^ (7 - i)
Next
BtoD = t2
End Function
Private Function Complement(data As String) As String
        tstr = Cbin(CLng("&H" & data))
        Dim i As Integer
        For i = 0 To 7
            tstr.SelStart = i
            tstr.SelLength = 1
            If tstr.SelText = 1 Then
                tstr.SelText = 0
            Else
                tstr.SelText = 1
            End If
        Next
        Complement = Hex(BtoD(tstr))
End Function


Private Sub Subtract_SUB(data As String)
        tstr = Hex((CLng("&H" & A) - CLng("&H" & data)))
        If Len(tstr) > 2 Then
            tstr = CLng("&H" & Complement(data)) + 1 + CLng("&H" & A)
            A = Hex(tstr)
            F(0) = 1
        Else
            tstr = CLng("&H" & Complement(data)) + 1 + CLng("&H" & A)
            A = Hex(tstr)
            F(0) = 0
        End If
End Sub

Private Sub Subtract_SBB(data As String)
        tstr = Hex((CLng("&H" & A) + CLng("&H" & F(0)) - CLng("&H" & data)))
        If Len(tstr) > 2 Then
            tstr = CLng("&H" & Complement(Hex(CLng("&H" & data) + F(0)))) + 1 + CLng("&H" & A)
            A = Hex(tstr)
            F(0) = 1
        Else
            tstr = CLng("&H" & Complement(Hex(CLng("&H" & data) + F(0)))) + 1 + CLng("&H" & A)
            A = Hex(tstr)
            F(0) = 0
        End If
End Sub

Private Sub Add_Description(Hcode As String)
'< Adding Description of Instruction '''''''''''''''''''''''''''''''''''''''''''''''''''''

Select Case CLng("&H" & Hcode)
Case CLng("&H" & "88") To CLng("&H" & "8F")
    G ptr, 4, "Add Register to Accumulator with Carry"
    Tstates = 4
    Description = ""
    AR = True
Case CLng("&H" & "A0") To CLng("&H" & "A7")
    G ptr, 4, "Logical AND with Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "80") To CLng("&H" & "87")
    G ptr, 4, "Add Register to Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "40") To CLng("&H" & "7F")
    G ptr, 4, "Copy from Source to Destination"
Case CLng("&H" & "B8") To CLng("&H" & "BF")
    G ptr, 4, "Compare with Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "B0") To CLng("&H" & "B7")
    G ptr, 4, "Logically OR with Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "98") To CLng("&H" & "9F")
    G ptr, 4, "Subtract Source and Borrow from Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "90") To CLng("&H" & "97")
    G ptr, 4, "Subtract Register or Memory from Accumulator"
    Tstates = 4
    AR = True
Case CLng("&H" & "A8") To CLng("&H" & "AF")
    G ptr, 4, "Exclusive OR with Accumulator"
    Tstates = 4
    AR = True
End Select


Select Case Hcode
Case "06", "0E", "16", "1E", "26", "2E", "36", "3E"
    G ptr, 4, "Move Immediate 8-Bit"
    Tstates = 7
Case "CE":
    G ptr, 4, "Add Immediate to Accumulator with Carry"
    Tstates = 7
    AR = True
Case "C6"
    G ptr, 4, "ADD Immediate to Accumulator"
    Tstates = 7
    AR = True
Case "E6"
    G ptr, 4, "AND Immediate with Accumulator"
    Tstates = 7
    AR = True
Case "CD"
    G ptr, 4, "Unconditional Subroutine Call"
    Tstates = 18
Case "DC"
    G ptr, 4, "Call on Carry"
Case "D4"
    G ptr, 4, "Call on No Carry"
Case "F4"
    G ptr, 4, "Call on positive"
Case "FC"
    G ptr, 4, "Call on minus"
Case "EC"
    G ptr, 4, "Call on Parity Even"
Case "E4"
    G ptr, 4, "Call on Parity Odd"
Case "CC"
    G ptr, 4, "Call on Zero"
Case "C4"
    G ptr, 4, "Call on No Zero"
Case "3F"
    G ptr, 4, "Complement Carry"
Case "2F"
    G ptr, 4, "Complement Accumulator"
    Tstates = 4
Case "09", "19", "29", "39"
    G ptr, 4, "Add Register Pair to H and L Registers"
    Tstates = 10
    AR = True
Case "FE"
    G ptr, 4, "Compare Immediate with Accumulator"
    Tstates = 7
    AR = True
Case "05", "0D", "15", "1D", "25", "2D", "35", "3D"
    G ptr, 4, "Decrement source by 1"
    Tstates = 4
    AR = True
Case "0B", "1B", "2B", "3B"
    G ptr, 4, "Decrement Register Pair by 1"
    Tstates = 6
Case "04", "14", "24", "34", "0C", "1C", "2C", "3C", "4C"
    G ptr, 4, "Increment Contents of Register/Memory by 1"
    Tstates = 4
    AR = True
Case "C3"
    G ptr, 4, "Jump Unconditionaly"
    Tstates = 10
Case "DA"
    G ptr, 4, "Jump on Carry"
Case "D2"
    G ptr, 4, "Jump on No Carry"
Case "F2"
    G ptr, 4, "Jump on positive"
Case "FA"
    G ptr, 4, "Jump on minus"
Case "EA"
    G ptr, 4, "Jump on Parity Even"
Case "E2"
    G ptr, 4, "Jump on Parity Odd"
Case "CA"
    G ptr, 4, "Jump on Zero"
Case "C2"
    G ptr, 4, "Jump on No Zero"
Case "3A"
    G ptr, 4, "Load Accumulator Direct"
    Tstates = 13
Case "0A", "1A"
    G ptr, 4, "Load Accumulator Indirect"
    Tstates = 7
Case "2A"
    G ptr, 4, "Load H and L Registers Direct"
    Tstates = 16
Case "01", "21", "31", "11"
    G ptr, 4, "Load Register Pair Immediate"
    Tstates = 10
Case "00"
    G ptr, 4, "No Operation"
    Tstates = 4
Case "F6"
    G ptr, 4, "Logically OR Immediate"
    Tstates = 7
    AR = True
Case "E9"
    G ptr, 4, "Load Program Counter with HL Contents"
    Tstates = 6
Case "C1", "D1", "E1", "F1"
    G ptr, 4, "Pop off Stack to Register Pair"
    Tstates = 10
Case "C5", "D5", "E5", "F5"
    G ptr, 4, "Push Register Pair onto Stack"
    Tstates = 12
Case "17"
    G ptr, 4, "Rotate Accumulator Left through Carry"
    Tstates = 4
Case "1F"
    G ptr, 4, "Rotate Accumulator Right through Carry"
    Tstates = 4
Case "07"
    G ptr, 4, "Rotate Accumulator Left"
    Tstates = 4
Case "0F"
    G ptr, 4, "Rotate Accumulator Right"
    Tstates = 4
Case "C9"
    G ptr, 4, "Return from Subroutine Unconditionaly"
    Tstates = 10
Case "D8"
    G ptr, 4, "Return on Carry"
Case "D0"
    G ptr, 4, "Return on No Carry"
Case "F0"
    G ptr, 4, "Return on positive"
Case "F8"
    G ptr, 4, "Return on minus"
Case "E8"
    G ptr, 4, "Return on Parity Even"
Case "E0"
    G ptr, 4, "Return on Parity Odd"
Case "C8"
    G ptr, 4, "Return on Zero"
Case "C0"
    G ptr, 4, "Return on No Zero"
Case "C7", "CF", "D7", "DF", "E7", "EF", "F7", "FF"
    G ptr, 4, "Restart"
    Tstates = 12
Case "76"
    G ptr, 4, "Halt and Enter Wait State"
    Tstates = 5
Case "DE"
    G ptr, 4, "Subtract Immediate with Borrow"
    Tstates = 7
    AR = True
Case "22"
    G ptr, 4, "Stores H and L Registers Direct"
    Tstates = 16
Case "F9"
    G ptr, 4, "Copy H and L Registers to Stack Pointer"
    Tstates = 6
Case "32"
    G ptr, 4, "Store Accumulator Direct"
    Tstates = 13
Case "02", "12"
    G ptr, 4, "Store Accumulator Indirect"
    Tstates = 7
Case "37"
    G ptr, 4, "Set Carry"
    Tstates = 4
Case "D6"
    G ptr, 4, "Subtract Immediate with Accumulator"
    Tstates = 7
    AR = True
Case "EE"
    G ptr, 4, "Exclusive OR Immediate with Accumulator"
    Tstates = 7
    AR = True
Case "E3"
    G ptr, 4, "Exchange H and L with Top of Stack"
    Tstates = 16
End Select
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' >

End Sub

Private Sub CALL_UNconditional()
     
        W = Hcode               ' < Uses WZ Registers for temporarily storage
        Z = G(ptr - 1, 2)       '   of calling address
        
        sptr = sptr - 1         ' < Stack Pointer is decremented by 1
        PC.SelStart = 0
        PC.SelLength = 2
        G sptr, 2, PC.SelText   ' MSB of Program Counter is stored
        
        sptr = sptr - 1         ' < Stack Pointer is decremented by 1
        PC.SelStart = 2
        PC.SelLength = 2
        G sptr, 2, PC.SelText   ' LSB of Program Counter is stored
        
        PC = W & Z              ' < Program Counter Loaded With Calling Address
End Sub

Private Sub Colourise(row As Long, Col As Long, Colour As Long)
Grid.row = row
Grid.Col = Col
Grid.CellBackColor = Colour
End Sub

Private Sub Update_Grid(row As Long)
If ptr = 1 Then Exit Sub
G row, 13, A
G row, 14, B
G row, 15, C
G row, 16, D
G row, 17, E
G row, 18, H
G row, 19, L
G row, 20, M
G row, 21, SP
G row, 22, PC
'Select Case Index
'Case 7:
    G row, 5, F(7).Text
    
'Case 6:
    G row, 6, F(6).Text
'Case 5:
    G row, 7, F(5).Text
'Case 4:
    G row, 8, F(4).Text
'Case 3:
    G row, 9, F(3).Text
'Case 2:
    G row, 10, F(2).Text
'Case 1:
    G row, 11, F(1).Text
'Case 0:
    G row, 12, F(0).Text
'End Select

If A.Tag = "Changed" Then
    Colourise row, 13, vbCyan
    A.ForeColor = vbBlue
    A.Tag = ""
End If
If B.Tag = "Changed" Then
    Colourise row, 14, vbCyan
    B.ForeColor = vbBlue
    B.Tag = ""
End If
If C.Tag = "Changed" Then
    Colourise row, 15, vbCyan
    C.ForeColor = vbBlue
    C.Tag = ""
End If
If D.Tag = "Changed" Then
    Colourise row, 16, vbCyan
    D.ForeColor = vbBlue
    D.Tag = ""
End If
If E.Tag = "Changed" Then
    Colourise row, 17, vbCyan
    E.ForeColor = vbBlue
    E.Tag = ""
End If
If H.Tag = "Changed" Then
    Colourise row, 18, vbCyan
    H.ForeColor = vbBlue
    H.Tag = ""
End If
If L.Tag = "Changed" Then
    Colourise row, 19, vbCyan
    L.ForeColor = vbBlue
    L.Tag = ""
End If
If M.Tag = "Changed" Then
    Colourise row, 20, vbCyan
    M.ForeColor = vbBlue
    M.Tag = ""
End If
If SP.Tag = "Changed" Then
    Colourise row, 21, vbCyan
    SP.ForeColor = vbBlue
    SP.Tag = ""
End If
If PC.Tag = "Changed" Then
    Colourise row, 22, vbCyan
    PC.ForeColor = vbBlue
    PC.Tag = ""
End If

Grid.row = row
If F(7).Tag = "Changed" Then
    Grid.Col = 5
    Grid.CellBackColor = vbYellow
    F(7).Tag = ""
End If
If F(6).Tag = "Changed" Then
    Grid.Col = 6
    Grid.CellBackColor = vbYellow
    F(6).Tag = ""
End If
If F(5).Tag = "Changed" Then
    Grid.Col = 7
    Grid.CellBackColor = vbYellow
    F(5).Tag = ""
End If
If F(4).Tag = "Changed" Then
    Grid.Col = 8
    Grid.CellBackColor = vbYellow
    F(4).Tag = ""
End If
If F(3).Tag = "Changed" Then
    Grid.Col = 9
    Grid.CellBackColor = vbYellow
    F(3).Tag = ""
End If
If F(2).Tag = "Changed" Then
    Grid.Col = 10
    Grid.CellBackColor = vbYellow
    F(2).Tag = ""
End If
If F(1).Tag = "Changed" Then
    Grid.Col = 11
    Grid.CellBackColor = vbYellow
    F(1).Tag = ""
End If
If F(0).Tag = "Changed" Then
    Grid.Col = 12
    Grid.CellBackColor = vbYellow
    F(0).Tag = ""
End If
End Sub


Private Sub DAD(data1 As String, data2 As String)
        tstr = Format(Hex(CLng("&H" & data1 & data2) + CLng("&H" & H & L)), "0###")
        If CLng("&H" & tstr) > CLng("&H" & "FFFF") Then ' Check if result is larger than 16 bit
            F(0) = 1
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            L = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            H = tstr.SelText
        Else
            F(0) = 0
            tstr.SelStart = Len(tstr) - 2
            tstr.SelLength = 2
            L = tstr.SelText
            tstr.SelStart = Len(tstr) - 4
            tstr.SelLength = 2
            H = tstr.SelText
        End If
End Sub

Private Sub ORA(data As String)
A = Hex(CLng("&H" & A) Or CLng("&H" & data))
        F(4) = 0
        F(0) = 0
End Sub

Private Sub PUSH(data1 As String, data2 As String)
        sptr = sptr - 1
        G sptr, 2, data1
        sptr = sptr - 1
        G sptr, 2, data2
End Sub

Private Sub RET()
        
        sptr = sptr + 1
        PC = G(sptr, 2) & G(sptr - 1, 2)
        sptr = sptr + 1
End Sub

Private Sub XRA(data As String)
A = Hex(CLng("&H" & A) Xor CLng("&H" & data))
        F(0) = 0
        F(4) = 0
End Sub

Private Sub Auto_Size_Frames()
Frame1.Height = ScaleHeight / 3
Frame1.Move 0, -Frame1.Height + 200, ScaleWidth

Grid.Height = Frame1.Height - (Label1.Height + 200)
Grid.Width = Frame1.Width

Frame2.Move -Frame2.Width + 200, ScaleHeight - Frame2.Height

'Frame3.Height = ScaleHeight - Frame2.Height
Frame3.Move -Frame3.Width + 200, 0

Frame4.Move 0, 0, 200, ScaleHeight
End Sub

Private Sub reset_fore_colour()
Dim i As Integer
    For i = 0 To Label3.Count - 1
        Label3(i).ForeColor = vbBlack
    Next
End Sub

Private Sub reset_back_colour()
Frame1.BackColor = BackColor
Frame2.BackColor = BackColor
Frame3.BackColor = BackColor

End Sub

Private Sub save_grid()
Dim allCells As String
Dim fnum As Integer
Dim curRow, curCol As Integer

    curRow = Grid.row
    curCol = Grid.Col
    
'    CommonDialog1.DefaultExt = "GRD"
'    CommonDialog1.Action = 2
'    If CommonDialog1.FileName = "" Then Exit Sub
    EditSelect_Click
    allCells = Grid.Clip
    fnum = FreeFile
    Open OpenFile For Output As #fnum
    Write #fnum, allCells
    Close #fnum
    
'    Grid.Row = curRow
'    Grid.Col = curCol
'    Grid.RowSel = Grid.Row
'    Grid.ColSel = Grid.Col

End Sub
