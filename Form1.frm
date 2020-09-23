VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   9570
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":212A
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6450
      Left            =   6675
      TabIndex        =   24
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox Speed 
         Height          =   315
         ItemData        =   "Form1.frx":10F9EC
         Left            =   1850
         List            =   "Form1.frx":10F9F6
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   250
         Width           =   735
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   330
         Left            =   1560
         TabIndex        =   33
         Top             =   1395
         Width           =   1020
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         LargeChange     =   5
         Left            =   1680
         Max             =   5
         Min             =   100
         SmallChange     =   5
         TabIndex        =   32
         Top             =   1800
         Value           =   5
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   360
         Left            =   1920
         TabIndex        =   31
         Text            =   "000"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   480
         Top             =   0
      End
      Begin VB.PictureBox pic 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   80
         ScaleHeight     =   3375
         ScaleWidth      =   2595
         TabIndex        =   27
         Top             =   3000
         Width           =   2600
         Begin VB.Label finish 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Congratulations!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   -56
            TabIndex        =   28
            Top             =   1080
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Image Image1 
            Height          =   3315
            Left            =   0
            MouseIcon       =   "Form1.frx":10FA06
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            ToolTipText     =   "Nabeel Hosny Cairo / 2006 Click to Exit"
            Top             =   0
            Width           =   2595
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00008000&
         Caption         =   "&Select a picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   675
         Width           =   2520
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00008000&
         Caption         =   "&Four Pics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   25
         Top             =   250
         Value           =   -1  'True
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin PicClip.PictureClip Clip1 
         Left            =   960
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Rows            =   4
         Cols            =   4
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   ".45"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1920
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label CmdSolve 
         BackStyle       =   0  'Transparent
         Caption         =   "&Solve Me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   200
         MouseIcon       =   "Form1.frx":10FD10
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   1350
         Width           =   1440
      End
      Begin VB.Label CmdSort 
         BackStyle       =   0  'Transparent
         Caption         =   "&Scramble"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   200
         MouseIcon       =   "Form1.frx":11001A
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   4
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1005
         Index           =   2
         Left            =   80
         Shape           =   4  'Rounded Rectangle
         Top             =   1300
         Width           =   2595
      End
      Begin VB.Label lblElapsed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   2550
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   4
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1005
         Index           =   0
         Left            =   80
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   2595
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   4
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   1
         Left            =   80
         Shape           =   4  'Rounded Rectangle
         Top             =   2425
         Width           =   2595
      End
   End
   Begin VB.PictureBox PicPlate 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   120
      ScaleHeight     =   430
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   0
      Top             =   120
      Width           =   6450
      Begin VB.PictureBox PicsSlider 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   -1650
         MouseIcon       =   "Form1.frx":11016C
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   650
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   9750
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Index           =   0
            Left            =   0
            MouseIcon       =   "Form1.frx":1105AE
            MousePointer    =   99  'Custom
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   7
            Top             =   0
            Width           =   1650
         End
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Index           =   1
            Left            =   1650
            MouseIcon       =   "Form1.frx":1109F0
            MousePointer    =   99  'Custom
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   6
            Top             =   0
            Width           =   1650
         End
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Index           =   2
            Left            =   3300
            MouseIcon       =   "Form1.frx":110E32
            MousePointer    =   99  'Custom
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   5
            Top             =   0
            Width           =   1650
         End
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Index           =   3
            Left            =   4950
            MouseIcon       =   "Form1.frx":111274
            MousePointer    =   99  'Custom
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   4
            Top             =   0
            Width           =   1650
         End
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1650
            Index           =   4
            Left            =   6360
            MouseIcon       =   "Form1.frx":1116B6
            MousePointer    =   99  'Custom
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   3
            Top             =   840
            Width           =   1650
         End
         Begin VB.PictureBox PicsInSlider 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1500
            Index           =   5
            Left            =   8250
            MouseIcon       =   "Form1.frx":111AF8
            MousePointer    =   99  'Custom
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   100
            TabIndex        =   2
            Top             =   0
            Width           =   1500
         End
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   0
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   23
         Tag             =   "0000"
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   1
         Left            =   1650
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   22
         Tag             =   "0100"
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   2
         Left            =   3300
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   21
         Tag             =   "0200"
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   3
         Left            =   4950
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   20
         Tag             =   "0300"
         Top             =   0
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   4
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   19
         Tag             =   "0004"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   5
         Left            =   1650
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   18
         Tag             =   "0104"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   6
         Left            =   3300
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   17
         Tag             =   "0204"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   7
         Left            =   4950
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   16
         Tag             =   "0304"
         Top             =   1650
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   8
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   15
         Tag             =   "0008"
         Top             =   3300
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   9
         Left            =   1650
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   14
         Tag             =   "0108"
         Top             =   3300
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   10
         Left            =   3300
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   13
         Tag             =   "0208"
         Top             =   3300
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   11
         Left            =   4950
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   12
         Tag             =   "0308"
         Top             =   3300
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   12
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   11
         Tag             =   "0012"
         Top             =   4950
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   13
         Left            =   1650
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   10
         Tag             =   "0112"
         Top             =   4950
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   14
         Left            =   3300
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   9
         Tag             =   "0212"
         Top             =   4950
         Width           =   1500
      End
      Begin VB.PictureBox PicUser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Index           =   15
         Left            =   4950
         MousePointer    =   99  'Custom
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   8
         Tag             =   "0312"
         Top             =   4950
         Width           =   1500
      End
   End
   Begin VB.Image PicSource 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   0
      Left            =   9525
      MouseIcon       =   "Form1.frx":111F3A
      Picture         =   "Form1.frx":112804
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image PicSource 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   1
      Left            =   9525
      MouseIcon       =   "Form1.frx":1132BA
      Picture         =   "Form1.frx":113B84
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image PicSource 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   2
      Left            =   9525
      MouseIcon       =   "Form1.frx":1143B3
      Picture         =   "Form1.frx":114C7D
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image PicSource 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   3
      Left            =   9525
      MouseIcon       =   "Form1.frx":115667
      Picture         =   "Form1.frx":115F31
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image PicSource 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   4
      Left            =   9525
      MouseIcon       =   "Form1.frx":116781
      Picture         =   "Form1.frx":116BC3
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim ScrollDirection As Byte
Dim RowOrColNumber As Byte

Dim MyTime As Long, Level, movment, i As Integer

Private Sub CmdSort_Click()
Randomize: List1.Clear: CmdSolve.Enabled = False
movment = 0: finish.Visible = False
Option1_Click (Val(Label2.Caption))
For i = 1 To Text1.Text
ScrollDirection = Int(Rnd * 4 + 1)
PicUser_Click (Int(Rnd * 16))
DoEvents
Next
 CmdSolve.Enabled = True
End Sub



Private Sub Form_Load()
SetWindowRgn Frame1.hwnd, CreateRoundRectRgn(0, 0, Frame1.Width, Frame1.Height, 50, 50), True
SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 50, 50), True
SetWindowRgn PicPlate.hwnd, CreateRoundRectRgn(0, 0, PicPlate.Width, PicPlate.Height, 50, 50), True
SetWindowRgn pic.hwnd, CreateRoundRectRgn(0, 0, pic.Width / Screen.TwipsPerPixelX, pic.Height / Screen.TwipsPerPixelY, 50, 50), True
VScroll1_Change
movment = 0
Speed.ListIndex = 0
Option1_Click (0)
End Sub




Private Sub Image1_Click()
End
End Sub

Private Sub Option1_Click(Index As Integer)
Label2.Caption = Index
finish.Visible = False
Select Case Index
Case 0
Image1.Picture = PicSource(4).Picture
For i = 0 To 3
PicUser(i).Picture = PicSource(0).Picture
PicUser(i + 4).Picture = PicSource(1).Picture
PicUser(i + 8).Picture = PicSource(2).Picture
PicUser(i + 12).Picture = PicSource(3).Picture
Next i
Option1(0).Value = False
Case 1
CommonDialog1.CancelError = False
   CommonDialog1.Filter = "Picture Files (*.jpg, *gif, and " & _
      "others)|*.jpg;*.jpeg;*.gif;*.bmp;*.cur"
   CommonDialog1.FilterIndex = 1
   CommonDialog1.ShowOpen
   If CommonDialog1.FileName = "" Then Exit Sub
 Clip1.Picture = LoadPicture(CommonDialog1.FileName)
 Image1.Picture = Clip1.Picture
 For i = 0 To 15
       Clip1.StretchX = PicUser(i).ScaleWidth
       Clip1.StretchY = PicUser(i).ScaleHeight
        PicUser(i) = Clip1.GraphicCell(i)
      Next i
   Option1(1).Value = False
      End Select
End Sub

 Sub Scrol()
Dim H As Byte
Dim i As Byte
'1=left  2=up    3=right  4=down
Select Case ScrollDirection
       Case 1
           While (PicsSlider.Left > -210)
              PicsSlider.Left = PicsSlider.Left - 1: Delay Val(Label1.Caption)
              
           Wend
           H = 2
              For i = RowOrColNumber To RowOrColNumber + 3
              PicUser(i).Picture = PicsInSlider(H).Picture
              H = H + 1
              Next i
              Call ResetParan
           
           
       Case 2
           While (PicsSlider.Top > -210)
              PicsSlider.Top = PicsSlider.Top - 1: Delay Val(Label1.Caption)
           Wend
              H = 2
              For i = RowOrColNumber To RowOrColNumber + 12 Step 4
              PicUser(i).Picture = PicsInSlider(H).Picture
              H = H + 1
              Next i
              Call ResetParan
           
       Case 3
           While (PicsSlider.Left < 0)
              PicsSlider.Left = PicsSlider.Left + 1: Delay Val(Label1.Caption)
           Wend
              H = 0
              For i = RowOrColNumber To RowOrColNumber + 3
              PicUser(i).Picture = PicsInSlider(H)
              H = H + 1
              Next i
              Call ResetParan
           
      Case 4
           While (PicsSlider.Top < 0)
              PicsSlider.Top = PicsSlider.Top + 1: Delay Val(Label1.Caption)
           Wend
              H = 0
              For i = RowOrColNumber To RowOrColNumber + 12 Step 4
              PicUser(i).Picture = PicsInSlider(H).Picture
              H = H + 1
              Next i
              Call ResetParan
           
End Select

End Sub
Sub ResetParan()

PicsSlider.Visible = False
PicsSlider.Left = -110
End Sub

Private Sub CmdSolve_Click()
'1=left  2=up    3=right  4=down
Dim Solnouber, helpindex As Integer
CmdSort.Enabled = False: VScroll1.Enabled = False

Solnouber = List1.ListCount
 
For i = Solnouber - 1 To 0 Step -1
helpindex = Val(Left(List1.List(i), 2))
Select Case Right(List1.List(i), 1)
Case Is = "L": ScrollDirection = 1: SetCursorPos 500, 140 + 110 * (helpindex \ 4)
Case Is = "U": ScrollDirection = 2: SetCursorPos 140 + 110 * (helpindex Mod 4), 500
Case Is = "R": ScrollDirection = 3: SetCursorPos 100, 140 + 110 * (helpindex \ 4)
Case Is = "D": ScrollDirection = 4: SetCursorPos 140 + 110 * (helpindex Mod 4), 100
End Select
List1.Selected(i) = True
PicUser_Click (Left(List1.List(i), 2))

Next
finish.Visible = True
CmdSort.Enabled = True: VScroll1.Enabled = True


End Sub

Private Sub PicUser_Click(Index As Integer)
Dim i, l, cm As Integer
Dim H As Byte, Direction As String

'1=left  2=up    3=right  4=down
    Select Case ScrollDirection
           Case 1, 3
           If ScrollDirection = 1 Then Direction = "R"
           If ScrollDirection = 3 Then Direction = "L"
                PicsSlider.Width = 650
                PicsSlider.Height = 100
                PicsSlider.Left = -110
                H = 0
                For i = 0 To 650 Step 110
                PicsInSlider(H).Top = 0
                PicsInSlider(H).Left = i
                H = H + 1
                Next i
                RowOrColNumber = Val(Mid$(PicUser(Index).Tag, 3, 2))
                PicsSlider.Top = PicUser(RowOrColNumber).Top
                For i = 1 To 4
                PicsInSlider(i).Picture = PicUser(RowOrColNumber + (i - 1)).Picture
                Next i
                PicsInSlider(0).Picture = PicUser(RowOrColNumber + 3).Picture
                PicsInSlider(5).Picture = PicUser(RowOrColNumber).Picture
                PicsSlider.Visible = True
                Scrol
           Case 2, 4
           If ScrollDirection = 2 Then Direction = "D"
           If ScrollDirection = 4 Then Direction = "U"
                PicsSlider.Width = 100
                PicsSlider.Height = 650
                PicsSlider.Top = -110
                H = 0
                For i = 0 To 650 Step 110
                PicsInSlider(H).Top = i
                PicsInSlider(H).Left = 0
                H = H + 1
                Next i
                RowOrColNumber = Val(Mid$(PicUser(Index).Tag, 1, 2))
                PicsSlider.Left = PicUser(RowOrColNumber).Left
                For i = 1 To 4
                PicsInSlider(i).Picture = PicUser(RowOrColNumber + ((i - 1) * 4)).Picture
                Next i
                PicsInSlider(0).Picture = PicUser(12 + RowOrColNumber).Picture
                PicsInSlider(5).Picture = PicUser(RowOrColNumber).Picture
                PicsSlider.Visible = True
                Scrol
    End Select
If Index > 9 Then
List1.AddItem Index & " " & Direction
Else
List1.AddItem "0" & Index & " " & Direction
End If
List1.Selected(movment) = True
movment = movment + 1
End Sub
Private Sub PicUser_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   Select Case Y - X
           Case Is < 0
                Select Case (100 - Y) - X
                       Case Is < 0
                            PicUser(Index).MouseIcon = PicSource(1).MouseIcon  'right
                            ScrollDirection = 3    ' right =3
                       Case Else
                            PicUser(Index).MouseIcon = PicSource(2).MouseIcon  'up
                            ScrollDirection = 2    '  up =2
                       End Select
           Case Else
                Select Case (100 - Y) - X
                       Case Is < 0
                            PicUser(Index).MouseIcon = PicSource(3).MouseIcon  ' down
                            ScrollDirection = 4    ' down =4
                       Case Else
                            PicUser(Index).MouseIcon = PicSource(0).MouseIcon ' left
                            ScrollDirection = 1    ' left =1
                End Select
    End Select
 

End Sub

Private Sub Speed_Click()
' fast = 0  .45   Slow =1  .6
Label1.Caption = 0.45 + Speed.ListIndex * 0.15
End Sub

Private Sub Timer_Timer()
 Dim t As Date
    Dim M, S As Integer
    
    MyTime = MyTime + 1
       
    t = TimeSerial(0, 0, MyTime)
    lblElapsed.Caption = IIf(Level = 0, "Easy", IIf(Level = 1, "Normal", "Hard")) & " - " & Format(movment, "000") & " - " & Format(t, "hh:nn:ss")

End Sub





Public Sub Delay(S As Single)
Dim running As Boolean
    Dim N As Long
  
    N = 0
       N = GetTickCount + S
    Do
        DoEvents
    Loop Until GetTickCount >= N
    
    
End Sub

Private Sub VScroll1_Change()
Text1.Text = Format(VScroll1.Value, "000")
Select Case VScroll1.Value
Case Is <= 25: Level = 0
Case 30 To 50: Level = 1
Case Else: Level = 2
End Select
End Sub
