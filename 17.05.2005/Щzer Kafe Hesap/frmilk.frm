VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmilk 
   Appearance      =   0  'Flat
   BackColor       =   &H00F8BA87&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Max             =   120
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      Picture         =   "frmilk.frx":0000
      ScaleHeight     =   1185
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Timer Timer2 
         Left            =   3360
         Top             =   0
      End
   End
   Begin VB.Label lblprogram 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Özer Kafe Hesap V "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
End
Attribute VB_Name = "frmilk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
lblprogram = "Özer Kafe Hesap V " & App.Major & "." & App.Minor & "." & App.Revision
Timer1.Interval = 1000
Timer2.Interval = 10
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
frmana.Show
Unload Me
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 1
End Sub
