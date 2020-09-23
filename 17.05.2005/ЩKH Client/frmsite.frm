VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmsite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Site Seçimleri"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmsite.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Git >>>"
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   7560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2400
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2520
      Width           =   7455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   9015
      ExtentX         =   15901
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin VB.Label Label2 
      Caption         =   "Kategori"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Adres"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblguncelleme 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Zararlý Site Listenizin En Son Güncellemesi 30.11.2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   9255
   End
End
Attribute VB_Name = "frmsite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbsite As Database
Dim rstporno As Recordset
Dim rstkumar As Recordset
Dim rstvirus As Recordset
Private Sub Form_Load()
On Error Resume Next
'***************data iþlemleri*************************
Set dbsite = OpenDatabase(App.Path & "\datasite.mdb")
Set rstporno = dbsite.OpenRecordset("porno")
Set rstkumar = dbsite.OpenRecordset("kumar")
Set rstvirus = dbsite.OpenRecordset("virus")



End Sub
