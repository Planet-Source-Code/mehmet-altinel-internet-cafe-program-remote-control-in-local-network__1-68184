VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmyedek 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Yedekle::."
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   Icon            =   "frmyedek.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "KAPAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "PROGRAMI KAPATIR"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdyedek 
      BackColor       =   &H006CFBD3&
      Caption         =   "Yedekle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "VERÝ TABANINI YEDEKLER"
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgislem 
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5235
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "20.02.2005"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lbldurum 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label lblhedef 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblkaynak 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KAYNAK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006CFBD3&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HEDEF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006CFBD3&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuyeni 
         Caption         =   "Yeni Klasör"
      End
   End
End
Attribute VB_Name = "frmyedek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstucret As Recordset
Dim dtkafe As Database
Private Sub cmdkapat_Click()
On Error Resume Next
End
End Sub

Private Sub cmdyedek_Click()
On Error Resume Next
    FileCopy lblkaynak, lblhedef
    Timer1.Interval = 1
End Sub

Private Sub Dir1_Change()
On Error Resume Next

If Len(Dir1) = 3 Then
    lblhedef = Dir1.Path & "datakafe.mdb"
Else
    lblhedef = Dir1.Path & "\datakafe.mdb"
End If

End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu mnu
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")

lblkaynak = App.Path & "\datakafe.mdb"
lblhedef = Dir1.Path
RENK_VER

dtkafe.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Shell App.Path & "\Özer Kafe Hesap.exe"
End Sub

Private Sub mnuyeni_Click()
On Error Resume Next
cevap = InputBox("Klasör adýný giriniz")
    If cevap <> "" Then
        ChDrive Drive1
        ChDir Dir1
        MkDir cevap
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If prgislem.Value < 100 Then
        prgislem.Value = prgislem + 1
        lbldurum = "Yedek alýnýyor..."
    Else
        lbldurum = "Yedeklendi"
        Timer1.Interval = 0
        prgislem.Value = 0
    End If
End Sub
Private Sub RENK_VER()
On Error Resume Next
'renk deðiþimi**************************************************************

With rstucret
.MoveFirst
Me.BackColor = !arkarenk

Dim C
For Each C In Me.Controls
    If TypeOf C Is CommandButton Then C.BackColor = !tusrenk
    If TypeOf C Is Label Then C.ForeColor = !onrenk
    If TypeOf C Is CheckBox Then C.BackColor = !arkarenk
    If TypeOf C Is CheckBox Then C.ForeColor = !onrenk
    If TypeOf C Is OptionButton Then C.BackColor = !arkarenk
    If TypeOf C Is OptionButton Then C.ForeColor = !onrenk
    If TypeOf C Is Frame Then C.BackColor = !arkarenk
Next
End With
'****************************************************************************
End Sub
