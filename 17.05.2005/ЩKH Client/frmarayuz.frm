VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmarayuz 
   Appearance      =   0  'Flat
   BackColor       =   &H00AC7222&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   12660
   ClientTop       =   -180
   ClientWidth     =   2070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtindex 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdrsec 
      Caption         =   "Sec Resim"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdsec 
      Caption         =   "Seç"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraicon 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   855
      Begin VB.Label lblicon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OYUNLAR"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Image imgicon 
         Appearance      =   0  'Flat
         Height          =   570
         Index           =   0
         Left            =   120
         MousePointer    =   99  'Custom
         Picture         =   "frmarayuz.frx":0000
         Top             =   120
         Width           =   600
      End
   End
   Begin VB.CheckBox chkduzen 
      Caption         =   "Check1"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraicon 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   855
      Begin VB.Image imgicon 
         Appearance      =   0  'Flat
         Height          =   570
         Index           =   1
         Left            =   120
         MousePointer    =   99  'Custom
         Picture         =   "frmarayuz.frx":061E
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblicon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OYUNLAR"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraicon 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   855
      Begin VB.Image imgicon 
         Appearance      =   0  'Flat
         Height          =   570
         Index           =   2
         Left            =   120
         MousePointer    =   99  'Custom
         Picture         =   "frmarayuz.frx":0C3C
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblicon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OYUNLAR"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraicon 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   3
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   855
      Begin VB.Image imgicon 
         Appearance      =   0  'Flat
         Height          =   570
         Index           =   3
         Left            =   120
         MousePointer    =   99  'Custom
         Picture         =   "frmarayuz.frx":125A
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblicon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OYUNLAR"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnu1 
      Caption         =   "mnu1"
      Visible         =   0   'False
      Begin VB.Menu mnuad 
         Caption         =   "Ad Deðiþtir"
      End
      Begin VB.Menu mnuyol 
         Caption         =   "Yol Deðiþtir"
      End
      Begin VB.Menu mnuresim 
         Caption         =   "Resim Deðiþtir"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmarayuz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbclient As Database
Dim rstarayuz As Recordset






Private Sub cmdgiris_Click()

End Sub

Private Sub cmdx_Click()
On Error Resume Next
txtsifresor = ""
frasifre.Visible = False
End Sub

Private Sub cmdrsec_Click()
On Error Resume Next
cevap = frmapi.lblyol
If cevap <> "" Then
With rstarayuz
        .MoveFirst
        For i = 0 To .RecordCount - 1
        If txtindex = !kod Then
            .Edit
                !resim = cevap
            .Update
        End If
            .MoveNext
        Next i
End With
imgicon(txtindex).Picture = LoadPicture(cevap)
End If
End Sub

Private Sub cmdsec_Click()
cevap = frmapi.lblyol
If cevap <> "" Then
With rstarayuz
        .MoveFirst
        For i = 0 To .RecordCount - 1
        If txtindex = !kod Then
            .Edit
                !yol = cevap
            .Update
        End If
            .MoveNext
        Next i
    End With
End If
End Sub

Private Sub Form_Click()
On Error Resume Next
For i = 0 To 4
    imgicon(i).BorderStyle = 0
Next i
End Sub

Private Sub Form_Load()
On Error Resume Next
Set dbclient = OpenDatabase(App.Path & "\dataclient.mdb")
Set rstarayuz = dbclient.OpenRecordset("arayuz")

With rstarayuz
    .MoveFirst
    For i = 0 To .RecordCount - 1
    lblicon(i) = !ad
    imgicon(i).Picture = LoadPicture(!resim)
    .MoveNext
    Next i
End With

'icon renkleri
For i = 0 To 6
    fraicon(i).BackColor = Me.BackColor
Next i

Timer1.Interval = 1000

frmarayuz.Move Screen.Width - Me.Width, (Screen.Height - Me.Height) - 450
End Sub

Private Sub fraicon_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
On Error Resume Next
For i = 0 To 4
    imgicon(i).BorderStyle = 0
Next i

imgicon(Index).BorderStyle = 1
End Sub

Private Sub imgicon_Click(Index As Integer)
On Error Resume Next
For i = 0 To 4
    imgicon(i).BorderStyle = 0
Next i
txtindex = Index
imgicon(Index).BorderStyle = 1
End Sub

Private Sub imgicon_DblClick(Index As Integer)
On Error Resume Next
txtindex = Index
With rstarayuz
    .MoveFirst
    For i = 0 To .RecordCount - 1
        If txtindex = !kod Then
            Shell "explorer.exe" & " " & !yol, vbNormalFocus
        End If
    .MoveNext
    Next i
End With
End Sub

Private Sub imgicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If chkduzen.Value = 1 Then
        PopupMenu mnu1
        txtindex = Index
    End If
End If
End Sub

Private Sub imgicon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
imgicon(Index).ToolTipText = "ÇÝFT TIKLAYINIZ"
End Sub

Private Sub lblicon_Click(Index As Integer)
On Error Resume Next
For i = 0 To 4
    imgicon(i).BorderStyle = 0
Next i
txtindex = Index
imgicon(Index).BorderStyle = 1
End Sub

Private Sub mnuad_Click()
On Error Resume Next
cevap = InputBox("Klasör adýný giriniz", "Ad Deðiþtir")
If cevap <> "" Then
    lblicon(txtindex) = cevap
    With rstarayuz
        .MoveFirst
        For i = 0 To .RecordCount - 1
        If txtindex = !kod Then
            .Edit
                !ad = lblicon(txtindex)
            .Update
        End If
            .MoveNext
        Next i
    End With
    
End If
End Sub

Private Sub mnuresim_Click()
On Error Resume Next
frmapi.Show
frmapi.File1.Enabled = True
frmapi.File1.Pattern = "*.bmp"
End Sub

Private Sub mnuyol_Click()
On Error Resume Next
frmapi.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
frmarayuz.Move Screen.Width - Me.Width, (Screen.Height - Me.Height) - 450
End Sub

