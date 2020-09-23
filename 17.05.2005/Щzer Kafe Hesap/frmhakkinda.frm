VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmhakkinda 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Program Hakkýnda::."
   ClientHeight    =   7935
   ClientLeft      =   8025
   ClientTop       =   3870
   ClientWidth     =   7395
   Icon            =   "frmhakkinda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox fralogo 
      BackColor       =   &H00404040&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7515
      ScaleWidth      =   7395
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkveraciklama 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "V 1.6.0 Açýklamasý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4920
         TabIndex        =   35
         Top             =   6960
         Width           =   2295
      End
      Begin VB.TextBox txtveraciklama 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   5775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Text            =   "frmhakkinda.frx":144A
         Top             =   1080
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.TextBox txtaciklama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmhakkinda.frx":1FC3
         Top             =   1080
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "www.ozerkafe.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmhakkinda.frx":253B
         MousePointer    =   99  'Custom
         TabIndex        =   21
         ToolTipText     =   "WEB SÝTEMÝZÝ ZÝYARET EDÝN DÝÐER PROGRAMLARDAN HABERDAR OLUN"
         Top             =   6960
         Width           =   4695
      End
      Begin VB.Label lblprogram 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÖZER KAFE HESAP V 1.6.0 "
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   24
            Charset         =   162
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   7095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   620
      Left            =   6120
      ScaleHeight     =   585
      ScaleWidth      =   975
      TabIndex        =   36
      Top             =   1560
      Width           =   1005
      Begin VB.Image Image4 
         Height          =   525
         Left            =   0
         Picture         =   "frmhakkinda.frx":2845
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Frame fraresim 
      BackColor       =   &H0091FBEB&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7935
      Left            =   -6840
      TabIndex        =   29
      Top             =   7560
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton cmdogiriscikis 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   31
         Top             =   120
         Width           =   255
      End
      Begin VB.Image imgasuman 
         BorderStyle     =   1  'Fixed Single
         Height          =   5415
         Left            =   120
         Picture         =   "frmhakkinda.frx":7181
         Stretch         =   -1  'True
         Top             =   1320
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label lblsumayra 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sumeyra'ya Sevgilerimle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Label lblasuman 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Asuman Hacer ÖZGÖNÜL Hocama Sevgilerimle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Image imgsumeyra 
         BorderStyle     =   1  'Fixed Single
         Height          =   7335
         Left            =   120
         Picture         =   "frmhakkinda.frx":4EC11
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   7200
      End
   End
   Begin VB.CommandButton cmdaciklama 
      BackColor       =   &H006CFBD3&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "AÇIKLAMA ÝÇÝN TIKLAYINIZ"
      Top             =   80
      Width           =   350
   End
   Begin VB.Frame fraogiris 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   4560
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdogirisac 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Özel Giriþi  AÇ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdogiriskapat 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Özel Giriþi KAPAT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtogiris 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblogiris 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Özel Atýf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdgit2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Git"
      Height          =   255
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdgit1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Git"
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      ItemData        =   "frmhakkinda.frx":604AE
      Left            =   240
      List            =   "frmhakkinda.frx":604C4
      MouseIcon       =   "frmhakkinda.frx":60527
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox txthakkinda2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmhakkinda.frx":60831
      Top             =   2400
      Width           =   5415
   End
   Begin VB.TextBox txthakkinda1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmhakkinda.frx":6089F
      Top             =   600
      Width           =   5415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   6960
      Width           =   3255
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   2880
         TabIndex        =   11
         Text            =   "9"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1800
         TabIndex        =   10
         Text            =   "6"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   720
         TabIndex        =   9
         Text            =   "3"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   2520
         TabIndex        =   8
         Text            =   "8"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1440
         TabIndex        =   7
         Text            =   "5"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   360
         TabIndex        =   6
         Text            =   "2"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   2160
         TabIndex        =   5
         Text            =   "7"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   1080
         TabIndex        =   4
         Text            =   "4"
         Top             =   0
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Text            =   "1"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Made By Med & batiidirbluessever         Copyright 2004-2005 ©"
            TextSave        =   "Made By Med & batiidirbluessever         Copyright 2004-2005 ©"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Son Güncelleme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmhakkinda.frx":6092A
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   4440
      MouseIcon       =   "frmhakkinda.frx":60C34
      MousePointer    =   99  'Custom
      Picture         =   "frmhakkinda.frx":616EE
      Stretch         =   -1  'True
      ToolTipText     =   "BARIÞ VE ÖZGÜRLÜK TÜM DÜNYA' YA ELBET BÝR GÜN HAKÝM OLACAKTIR !"
      Top             =   4200
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAMIN YAPIM AÞAMASINDA EMEÐÝ GEÇENLER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".::YAZILIMCILAR::."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7575
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   240
      MouseIcon       =   "frmhakkinda.frx":621D9
      MousePointer    =   99  'Custom
      Picture         =   "frmhakkinda.frx":624E3
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   240
      MouseIcon       =   "frmhakkinda.frx":63807
      MousePointer    =   99  'Custom
      Picture         =   "frmhakkinda.frx":63B11
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblguncel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "08.04.2005"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      MouseIcon       =   "frmhakkinda.frx":67481
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6960
      Width           =   1575
   End
End
Attribute VB_Name = "frmhakkinda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstucret As Recordset
Dim dtkafe As Database
Dim AAA As String

Private Sub chkveraciklama_Click()
On Error Resume Next
If chkveraciklama.Value = 1 Then
    txtveraciklama.Visible = True
Else
    txtveraciklama.Visible = False
End If

End Sub

Private Sub cmdaciklama_Click()
On Error Resume Next
lblprogram = "ÖZER KAFE HESAP V " & App.Major & "." & App.Minor & "." & App.Revision
chkveraciklama.Caption = "V " & App.Major & "." & App.Minor & "." & App.Revision & " Açýklamasý"
If AAA = 1 Then
    Timer1.Interval = 0
    fraogiris.Visible = True
    txtogiris.SetFocus
Else
    If cmdaciklama.Caption = "?" Then
        fralogo.Move 0, 0
        txtaciklama.Visible = True
        fralogo.Visible = True
        cmdaciklama.Caption = "X"
        Timer1.Enabled = False
    Else
        txtaciklama.Visible = False
        fralogo.Visible = False
        cmdaciklama.Caption = "?"
        Timer1.Enabled = True
    End If
End If
End Sub

Private Sub cmdaciklama_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim Tctrl, Tshift, Talt
Tctrl = (Shift And vbCtrlMask) > 0
Talt = (Shift And vbAltMask) > 0
Tshift = (Shift And vbShiftMask) > 0
If Tctrl = True And Talt = True And Tshift = True Then
    AAA = 1
End If
End Sub

Private Sub cmdgit1_Click()
On Error Resume Next
frmsohbet.Show
frmsohbet.cmbadres.Text = "http://www.mehmetaltinel.com.tr.tc"
frmsohbet.cmdGO = True
Shell "start http://www.mehmetaltinel.com.tr.tc", vbHide
End Sub

Private Sub cmdgit2_Click()
On Error Resume Next
frmsohbet.Show
frmsohbet.cmbadres.Text = "http://www.ozerkafe.com"
frmsohbet.cmdGO = True
Shell "start http://www.ozerkafe.com", vbHide
End Sub

Private Sub cmdogirisac_Click()
On Error Resume Next

If txtogiris = "asuman hacer ozgonul" Then
    fraresim.Move 0, 0
    fraresim.Visible = True
    imgasuman.Visible = True
    lblasuman.Visible = True
End If

If txtogiris = "sumeyra" Then
    fraresim.Move 0, 0
    fraresim.Visible = True
    imgsumeyra.Visible = True
    lblsumayra.Visible = True
End If
txtogiris = ""
End Sub

Private Sub cmdogiriscikis_Click()
On Error Resume Next
    
imgasuman.Visible = False
lblasuman.Visible = False
imgsumeyra.Visible = False
lblsumayra.Visible = False

fraresim.Visible = False
cmdogiriskapat_Click
End Sub

Private Sub cmdogiriskapat_Click()
On Error Resume Next
AAA = 0
fraogiris.Visible = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'-----------------------------------------------
Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyF2 Then
    If ShiftDown And CtrlDown And AltDown Then
'-----site bilgisi----------
cevap = InputBox("Güncelleme þifresi(MED)", ".::Programcý Þifre Kontrol::.")
If cevap = "/***/" Then
cevap2 = InputBox("Güncelleme Tarihi(MED)", ".::Program Son Güncelleme::.")
rstucret.MoveFirst
rstucret.Edit
rstucret![songuncelleme] = cevap2
rstucret.Update
MsgBox "Tarih Deðiþtirildi :)"
End If
'---------------------------------------
    End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Picture1.Picture = LoadPicture(App.Path & "\peace.gif")
'----------
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")
rstucret.MoveFirst
lblguncel = rstucret![songuncelleme]
'----------
AAA = 0
End Sub





Private Sub Image5_Click()

End Sub

Private Sub Label2_Click()
On Error Resume Next
frmsohbet.Show
frmsohbet.cmbadres.Text = "http://www.ozerkafe.com"
frmsohbet.cmdGO = True
Shell "start http://www.ozerkafe.com", vbHide
End Sub





Private Sub Timer1_Timer()
On Error Resume Next
Randomize Timer
Dim sayi As Integer
For i = 1 To Timer1.Interval
Text1 = Int(Rnd * (10))
Text2 = Int(Rnd * (10))
Text3 = Int(Rnd * (10))
Text4 = Int(Rnd * (10))
Text5 = Int(Rnd * (10))
Text6 = Int(Rnd * (10))
Text7 = Int(Rnd * (10))
Text8 = Int(Rnd * (10))
Text9 = Int(Rnd * (10))
Next i
End Sub

Private Sub txtogiris_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdogirisac_Click
End If
End Sub
