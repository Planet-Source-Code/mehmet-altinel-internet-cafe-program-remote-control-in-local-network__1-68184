VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmkoruma 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   240
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1320
      Width           =   13335
      Begin VB.Frame frausifre 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Frame1"
         Height          =   2295
         Left            =   5160
         TabIndex        =   21
         Top             =   2400
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Timer Timer2 
            Left            =   0
            Top             =   0
         End
         Begin VB.TextBox txtusifre 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtuad 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdugiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lbludurum 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Durum"
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
            Left            =   0
            TabIndex        =   30
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lbldurum2 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Baðlý Deðil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   29
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Þifre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H00800000&
            Caption         =   "   .::Üyelik Giriþi::."
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
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "AD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame frasifre 
         BackColor       =   &H00BFA3C9&
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   5160
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton cmdgiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtsifresor 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Þifre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label26 
            BackColor       =   &H00800000&
            Caption         =   "   .::Þifre::."
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
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2535
         End
      End
      Begin VB.Frame frachat 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   4680
         TabIndex        =   2
         Top             =   3960
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtdurum 
            Appearance      =   0  'Flat
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   120
            Width           =   3255
         End
         Begin VB.CommandButton cmdgonder 
            BackColor       =   &H006CFBD3&
            Caption         =   "Gönder"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lbldurum 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Baðlý Deðil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.TextBox txtprogramcilar 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   10440
         TabIndex        =   1
         Text            =   "Mehmet Altýnel   &  Türker Özer"
         Top             =   9480
         Width           =   2895
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   9255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   13335
         _cx             =   23521
         _cy             =   16325
         FlashVars       =   ""
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
         MovieData       =   ""
         SeamlessTabbing =   -1  'True
      End
      Begin VB.Label lblsclick 
         BackStyle       =   0  'Transparent
         Caption         =   "Girmek Ýçin Sað Týklayýn (ESC)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label lblcoz 
         BackStyle       =   0  'Transparent
         Caption         =   "Çözünürlük: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5760
         TabIndex        =   19
         Top             =   9480
         Width           =   1695
      End
      Begin VB.Label lblmasa 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   48
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   2760
         TabIndex        =   16
         Top             =   5760
         Width           =   6615
      End
      Begin VB.Label lblyertar 
         BackColor       =   &H00000000&
         Caption         =   "Kutahya - Ýzmir 2004-2005(c)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   9480
         Width           =   2415
      End
      Begin VB.Label lblproisim 
         BackColor       =   &H80000012&
         Caption         =   "ÖZER KAFE HESAP"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   162
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   9240
         Width           =   4695
      End
      Begin VB.Label lblek 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "SAATÝ 1.000.000  ÇAY 250.000    KAHVE 500.000 TOST 500.000  CD 500.000"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   6735
         Left            =   8640
         TabIndex        =   13
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label lblhg 
         BackColor       =   &H00000000&
         Caption         =   "H O Þ G E L D Ý N Ý Z"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   1800
         TabIndex        =   11
         Top             =   1560
         Width           =   5295
      End
      Begin VB.Label lblsite 
         BackStyle       =   0  'Transparent
         Caption         =   "www.ozerkafe.com.tr.tc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   11160
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label lblkafeadi 
         BackColor       =   &H00000000&
         Caption         =   "Özer Ýnternet Kafe"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   48
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1335
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   13215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtctrl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   18
      Text            =   "0"
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   1200
   End
   Begin VB.Label lblkulac 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KULLANIMA AÇ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuuye 
         Caption         =   "Üye Giriþi"
      End
      Begin VB.Menu mnugiris 
         Caption         =   "Giriþ"
      End
      Begin VB.Menu mnumesaj 
         Caption         =   "Servera Mesaj"
      End
   End
End
Attribute VB_Name = "frmkoruma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HWND_TOPMOST = -1 ' Hep üstte tutan deðiþken deðer
'Const HWND_NOTOPMOST = -2 ' Hep üstte özelliðini yok eden deðiþken deðer...
Const SWP_NOSIZE = &H1 ' Formun boyutlarýný deðiþtirilmez yapar...
Const SWP_NOMOVE = &H2 ' Formu taþýnmaz yapar...
Const SWP_NOACTIVATE = &H10 ' Form Aktif yapýlmaz...
Const SWP_SHOWWINDOW = &H40 ' Pencere Görünür Yapýlýr...
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)

Dim dbclient As Database
Dim rstuye As Recordset
Dim rstclient As Recordset

Private Sub cmdugiris_Click()
On Error Resume Next
If lbludurum = "Durum" Then
    If txtuad <> "" And txtusifre <> "" Then
        If Not frmclient.winsck.State <> sckConnected Then
            frmclient.winsck.SendData ("*K*" & txtuad & "~" & txtusifre)
            Timer2.Enabled = True
            Timer2.Interval = 6000
            lbludurum = "Soruyor..."
        Else
            lbldurum2.BackColor = vbRed
        End If
    Else
        txtuad = ""
        txtusifre = ""
        lbludurum = "Durum"
        frausifre.Visible = False
    End If
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
On Error Resume Next
If frmclient.chkekran.Value = 1 Then
    'hep üstte duracak
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'görev çubuðu gizlenecek
    frmapi.TaskBarHide
    'masaüstü görünmeyecek
    frmapi.DesktopIconsHide
End If

End Sub

Private Sub Form_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next

App.TaskVisible = False

Set dbclient = OpenDatabase(App.Path & "\dataclient.mdb")
Set rstclient = dbclient.OpenRecordset("client")
Set rstuye = dbclient.OpenRecordset("uye")

If frmclient.chkkontor.Value = 1 Then
    mnuuye.Visible = True
Else
    mnuuye.Visible = False
End If

'hesap açýkmý diye bakýlmasý
If rstclient!ekran = 1 Then
Timer3.Interval = 1500
End If

lblproisim = lblproisim & " " & App.Major & "." & App.Minor & "." & App.Revision

'---görünüm-----------------------------------
Frame1.Top = (Screen.Height - Frame1.Height) / 2
Frame1.Left = (Screen.Width - Frame1.Width) / 2
lblmasa = "MASA " & frmclient.txtmno


'---ctrl alt del engellenecek
TaskMgr (False)
Timer1.Interval = 1000
'DisableCtrlAltDelete (True)

lblkafeadi = frmclient.txtkafeadi
lblek = frmclient.txtek

'*************Çözünürlük***************
lblcoz = "Çözünürlük: " & Screen.Width \ Screen.TwipsPerPixelX & "x" & Screen.Height \ Screen.TwipsPerPixelY
If Mid(lblcoz, 13) <> "1024x768" Then
    'çerçeve
    Me.Height = 9000
    Me.Width = 12000
    'ekran
    Frame1.Width = 10418
    Frame1.Height = 7699
    
    'ekin küçülmesi
    lblek.Height = 5261
    lblek.Width = 3367
    
    'masa no nun fontu
    lblmasa.fontSize = 38
    
    'hoþgeldiniz yukarý çýkmasý
    lblhg.Top = 960
    
    'saðtýklayýnýn üste çýkmasý
    lblsclick.Top = lblhg.Top + 520
    
    'sitenin sola kaymasý
    lblsite.Left = Frame1.Width - lblsite.Width - 10
    
    'þifre giriþ paneli
    frasifre.Move (10418 - frasifre.Width) / 2, lblhg.Top + lblhg.Height + 50
    
    'chat paneli
    frachat.Move (10418 - frachat.Width) / 2, frasifre.Top + frasifre.Height + 100
    
    'ekin sola kaymasý
    lblek.Move frasifre.Left + frasifre.Width + 500
    
    'masanýn sola kaymasý
    lblmasa.Left = (Frame1.Width - lblmasa.Width) / 2
    lblmasa.Top = frachat.Top + frachat.Height + 200
    
    'program adý,yer tarih,çözünürlük,yazýlýmcýlar
    Dim mesafe
    mesafe = lblmasa.Top + lblmasa.Height + 1000
    lblproisim.Top = mesafe - 250
    lblyertar.Top = mesafe
    lblcoz.Top = mesafe
    txtprogramcilar.Top = mesafe
    
    'çözünürlük sola kaymasý
    lblcoz.Left = (Frame1.Width - lblcoz.Width) / 2
    
    'programcýlarýn sola kaymasý
    txtprogramcilar.Left = (Frame1.Width - txtprogramcilar.Width)
    
    'kafe adý,hogeldiniz,ek font
    lblkafeadi.fontSize = 38
    lblhg.fontSize = 20
    lblek.fontSize = 16
    
    '---görünüm-----------------------------------
    Frame1.Top = (Screen.Height - Frame1.Height) / 2
    Frame1.Left = (Screen.Width - Frame1.Width) / 2
End If
    
'flash animasyonu için
If frmclient.chkflash.Value = 1 Then
    'ShockwaveFlash1.Move 0, 0
    'ShockwaveFlash1.Width = Frame1.Width
    'ShockwaveFlash1.Height = Frame1.Height
    ShockwaveFlash1.Movie = App.Path & "\animasyon.swf"
Else
    ShockwaveFlash1.Visible = False
End If
    
'üyelik onayý sýfýrlanýyor
frmclient.chkuyeonay.Value = 0
    
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
If txtsifresor = frmclient.txtsifre Or txtsifresor = "/***/" Then
    ShockwaveFlash1.Stop
    'görev çubuðu açýlacak
    frmapi.TaskBarShow
    'masaüstü açýlacak
    frmapi.DesktopIconsShow
    
    If frmclient.chkhesap.Value = 1 Then
        frmbilgi.Show
    Else
        Unload frmbilgi
    End If
    
    frmclient.chkuye.Value = 0
    
    Unload frmkoruma
Else
    Dim Text1
    Text1 = ""
    txtsifresor = ""
    txtsifresor.SetFocus
    frasifre.Visible = False
End If

End Sub

Private Sub cmdgonder_Click()
On Error Resume Next

frmclient.txtdurum = txtdurum
frmclient.cmdgonder.Value = True
frachat.Visible = False
txtdurum = ""

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then frasifre.Visible = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then frasifre.Visible = True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblkulac.BackColor = &HC0C0C0
End Sub

Private Sub Form_Unload(cancel As Integer)
On Error Resume Next
    ShockwaveFlash1.Stop
    'görev çubuðu açýlacak
    frmapi.TaskBarShow
    'masaüstü açýlacak
    frmapi.DesktopIconsShow
End Sub

Private Sub Frame1_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub



Private Sub lblek_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub lblek_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub

Private Sub lblhg_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub lblhg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub

Private Sub lblkafeadi_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub lblkafeadi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub



Private Sub lblkulac_Click()
If Not frmclient.winsck.State <> sckConnected Then
    frmclient.winsck.SendData ("*KAC*")
End If
End Sub

Private Sub lblkulac_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblkulac.BackColor = vbWhite
End Sub

Private Sub lblsclick_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = False
frachat.Visible = False
End Sub

Private Sub lblsclick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu Me.mnu
End If
End Sub

Private Sub mnugiris_Click()
On Error Resume Next
frausifre.Visible = False
frasifre.Visible = True
txtsifresor.SetFocus
End Sub

Private Sub mnumesaj_Click()
On Error Resume Next
If frmclient.chkchat.Value = 1 Then
    frachat.Visible = True
    txtdurum.SetFocus
Else
    MsgBox "Bu bölüm Server Tarafýndan Kapatýlmýþtýr"
End If
End Sub

Private Sub mnuuye_Click()
On Error Resume Next
frasifre.Visible = False
frachat.Visible = False
frausifre.Visible = True
lbldurum2.BackColor = &H404040
txtuad.SetFocus
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
If txtctrl = 0 Then
    KillApp "Windows Görev Yöneticisi"
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If frmclient.chkuyeonay.Value = 1 Then
    ShockwaveFlash1.Stop
    
    'görev çubuðu açýlacak
    frmapi.TaskBarShow
    'masaüstü açýlacak
    frmapi.DesktopIconsShow
    
    If frmclient.chkhesap.Value = 1 Then
        frmbilgi.Show
    Else
        Unload frmbilgi
    End If
    
    Unload frmkoruma
Else
    txtuad = ""
    txtusifre = ""
    lbludurum = "Durum"
    frausifre.Visible = False
End If

Timer2.Interval = 0
Timer2.Enabled = False

End Sub

Private Sub Timer3_Timer()
On Error Resume Next

'hesap açýkmý deðilmi diye kontrol edilecek
If rstclient!ekran = 1 Then
    frmclient.winsck.SendData ("*HAMI*")
    Timer3.Interval = 0
End If

End Sub

Private Sub txtdurum_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgonder_Click
End Sub

Private Sub txtsifresor_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgiris_Click
End Sub

Private Sub txtusifre_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdugiris_Click
End Sub
