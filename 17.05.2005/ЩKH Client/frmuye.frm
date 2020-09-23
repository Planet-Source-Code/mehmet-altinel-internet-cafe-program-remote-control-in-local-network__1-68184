VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmuye 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Üyelik Sistemi::."
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmuye.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcikis 
      BackColor       =   &H006CFBD3&
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "ÇIKIÞ"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pctkkontor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   1440
      ScaleHeight     =   1905
      ScaleWidth      =   3225
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox txtsonkkontor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   645
         IMEMode         =   3  'DISABLE
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "KALAN KONTÖRÜNÜZ"
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Teþekkür Ederiz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kalan Kontörünüz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame fraengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   3480
      TabIndex        =   8
      Top             =   3480
      Width           =   6015
      Begin VB.CheckBox chkuye 
         Caption         =   "uye"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkuyeonay 
         Caption         =   "uyeonay"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame frausifre 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   1800
         TabIndex        =   9
         Top             =   720
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
            TabIndex        =   12
            ToolTipText     =   "KULLANICI ÞÝFRESÝ"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtuad 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   11
            ToolTipText     =   "KULLANICI ADI"
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdugiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   28
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label4 
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
            TabIndex        =   15
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label5 
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
            TabIndex        =   14
            Top             =   0
            Width           =   2535
         End
         Begin VB.Label Label6 
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
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdhkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Hesap Kapat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "HESABI KAPAT"
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404080&
      Caption         =   "Durum"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2535
      Begin VB.TextBox txtanasure 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "BAÞLANGIÇ KONTÖRÜNÜZ"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtanakontor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "BAÞLANGIÇ KONTÖRÜNÜZ"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton cmdsorgula 
         BackColor       =   &H006CFBD3&
         Caption         =   "Sorgula"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "HESAP SORGULA"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txthkontor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "HARCANAN KONTÖR"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtkkontor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         ToolTipText     =   "KALAN KONTÖR"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtsure 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "GEÇEN SÜRE"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtacilis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "AÇILIÞ SAATÝ"
         Top             =   480
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   2520
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Harcanan Kontör"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Kalan Kontör"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "   Açýlýþ         Süre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404080&
      Caption         =   "Þifre Deðiþitir"
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
      Height          =   2175
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtad 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "KULLANICI ADI"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtysifret 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   5
         ToolTipText     =   "YENÝ ÞÝFRE TEKRAR"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmddsifre 
         BackColor       =   &H006CFBD3&
         Caption         =   "Deðiþtir"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "ÞÝFREMÝ DEÐÝÞTÝR"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtesifre 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "ESKÝ ÞÝFRE"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtysifre 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "YENÝ ÞÝFRE"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Yeni Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Yeni Þifre T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Eski Þifre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   3435
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3087
            MinWidth        =   3087
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1676
            MinWidth        =   1676
            TextSave        =   "11.02.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "00:58"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4621
            MinWidth        =   4621
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmuye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbclient As Database
Dim rstclient As Recordset
Dim rstuye As Recordset

Private Sub cmdcikis_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmddsifre_Click()
On Error Resume Next

If txtad <> "" And txtesifre <> "" And txtysifre <> "" And txtysifret <> "" Then
    If txtysifre = txtysifret Then
        frmclient.winsck.SendData ("*SFR*" & txtad & "~" & txtesifre & "~" & txtysifre)
    Else
        MsgBox "Yeni Þifreler Uymuyor", vbInformation
    End If
Else
    MsgBox "Boþluklarý doldurunuz !!!", vbInformation
End If

End Sub

Private Sub cmdhkapat_Click()
On Error Resume Next
If chkuye.Value = 1 Then
    cevap = MsgBox("Sayýn Müþterimiz Hesabýnýzý Kapatmak Ýstiyormusunuz?", vbYesNo + vbInformation)
    If cevap = vbYes Then
        If Not frmclient.winsck.State <> sckConnected Then
            'kalan kontor bildirimi
            cmdsorgula_Click
            pctkkontor.Visible = True
            fraengel.Visible = True
            cmdcikis.Enabled = False
            
            txtsonkkontor = txtkkontor
            Timer1.Interval = 5000
            
        End If
    End If
Else
    MsgBox "Üye Giriþi yapýlmadan hesap kapatýlamaz!!!"
End If
    
End Sub

Private Sub cmdsorgula_Click()
On Error Resume Next
If chkuye.Value = 1 Then
    With rstuye
    .MoveFirst
        frmclient.winsck.SendData ("*SORGU*")
    End With
Else
    StatusBar1.Panels(4).Text = "Üye Giriþi Yapýnýz !"
End If

End Sub

Private Sub cmdugiris_Click()
On Error Resume Next
If lbludurum = "Durum" Then
    If txtuad <> "" And txtusifre <> "" Then
        If Not frmclient.winsck.State <> sckConnected Then
            chkuyeonay.Value = 0
            frmclient.winsck.SendData ("*KO*" & txtuad & "~" & txtusifre)
            
            Timer2.Enabled = True
            Timer2.Interval = 8000
            
            lbludurum = "Soruyor..."
        Else
            lbludurum.BackColor = vbRed
            lbludurum = "Baðlý Deðil"
        End If
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

Set dbclient = OpenDatabase(App.Path & "\dataclient.mdb")
Set rstclient = dbclient.OpenRecordset("client")
Set rstuye = dbclient.OpenRecordset("uye")

'görünüm
fraengel.Move 0, 0

StatusBar1.Panels(4).Text = "Durum"

chkuye.Value = frmclient.chkuye.Value

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
With rstuye
    .MoveFirst
    frmclient.winsck.SendData ("*HK*" & !AD & "-" & !SIFRE & "~" & txtsonkkontor)
End With
Unload Me
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

cmdcikis.Enabled = False

If chkuyeonay.Value = 1 Then
    fraengel.Visible = False
    
    With rstuye
    .MoveFirst
        txtanakontor = !KONTOR
        If chkuye.Value = 1 Then
            cmdsorgula.Value = True
        End If
    End With
    
Else
    txtuad = ""
    txtusifre = ""
    lbludurum = "Geçersiz"
End If

cmdcikis.Enabled = True

Timer2.Interval = 0
Timer2.Enabled = False

End Sub

Private Sub txtanakontor_Change()
On Error Resume Next
If txtanakontor / 60 > 0 Then
    If (txtanakontor \ 60) < 10 Then
        HH = "0" & (txtanakontor \ 60)
    Else
        HH = (txtanakontor \ 60)
    End If
    
    If txtanakontor - ((txtanakontor \ 60) * 60) < 10 Then
        MM = "0" & txtanakontor - ((txtanakontor \ 60) * 60)
    Else
        MM = txtanakontor - ((txtanakontor \ 60) * 60)
    End If
    
    txtanasure = HH & ":" & MM
Else
    If txtanakontor - ((txtanakontor \ 60) * 60) < 10 Then
        MM = "0" & txtanakontor - ((txtanakontor \ 60) * 60)
    Else
        MM = txtanakontor - ((txtanakontor \ 60) * 60)
    End If
    
    txtanasure = "00" & ":" & MM
End If

End Sub

Private Sub txtkkontor_Change()
On Error Resume Next
txtsonkkontor = txtkkontor
End Sub

Private Sub txtsure_Change()
On Error Resume Next
txthkontor = (Val(Mid(txtsure, 1, 2)) * 60) + Val(Mid(txtsure, 4, 2))
txtkkontor = Val(txtanakontor) - Val(txthkontor)
StatusBar1.Panels(4).Text = "Hesap bilgileri"
End Sub

Private Sub txtuad_Click()
On Error Resume Next
txtuad = ""
lbludurum.BackColor = &H404040
lbludurum = "Durum"
End Sub


Private Sub txtusifre_Click()
On Error Resume Next
txtusifre = ""
lbludurum.BackColor = &H404040
lbludurum = "Durum"
End Sub

Private Sub txtusifre_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdugiris.Value = True
End Sub
