VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Özer Kafe Client V"
   ClientHeight    =   8295
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8145
   ControlBox      =   0   'False
   Icon            =   "frmclient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frayardim 
      BackColor       =   &H00C0FFFF&
      Height          =   7575
      Left            =   30
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdyardimkapat 
         BackColor       =   &H0080FFFF&
         Caption         =   "Kapat"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   7080
         Width           =   1095
      End
      Begin VB.TextBox txtyardim 
         Height          =   6615
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Text            =   "frmclient.frx":144A
         Top             =   360
         Width           =   3725
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "YARDIM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   41
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.CheckBox chkhesapacik 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Hesap Açýk"
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
      Left            =   6240
      TabIndex        =   58
      ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   2400
   End
   Begin VB.CheckBox chkuye 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Üye"
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
      Left            =   5280
      TabIndex        =   56
      ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CheckBox chkuyeonay 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Üye Onay"
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
      Left            =   4080
      TabIndex        =   55
      ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtosdmesaj 
      Height          =   615
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   51
      Text            =   "frmclient.frx":14E8
      Top             =   6720
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   4680
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   50
      Top             =   6960
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   49
      Top             =   6960
      Width           =   375
   End
   Begin VB.Frame frasifre 
      BackColor       =   &H00ED9EDB&
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   720
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox chksifre 
         Caption         =   "Check1"
         Height          =   315
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdx 
         BackColor       =   &H006CFBD3&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtsifresor 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   26
         ToolTipText     =   "ÞÝFRE GÝRÝNÝZ"
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdgiris 
         BackColor       =   &H006CFBD3&
         Caption         =   "*GÝRÝÞ* )>)>)>"
         Height          =   375
         Left            =   600
         MouseIcon       =   "frmclient.frx":14EE
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "GÝRÝÞ"
         Top             =   720
         Width           =   1215
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
         TabIndex        =   28
         Top             =   0
         Width           =   2535
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
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.Slider sldses 
      Height          =   375
      Left            =   4080
      TabIndex        =   44
      Top             =   5520
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      Min             =   10
      Max             =   65535
      SelStart        =   32500
      Value           =   32500
   End
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdyardim 
      BackColor       =   &H006CFBD3&
      Caption         =   "?"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdprokapat 
      BackColor       =   &H000000FF&
      Caption         =   "Kapat"
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdgoster 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Göster"
      Height          =   255
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstdurum 
      Height          =   4935
      ItemData        =   "frmclient.frx":2760
      Left            =   4080
      List            =   "frmclient.frx":2762
      TabIndex        =   33
      Top             =   480
      Width           =   3975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   7920
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "27.06.2006"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "19:28"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdayarla 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ayarla"
      Height          =   375
      Left            =   960
      MouseIcon       =   "frmclient.frx":2764
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "CLIENT PROGRAMINI AYARLA"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtdene 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "1"
      ToolTipText     =   "SERVERA YAPILAN BAÐLANMA DENEMELERÝ SAYISI"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   1920
   End
   Begin VB.TextBox txtekran 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      ToolTipText     =   "MESAJLAÞMA EKRANI"
      Top             =   840
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock winsck 
      Left            =   0
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdgonder 
      BackColor       =   &H006CFBD3&
      Caption         =   "Gönder"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmclient.frx":39D6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "MESAJ GÖNDER"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtsip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "SERVER IP NUMARASI"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtbport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "BAÐLANTI PORTU"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtcip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "CLIENT IP NUMARASI"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   3735
      Begin VB.CheckBox chkarayuz 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Açýlýþta arayüz çýksýn"
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
         TabIndex        =   57
         ToolTipText     =   "WINDOWS AÇILIÞINDA MASA ÜSTÜ ARAYÜZÜ ÇIKSIN"
         Top             =   3000
         Width           =   3615
      End
      Begin VB.CheckBox chkkontor 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Kontörlü sistem devrede"
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
         TabIndex        =   54
         ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox chkhesap 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Masa üstünde hesap göster"
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
         TabIndex        =   53
         ToolTipText     =   "MASA ÜSTÜNDE HESAP GÖSTERÝLMESÝNÝ ÝSTÝYORSANIZ SEÇÝNÝZ"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CheckBox chkflash 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Flash Animasyonu"
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
         TabIndex        =   52
         ToolTipText     =   "EKRAN KORUYUCUDA FLASH ANÝMASYONU GÖSTERÝR"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.CheckBox chkeyaz 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Mesajlarý ekrana yaz"
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
         TabIndex        =   48
         ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtek 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmclient.frx":4C48
         ToolTipText     =   "EKLEMEK ÝSTEDÝKLERÝNÝZ EKRAN KORUYUCUDA GÖRÜNÜR"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtkafeadi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   ".::Özer Ýnternet Kafe::."
         ToolTipText     =   "KAFENÝZÝN ÝSMÝNÝ GÝRÝN"
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox chkchat 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Server ' la Görüþmeye Ýzin ver"
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
         TabIndex        =   8
         ToolTipText     =   "BÝLGÝSAYAR KULLANIMA KAPALIYKEN MÜÞTERÝNÝN SERVERLA  MESAJLAÞMASINI ÝSTÝYORSANIZ SEÇÝN"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.CheckBox chkekran 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "Açýlýþta Ekran Koruyucu Gelsin"
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
         TabIndex        =   5
         ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label cmdsite 
         Alignment       =   2  'Center
         BackColor       =   &H006CFBD3&
         Caption         =   "Siteler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   47
         ToolTipText     =   "YASAKLANAN SÝTELER"
         Top             =   3720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label cmdiptal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006CFBD3&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ýptal"
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
         Height          =   315
         Left            =   1920
         MouseIcon       =   "frmclient.frx":4C84
         MousePointer    =   99  'Custom
         TabIndex        =   46
         ToolTipText     =   "ÝPTAL"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label cmdkaydet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006CFBD3&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ayarlarý Kaydet"
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
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmclient.frx":5EF6
         MousePointer    =   99  'Custom
         TabIndex        =   45
         ToolTipText     =   "AYARLARI  KAYDET"
         Top             =   3600
         Width           =   1635
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Sizin Eklemek Ýstedikleriniz"
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
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Kafenizin Ýsmi Görünsün "
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
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   2175
      End
   End
   Begin MSComCtl2.UpDown updmno 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      Max             =   50
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtsifre 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   4
      ToolTipText     =   "CLIENT PROGRAMINA ÞÝFRE VERÝN AYARLARA SÝZDEN BAÞKASI ULAÞAMASIN"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtmno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      ToolTipText     =   "MASA NUMARASI SEÇÝNÝZ PROGRAM OTAMATÝK PORT NUMARASI ATAYACAKTIR"
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtdurum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   405
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      ToolTipText     =   "MESAJ YAZMA EKRANI"
      Top             =   1920
      Width           =   3735
   End
   Begin VB.PictureBox Messenger 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblguncelleme 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Son Güncelleme 17.05.2005"
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
      TabIndex        =   43
      Top             =   7560
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Çalýþan Programlar"
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
      Left            =   4080
      TabIndex        =   34
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H006CFBD3&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   4080
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lbldurum 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bað.Denemesi"
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
      Left            =   1800
      TabIndex        =   17
      ToolTipText     =   "DURUM"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "S.IP"
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
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Baðlantý Portu"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
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
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MASA NO"
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
      TabIndex        =   21
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Menu mnuclient 
      Caption         =   "Client"
      Visible         =   0   'False
      Begin VB.Menu mnuuye 
         Caption         =   "Üyelik Sistemi"
      End
      Begin VB.Menu mnucizgi 
         Caption         =   "-"
      End
      Begin VB.Menu mnuac 
         Caption         =   "Programý Aç"
      End
      Begin VB.Menu mnuekaktif 
         Caption         =   "Ekran Koruyucu Aktif"
      End
      Begin VB.Menu mnucik 
         Caption         =   "Çýkýþ"
      End
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*****ses ayarý yapmak için*****
Private Declare Function waveOutSetVolume Lib "Winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Private Declare Function waveOutGetVolume Lib "Winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
'******************************************

'*****çalýþan programlarý görmek için******
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'*******************************************

'***bilgisayarý kapatma veya yeniden baþlatmak için
'***************************************************
'Private Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const KapatAc = 2
Const Kapat = 1
'***************************************************

'***cdrom açmak veya kapatmak için***********************************
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
'********************************************************************

'***********bu kýsm msn gibi simge halinde görünmesi için***************************************************

Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uId As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

'Bu Bölümde Win32 Apisindeki  bazý sabitler kopyalanmýþtýr
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

'Bu Bölümde Win32 Apisindeki   "Shell_NotifyIcon" fonksiyonu alýnmýþ ve burdaki alias
'seçeneði kaldýrýlmýþtýr. Kaldýrmassanýz giriþ kýsmý bulunamadý þeklinde bur hata
'mesajý alýrsýnýz
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Dim Tray As NOTIFYICONDATA   'belirttiðimiz tipte bir deðiþken tanýmlýyoruz
'************************************************************************************************************************


'****************************************************************************
'ekrana yazý yazmak için
Private myText As pAttributes

Dim i
Dim dbsite As Database
Dim rstporno As Recordset
Dim rstkumar As Recordset
Dim rstvirus As Recordset

Dim dbclient As Database
Dim rstclient As Recordset
Dim rstuye As Recordset

Dim mesaj
Dim ikontor
Dim txtskontor

Private Sub chkekran_Click()
On Error Resume Next
If chkekran.Value = 1 Then
    txtkafeadi.Enabled = True
    txtek.Enabled = True
    chkflash.Enabled = True
    chkkontor.Enabled = True
    txtkafeadi.SetFocus
Else
    txtkafeadi.Enabled = False
    txtek.Enabled = False
    chkflash.Value = 0
    chkkontor.Value = 0
    chkflash.Enabled = False
    chkkontor.Enabled = False
End If
End Sub





Private Sub cmdcikis_Click()
'***görünüm*****
Me.Height = 3770
Me.Width = 4050
frayardim.Visible = False
txtsifre = ""
'***************
Me.Hide
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
If chksifre.Value = 1 Then
    If txtsifresor = txtsifre Or txtsifresor = "/***/" Then
        Me.Height = 8670
        frasifre.Visible = False
        '-------------------------------
        txtsip.Locked = False
        txtsip.BackColor = vbGreen
        txtsip.SetFocus
        cmdyardim.Enabled = False
        cmdcikis.Enabled = False
        '-------------------------------
        frmarayuz.chkduzen.Value = 1
    Else
        txtsifresor = ""
        winsck.SendData ("*M*" & "Yanlýþ þifre girþimi oldu")
        
        frmarayuz.chkduzen.Value = 0
        
        txtsifresor.SetFocus
    End If
Else
    If txtsifresor = txtsifre Then
        Dim cevap2
        cevap2 = MsgBox("Programdan çýkmanýz dahilinde serverdan kopacaktýr çýkmak istiyor musunuz?", vbYesNo + vbCritical)
        If cevap2 = vbYes Then
            Unload Me
            End
        End If
    Else
        txtsifresor = ""
        winsck.SendData ("Yanlýþ þifre giriþimi oldu")
        txtsifresor.SetFocus
    End If
End If

End Sub

Private Sub cmdkapat_Click()
Me.Height = 3770
Me.Hide
End Sub

Private Sub cmdgoster_Click()
On Error Resume Next
LISTELE
End Sub

Private Sub cmdiptal_Click()
On Error Resume Next
Me.Height = 3770
txtsifresor = ""
rstclient.MoveFirst
txtsifre = rstclient!SIFRE

txtsip.Locked = True
txtsip.BackColor = &HC0FFFF
cmdyardim.Enabled = True
cmdcikis.Enabled = True
chksifre.Value = 0

frmarayuz.chkduzen.Value = 0

cmdayarla.SetFocus
End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next
Dim cevap
cevap = MsgBox("Ayarlarý deðiþtirmek istiyor musunuz?", vbYesNo + vbInformation)

With rstclient
    If cevap = vbYes Then
    .Edit
        !cip = txtcip
        !sip = txtsip
        !Port = txtbport
        !SIFRE = txtsifre
        !kafeadi = txtkafeadi
        !ek = txtek
        !masano = txtmno
        !ekran = chkekran.Value
        !chat = chkchat.Value
        !eyaz = chkeyaz.Value
        !hesap = chkhesap.Value
        !flash = chkflash.Value
        !KONTOR = chkkontor.Value
        !arayuz = chkarayuz.Value
        
        frmarayuz.chkduzen.Value = 0
    
    .Update
        Unload Me
        Unload frmkoruma
        Me.Show
        Me.Hide
    End If
End With
End Sub

Private Sub cmdprokapat_Click()
On Error Resume Next
PROKAPAT
End Sub

Private Sub cmdsite_Click()
On Error Resume Next
MsgBox "BU BÖLÜM HAZIRLANMA AÞAMASINDADIR"
'Me.Hide
'cmdiptal_Click
'frmsite.Show
End Sub

Private Sub cmdx_Click()
On Error Resume Next
frasifre.Visible = False
txtsifresor = ""
End Sub

Private Sub cmdyardim_Click()
On Error Resume Next
frayardim.Move 0, 0
frayardim.Visible = True
Me.Height = 8670
End Sub

Private Sub cmdyardimkapat_Click()
On Error Resume Next
frayardim.Visible = False
Me.Height = 3770
Me.Width = 4050
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'-------------------------

Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

If KeyCode = vbKeyF1 Then
    If ShiftDown And CtrlDown And AltDown Then
        If Me.Height = 8130 Then
            Me.Width = 8235
            Me.Height = 7770
        End If
    End If
End If

If KeyCode = vbKeyF1 Then
    frayardim.Move 0, 0
    frayardim.Visible = True
    Me.Height = 7770
End If
'*************************************************************

If KeyCode = vbKeyControl + vbKeyEscape Then
MsgBox "mememe"
End If

End Sub

Private Sub Form_Unload(cancel As Integer)
On Error Resume Next
'DisableCtrlAltDelete (False)
TaskMgr (True)
End Sub



Private Sub Messenger_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Me.PopupMenu mnuclient
End Sub

Private Sub Messenger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Rec  As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_RBUTTONUP:
                Me.PopupMenu mnuclient
            Case WM_LBUTTONDBLCLK:
                mnuac_Click
        End Select
        Rec = False
    End If
End Sub
Private Sub cmdgonder_Click()
On Error Resume Next
BILGI_GONDER
End Sub

Private Sub cmdayarla_Click()
On Error Resume Next
frasifre.Visible = True
chksifre.Value = 1
txtsifresor.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
mnuclient.Visible = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision & "::."

'***baþlangýçta çalýþtýr****
App.TaskVisible = False

If App.PrevInstance Then
    MsgBox "Program zaten çalýþýyor !!!", vbInformation
    End
End If
'******************************
Call AddInRun
'******************************


'***************

'***************data iþlemleri*************************
Set dbsite = OpenDatabase(App.Path & "\datasite.mdb")
Set rstporno = dbsite.OpenRecordset("porno")
Set rstkumar = dbsite.OpenRecordset("kumar")
Set rstvirus = dbsite.OpenRecordset("virus")

Set dbclient = OpenDatabase(App.Path & "\dataclient.mdb")
Set rstclient = dbclient.OpenRecordset("client")
Set rstuye = dbclient.OpenRecordset("uye")

With rstclient
    txtyardim = !yardim
    txtcip = winsck.LocalIP
    txtsip = !sip
    txtbport = !Port
    txtsifre = !SIFRE
    txtkafeadi = !kafeadi
    txtek = !ek
    
    updmno.Value = CLng(Mid(txtbport, 4, 2))
    txtmno = CLng(Mid(txtbport, 4, 2))
    
    chkhesap.Value = !hesap
    chkflash.Value = !flash
    chkekran.Value = !ekran
    chkchat.Value = !chat
    chkeyaz.Value = !eyaz
    chkkontor.Value = !KONTOR
    chkarayuz.Value = !arayuz
    
End With
'*******************************************************

If chkkontor.Value = 1 Then
    mnuuye.Visible = True
    mnucizgi.Visible = True
Else
    mnuuye.Visible = False
    mnucizgi.Visible = False
End If

'*********************icon load**************************
    Tray.cbSize = Len(Tray)
    Tray.hWnd = Messenger.hWnd
    Tray.uId = 1&
    Tray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Tray.uCallBackMessage = WM_MOUSEMOVE
    Tray.hIcon = Me.Icon
    Tray.szTip = "ÖKH Client" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, Tray
    Me.Hide
'*********************************************************

'****timer baþlat ve baðlanmaya çalýþ
Timer1.Interval = 1000
Timer2.Interval = 0
ikontor = 0

'***masa üstü arayüz olayý***
If chkarayuz.Value = 1 Then
    frmarayuz.Show
End If

'****koruma olayý*******
If chkekran.Value = 1 Then
    frmkoruma.Show
    Me.Hide
End If

'ses için
sldses.Value = 65536 \ 2
 
'bilgi formu
If frmclient.chkhesap.Value = 1 Then
    frmbilgi.Show
    rstclient.MoveFirst
    rstclient.Edit
    rstclient!sinir = "0"
    rstclient.Update
End If

End Sub

Private Sub mnuac_Click()
'***görünüm*****
Me.Height = 3770
Me.Width = 4050
'***************
Me.Visible = True
cmdiptal_Click
End Sub

Private Sub mnucik_Click()
On Error Resume Next
'***görünüm*****
Me.Height = 3770
Me.Width = 4050
'***************
Me.Visible = True
frasifre.Visible = True
txtsifresor.SetFocus
cmdiptal_Click
End Sub

Private Sub mnuekaktif_Click()
On Error Resume Next
If chkuye.Value = 0 Then
    Me.Hide
    frmkoruma.Show
Else
    MsgBox "Masa ÜYE Tarafýndan Kullanýlmakta Ekran Koruyucu Aktifleþtirilemez!!!", vbInformation
End If
End Sub

Private Sub mnuuye_Click()
On Error Resume Next
frmuye.Show
End Sub

Private Sub sldses_Change()
On Error Resume Next

Dim a, i As Long
Dim tmp, vol As String

vol = sldses.Value
tmp = Right((Hex$(vol + 65536)), 4)
vol = CLng("&H" & tmp & tmp)
a = waveOutSetVolume(0, vol)

End Sub





Private Sub Timer2_Timer()
On Error Resume Next
    ikontor = ikontor + 1
    If ikontor >= txtskontor Then
        If Not frmclient.winsck.State <> sckConnected Then
            With rstuye
                .MoveFirst
                frmclient.winsck.SendData ("*HK*" & !ad & "-" & !SIFRE)
            End With
        Else
            Unload frmuye
            frmkoruma.Show
        End If
        Timer2.Interval = 0
        txtskontor = 0
        ikontor = 0
    End If
    
End Sub



Private Sub txtbport_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim CtrlDown
If Me.Height > 5000 Then
    CtrlDown = (Shift And vbCtrlMask) > 0
    If CtrlDown = True Then
        If KeyCode = vbKeyF2 Then
            txtbport.Locked = False
        End If
        If KeyCode = vbKeyF3 Then
            txtbport.Locked = True
        End If
    End If
End If
End Sub

Private Sub txtekran_Change()
On Error Resume Next
frmkoruma.txtdurum = txtekran
End Sub

Private Sub txtsifresor_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgiris_Click
End Sub

Private Sub updmno_Change()
txtmno = updmno.Value
txtbport = CLng(20000) + CLng(txtmno)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

If Not winsck.State = sckConnected Then
    winsck.Close
    lbldurum = "Bað.Denemesi"
    frmkoruma.lbldurum = lbldurum
    frmkoruma.lbldurum2 = lbldurum
    frmbilgi.lblucret = "0"
    BAGLAN
Else
    winsck.SendData ("*A*")
    lbldurum = "Baðlý"
    txtdene = 0
    
    frmkoruma.lbldurum = lbldurum
    frmkoruma.lbldurum2 = lbldurum
    
End If

End Sub

Private Sub txtdurum_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgonder_Click
End Sub

Private Sub winsck_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
'---------
winsck.GetData mesaj, vbString, bytesTotal
'---------

'üye þifre deðiþimi için yanlýþsa
If Left(mesaj, 6) = "*SFRY*" Then
    MsgBox "Ýsim ve Þifre Geçersiz !!!", vbInformation
    
    frmuye.txtad = ""
    frmuye.txtesifre = ""
    frmuye.txtysifre = ""
    frmuye.txtysifret = ""
Else


    'üye þifre deðiþimi için doðruysa
    If Left(mesaj, 6) = "*SFRD*" Then
        With rstuye
            .MoveFirst
            .Edit
                !SIFRE = frmuye.txtysifre
            .Update
        End With
        
        MsgBox "Þifreniz Deðiþtirildi yeni þifreniz (" & frmuye.txtysifre & ")", vbInformation
        
        frmuye.txtad = ""
        frmuye.txtesifre = ""
        frmuye.txtysifre = ""
        frmuye.txtysifret = ""
    Else
    
        'üye saat kontor sorgulama üyelik sisteminde
        If Left(mesaj, 3) = "*S*" Then
            frmuye.txtacilis = Mid(mesaj, 4, 5)
            frmuye.txtsure = Mid(mesaj, 10, 5)
        Else
            
            'üyelik onayý için üyelik sisteminde
            If Left(mesaj, 4) = "*OO*" Then
                frmuye.txtanakontor = Mid(mesaj, 4)
                frmuye.chkuyeonay.Value = 1
            Else
            
                'üyelik onayý için ekran koruyucuda
                If Left(mesaj, 3) = "*O*" Then
                    UYE_BASLAT
                Else
                
                    'sýnýrlama kalkmasý için
                    If Left(mesaj, 8) = "*KSINIR*" Then
                        SINIR_KALDIR
                        MASAK
                    Else
                    
                        'sýnýrlama uyarýsý için
                        If Left(mesaj, 7) = "*SINIR*" Then
                           SINIRLA
                        Else
                        
                            'clent ayarlarý için
                            If Left(mesaj, 4) = "AYAR" Then
                                CLIENT_AYARLA
                            Else
                                
                                'ctrl kapatmak için
                                If mesaj = "*CTRLKAPAT*" Then
                                    TaskMgr (False)
                                    frmkoruma.txtctrl = 0
                                Else
                                    
                                    'ctrl açmak için
                                    If mesaj = "*CTRLAC*" Then
                                        TaskMgr (True)
                                        frmkoruma.txtctrl = 1
                                    Else
                                        
                                        'ses ayarý
                                        If Left(mesaj, 5) = "*SES*" Then
                                           Dim ses As Integer
                                           sldses.Value = Mid(mesaj, 6)
                                        Else
                                            
                                            '***ucret gösterimi***************
                                            If Left(mesaj, 3) = "*U*" Then
                                                frmbilgi.lblucret = Mid(mesaj, 4)
                                                If frmbilgi.Visible = False Then frmbilgi.Show
                                            Else
                                            
                                                'program kapatma
                                                If Left(mesaj, 10) = "*PROKAPAT*" Then
                                                    lstdurum.ListIndex = Val(Mid(mesaj, 11))
                                                    PROKAPAT
                                                Else
                                                        '----
                                                    Select Case mesaj
                                                        Case "*KAPAT*": winKAPAT
                                                        Case "*YBASLAT*": winYBASLAT
                                                        Case "*PGONDER*": LISTELE
                                                        Case "*MASAAC*": MASAAC
                                                        Case "*MASAK*": MASAK
                                                        Case "*CDROMAC*": CDROMAC
                                                        Case "*CDROMKAPAT*": CDROMKAPAT
                                                        Case Else:
                                                        
                                                        If Me.Visible = True Then
                                                            txtekran.SelText = "<Server>(" & Time & ") " & mesaj + vbCrLf
                                                        Else
                                                            If frmkoruma.Visible = False Then
                                                                If rstclient!eyaz = 1 Then
                                                                    txtosdmesaj = mesaj
                                                                    myText.fontName = "Sans"
                                                                    myText.fontBold = True
                                                                    myText.fontSize = 24
                                                                    myText.fontColor = RGB(255, 0, 0)
                                                                    Set myText.scrBufferBox = Picture2
                                                                    Set myText.textBufferBox = Picture1
                                                                    myText.textBufferWidth = 1000
                                                                    myText.textBufferHeight = 500
                                                                    myText.textLocX = 300
                                                                    myText.textLocY = 50
                                                                    myText.textString = txtosdmesaj
                                                                    PrintOnScreen myText
                                                                    txtosdmesaj = ""
                                                                Else
                                                                    MsgBox "<Server>(" & Time & ") " + vbCrLf + vbCrLf & mesaj, vbInformation
                                                                End If
                                                            Else
                                                                frmkoruma.frachat.Visible = True
                                                                frmkoruma.txtdurum = "<Server>(" & Time & ") " + vbCrLf & mesaj
                                                            End If
                                                        End If
                                                    End Select
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
 End If
 
End Sub
Public Sub BILGI_GONDER()
On Error Resume Next
'----
If winsck.State <> sckConnected Then
Else
txtekran.SelText = "<*> " & txtdurum + vbCrLf
winsck.SendData ("*M*" & txtdurum)
End If
txtdurum = ""
'----
End Sub
Public Sub BAGLAN()
On Error Resume Next
'*********porta baðlanma************
With winsck
'-------baðlan---------------------
.RemotePort = txtbport
.RemoteHost = txtsip
.Connect
'----------------------------------
txtdene = Val(txtdene) + 1
End With
'***********************************
End Sub

Public Sub winKAPAT()
On Error Resume Next
ShutDownNT True
End Sub
Public Sub winYBASLAT()
On Error Resume Next
RebootNT True
End Sub
Public Sub CDROMAC()
On Error Resume Next
mciExecute ("Set CDAudio door Open")
End Sub
Public Sub CDROMKAPAT()
On Error Resume Next
mciExecute ("Set CDAudio door closed")
End Sub
Public Sub MASAAC()
On Error Resume Next
Unload frmkoruma
End Sub
Public Sub MASAK()
On Error Resume Next
Me.Hide
frmkoruma.Show

Timer2.Interval = 0
txtskontor = 0
ikontor = 0

End Sub
Public Sub LISTELE() 'çalýþan programlarý listele
lstdurum.Clear
Dim ActiveWindowhWnd, WindowTextLength As Integer, WindowCaption As String
ActiveWindowhWnd = GetWindow(Me.hWnd, 0)

While ActiveWindowhWnd <> 0
    WindowTextLength = GetWindowTextLength(ActiveWindowhWnd)
    WindowCaption = Space(WindowTextLength + 1)
    WindowTextLength = GetWindowText(ActiveWindowhWnd, WindowCaption, WindowTextLength + 1)
        
        If WindowTextLength > 0 Then
            WindowCaption = Left(WindowCaption, WindowTextLength)
            lstdurum.AddItem (WindowCaption)
        End If
    
    ActiveWindowhWnd = GetWindow(ActiveWindowhWnd, 2)
Wend
For i = 0 To lstdurum.ListCount - 1
Dim program As String
program = "*P*" & lstdurum.List(i)
winsck.SendData (program)
Next i

End Sub

Public Sub PROKAPAT()
  On Error Resume Next
  Dim Kapatildi As Boolean
  Kapatildi = KillApp(lstdurum.Text)
  If Kapatildi = True Then
    LISTELE
  End If
  
End Sub

Public Sub CLIENT_AYARLA()
On Error Resume Next
'-----------------------------------
Dim bul1, bul2, bul3, bul4, bul5, bul6, bul7, bul8, bul9
bul1 = InStr(1, mesaj, "~")
bul2 = InStr(bul1 + 1, mesaj, "~")
bul3 = InStr(bul2 + 1, mesaj, "~")
bul4 = InStr(bul3 + 1, mesaj, "~")
bul5 = InStr(bul4 + 1, mesaj, "~")
bul6 = InStr(bul5 + 1, mesaj, "~")
bul7 = InStr(bul6 + 1, mesaj, "~")
bul8 = InStr(bul7 + 1, mesaj, "~")
bul9 = InStr(bul8 + 1, mesaj, "~")
'------------------------------------

txtsifre = Mid(mesaj, 5, bul1 - 5)
chkekran.Value = Mid(mesaj, bul1 + 1, 1)
txtkafeadi = Mid(mesaj, bul2 + 1, bul3 - bul2 - 1)
txtek = Mid(mesaj, bul3 + 1, bul4 - bul3 - 1)
chkchat.Value = Mid(mesaj, bul4 + 1, 1)
chkeyaz.Value = Mid(mesaj, bul5 + 1, 1)
chkhesap.Value = Mid(mesaj, bul6 + 1, 1)
chkflash.Value = Mid(mesaj, bul7 + 1, 1)
chkkontor.Value = Mid(mesaj, bul8 + 1, 1)
chkarayuz.Value = Mid(mesaj, bul9 + 1, 1)
'***ayarlar kaydediliyor
With rstclient
        .Edit
        !SIFRE = txtsifre
        !kafeadi = txtkafeadi
        !ek = txtek
        !ekran = chkekran.Value
        !chat = chkchat.Value
        !eyaz = chkeyaz.Value
        !hesap = chkhesap.Value
        !flash = chkflash.Value
        !KONTOR = chkkontor.Value
        !arayuz = chkarayuz.Value
        .Update
End With

End Sub

Public Sub SINIRLA()
On Error Resume Next
rstclient.MoveFirst
rstclient.Edit
    rstclient!sinir = "1"
rstclient.Update
End Sub

Public Sub SINIR_KALDIR()
On Error Resume Next
rstclient.MoveFirst
rstclient.Edit
    rstclient!sinir = "0"
rstclient.Update
End Sub

Public Sub UYE_BASLAT()
On Error Resume Next
'-----------------------------------
Dim bul1, bul2, bul3, bul4, bul5, bul6, bul7, bul8
bul1 = InStr(1, mesaj, "~")
bul2 = InStr(bul1 + 1, mesaj, "~")
bul3 = InStr(bul2 + 1, mesaj, "~")
bul4 = InStr(bul3 + 1, mesaj, "~")
bul5 = InStr(bul4 + 1, mesaj, "~")
bul6 = InStr(bul5 + 1, mesaj, "~")
bul7 = InStr(bul6 + 1, mesaj, "~")
bul8 = InStr(bul7 + 1, mesaj, "~")
'------------------------------------

'uye onaylandý
If Val(Mid(mesaj, bul2 + 1, bul3 - (bul2 + 1))) > 0 Then
    
    chkuyeonay.Value = 1
    chkuye.Value = 1
    
    With rstuye
        .MoveFirst
        .Edit
            !ad = ""
            !SIFRE = ""
            !KONTOR = ""
            !acilis = ""
            !sure = ""
            !bitis = ""
        .Update
    End With
    
    With rstuye
        .MoveFirst
        .Edit
            !ad = Mid(mesaj, 4, bul1 - 4)
            !SIFRE = Mid(mesaj, bul1 + 1, bul2 - (bul1 + 1))
            !KONTOR = Mid(mesaj, bul2 + 1, bul3 - (bul2 + 1))
            !acilis = Mid(mesaj, bul3 + 1, 5)
            '!sure = Mid(mesaj, bul4 + 1, bul5 - 4)
            '!bitis = Mid(mesaj, bul5 + 1, bul6 - 4)
        .Update
        
        'kalan kontore göre sýnýrlama konuluyor
        .MoveFirst
        Dim kontorsinir
        kontorsinir = Mid(mesaj, bul2 + 1, bul3 - (bul2 + 1))
        Timer2.Interval = 10000
        txtskontor = CDbl(kontorsinir) * 6
        Timer2.Enabled = True
    End With
End If

End Sub


