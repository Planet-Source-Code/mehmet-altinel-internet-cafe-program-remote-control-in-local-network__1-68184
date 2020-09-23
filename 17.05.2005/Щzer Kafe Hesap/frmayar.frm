VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmayar 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Ayarlar::."
   ClientHeight    =   5400
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7620
   Icon            =   "frmayar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdverkontrol 
      BackColor       =   &H006CFBD3&
      Caption         =   "Yeni Versiyon Kontrol Et"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      MouseIcon       =   "frmayar.frx":144A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "VERÝLERÝMÝ YEDEKLE"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CheckBox chkotokapat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Süre dolduðunda oto. kapat"
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
      Height          =   315
      Left            =   3840
      TabIndex        =   39
      ToolTipText     =   "CLÝENTE SINIRLAMA DOLDUÐUNDA OTOMATÝK KAPATIR"
      Top             =   3360
      Width           =   3375
   End
   Begin VB.CheckBox chkuyeengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Üyelere giriþi engelle"
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
      Left            =   3840
      TabIndex        =   38
      ToolTipText     =   "ÜYELERE GÝRÝÞTE KULLANICI ÞÝFRESÝ SORAR"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CheckBox chkyyuvarla 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Hesaplarý yukarý yuvarla"
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
      Height          =   315
      Left            =   3840
      TabIndex        =   37
      ToolTipText     =   "HESAPLARI YUKARI YUVARLAR"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CheckBox chkvrenk 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Varsayýlan"
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
      Left            =   6120
      TabIndex        =   36
      ToolTipText     =   "VARSAYILAN RENK AYARLARINI UYGULAR"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   35
      Text            =   "frmayar.frx":1754
      Top             =   2760
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkotobaglan 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Açýlþta oto. clientlere baðlan"
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
      Height          =   315
      Left            =   3840
      TabIndex        =   30
      ToolTipText     =   "PROGRAM AÇILIÞINDA OTOMATÝK CLIENTLERE BAÐLAN"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.CommandButton cmddtrans 
      BackColor       =   &H006CFBD3&
      Caption         =   "Veri Transferi"
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
      Left            =   5880
      MouseIcon       =   "frmayar.frx":179E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdyedek 
      BackColor       =   &H006CFBD3&
      Caption         =   "Veri Tabaný Yedekle"
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
      Left            =   3840
      MouseIcon       =   "frmayar.frx":1AA8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "VERÝLERÝMÝ YEDEKLE"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CheckBox chkyedek 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Her kapanýþta yedek al"
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
      Left            =   3840
      TabIndex        =   13
      ToolTipText     =   "HER KAPANIÞTA YEDEKLEMEYÝ SOR"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CheckBox chkhesapac 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Hesap açýldýðýnda masa aktif"
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
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "HESAP AÇILDIÐINDA CLÝENT MASA KULLANIMA AÇ HESAP KAPATILDIÐINDA KULLANIMA KAPAT"
      Top             =   1920
      Width           =   3375
   End
   Begin MSComCtl2.UpDown updsaatucret2 
      Height          =   300
      Left            =   3360
      TabIndex        =   27
      Top             =   840
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   1000
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtsaatucret2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "SAAT ÜCRETÝ......NOT:YTL UYGULAMAK ÝÇÝN  ÖRN: SAATÝ 1 YTL ÝSE ""100"" DEÐERÝNÝ GÝRÝNÝZ"
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkasaatengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Açýlýþ saati deðiþtirilemesin"
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
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "AÇILIÞ SAATÝNÝN DEÐÝÞTÝRÝLMESÝNÝ ENGELLE"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optytl 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Yeni Türk Lirasý"
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
      TabIndex        =   7
      ToolTipText     =   "YTL OLARAK KULLANMAK ÝÇÝN SEÇÝNÝZ"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton opttl 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Türk lirasý"
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
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "TL OLARAK KULLANMAK ÝÇÝN SEÇÝNÝZ"
      Top             =   2400
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkanaraporengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Rapora giriþi engelle"
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
      Left            =   3840
      TabIndex        =   10
      ToolTipText     =   "RAPORA GÝRÝÞTE KULLANICI ÞÝFRESÝ SORAR"
      Top             =   840
      Width           =   3375
   End
   Begin VB.CheckBox chkkasaengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Kasaya giriþi engelle"
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
      Left            =   3840
      TabIndex        =   9
      ToolTipText     =   "KSAYA GÝRÝÞTE KULLANICI ÞÝFRESÝ SORAR"
      Top             =   480
      Width           =   3375
   End
   Begin VB.CheckBox chkraporengel 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Tüm hesaplarý rapora kaydet"
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
      Left            =   3840
      TabIndex        =   8
      ToolTipText     =   "NAKÝT KAPATMA OLAYINDA ÜCRET RAPORA EKLENMEDEN HESAP KAPATILMAZ"
      Top             =   120
      Width           =   3495
   End
   Begin MSComCtl2.UpDown updbirim 
      Height          =   300
      Left            =   3360
      TabIndex        =   19
      Top             =   1560
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      Max             =   1000000
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updbasucret 
      Height          =   300
      Left            =   3360
      TabIndex        =   18
      Top             =   1200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   1000
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updsaatucret 
      Height          =   300
      Left            =   3360
      TabIndex        =   17
      Top             =   480
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   1000
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updmsayisi 
      Height          =   300
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      Max             =   50
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdkaydet 
      BackColor       =   &H006CFBD3&
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
      Height          =   615
      Left            =   240
      MouseIcon       =   "frmayar.frx":1DB2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtsifre 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtbirim 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "DUYARLILIK....YTL KULLANACAKSANIZ ÖRN:0,05 LÝRA ÝÇÝN 5 GÝRÝN"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtbasucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "BAÞLANGIÇ ÜCRETÝ.....YTL ÝÇÝN ÖRN:0,25 LÝRA KULLANACKSANIZ 25 GÝRÝN"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtsaatucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "SAAT ÜCRETÝ......NOT:YTL UYGULAMAK ÝÇÝN  ÖRN: SAATÝ 1 YTL ÝSE ""100"" DEÐERÝNÝ GÝRÝNÝZ"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtmsayisi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   5025
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "20.02.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "09:21"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   6615
            MinWidth        =   6615
            Text            =   "Mehmet ALTINEL & Türker ÖZER"
            TextSave        =   "Mehmet ALTINEL & Türker ÖZER"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   5040
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Renk Ayarlarý"
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
      Left            =   3840
      TabIndex        =   34
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lbltus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tuþ"
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
      Left            =   5280
      TabIndex        =   33
      ToolTipText     =   "TUÞ RENGÝ SEÇÝNÝZ"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblonrenk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Yazý"
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
      Left            =   4560
      TabIndex        =   32
      ToolTipText     =   "YAZI RENGÝ SEÇÝNÝZ"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblarkarenk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fon"
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
      Left            =   3840
      TabIndex        =   31
      ToolTipText     =   "FON RENGÝ SEÇÝNÝZ"
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Saat Ücreti 2"
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
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "50 Masaya kadar"
      ForeColor       =   &H006CFBD3&
      Height          =   255
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Þifre "
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
      Left            =   240
      TabIndex        =   24
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Duyarlýlýk"
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
      Left            =   240
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Baþ. Ücreti"
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
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Saat Ücreti 1"
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
      Left            =   240
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Masa Sayýsý"
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
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim dtkafe As Database
Dim rstucret As Recordset

Private Sub cmddatatransferi_Click()
On Error Resume Next
frmdtrans.Show
End Sub





Private Sub cmddtrans_Click()
On Error Resume Next
frmdtrans.Show
End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next
'****
cevap = MsgBox("Ayarlarý deðiþtirmek istediðinizden eminmisiniz? Not: Programýnýz Kapanacaktýr Tekrar baþlatýnýz.", vbYesNo + vbInformation)
'____________
If cevap = vbYes Then
If txtmsayisi <> 0 And txtmsayisi <> "" And txtmsayisi <= 50 Then
    '----
    If txtsaatucret = "" Then txtsaatucret = 0
    If txtbasucret = "" Then txtbasucret = 0
    If txtbirim = "" Then txtbirim = 1
    '---
    rstucret.MoveFirst
    With rstucret
    .Edit
        If !parabirimi = 0 Then
            !msayisi = Val(txtmsayisi)
            !ucret = CLng(txtsaatucret)
            !ucret2 = CLng(txtsaatucret2)
            !basucret = CLng(txtbasucret)
            !birim = CLng(txtbirim)
        Else
            !msayisi = Val(txtmsayisi)
            !ucret = txtsaatucret
            !ucret2 = txtsaatucret2
            !basucret = txtbasucret
            !birim = txtbirim
        End If
            
            !sifre = txtsifre
            !hesapac = chkhesapac.Value
            !yedek = chkyedek.Value
            !otobaglan = chkotobaglan.Value
            !konanarapor = chkanaraporengel.Value
            !konkasa = chkkasaengel.Value
            !konrapor = chkraporengel.Value
            !konasaat = chkasaatengel.Value
            !konuye = chkuyeengel.Value
            !yyuvarla = chkyyuvarla.Value
            !parabirimi = optytl.Value
            !vrenk = chkvrenk.Value
            !otokapat = chkotokapat.Value
            
            If chkvrenk.Value = 0 Then
                !onrenk = lblonrenk.Tag
                !arkarenk = lblarkarenk.Tag
                !tusrenk = lbltus.Tag
            Else
                !onrenk = ""
                !arkarenk = ""
                !tusrenk = ""
            End If
            
    .Update
    End With
Else
    MsgBox "Yanlýþ veya boþ deðerler girdiniz!!! Ýþlem iptal edildi", vbInformation
    Exit Sub
End If
'****refresh olacak
Unload frmayar
End
'---
End If

End Sub

Private Sub cmdverkontrol_Click()
On Error Resume Next
frmversiyon.Show
End Sub

Private Sub cmdyedek_Click()
On Error Resume Next
Shell App.Path & "\ÖKH Yedekle.exe", vbNormalFocus
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
On Error Resume Next
'****
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")
'****
With rstucret
    txtmsayisi = !msayisi
    txtsaatucret = !ucret
    txtsaatucret2 = !ucret2
    txtbasucret = !basucret
    txtbirim = !birim
    txtsifre = !sifre

    chkvrenk.Value = !vrenk
    chkotokapat.Value = !otokapat
    
    lblarkarenk.Tag = !arkarenk
    lblonrenk.Tag = !onrenk
    lbltus.Tag = !tusrenk
        
    '---
    If !hesapac = 1 Then chkhesapac.Value = 1
    If !yedek = 1 Then chkyedek.Value = 1
    If !otobaglan = 1 Then chkotobaglan.Value = 1
    If !yyuvarla = 1 Then chkyyuvarla.Value = 1
    If !konrapor = 1 Then chkraporengel.Value = 1
    If !konkasa = 1 Then chkkasaengel.Value = 1
    If !konanarapor = 1 Then chkanaraporengel.Value = 1
    If !konasaat = 1 Then chkasaatengel.Value = 1
    If !konuye = 1 Then chkuyeengel.Value = 1
    If !parabirimi = True Then optytl.Value = True
    '---
    
End With
'***
'küçük kontroller
updmsayisi.Value = Val(txtmsayisi)
'****************

If chkvrenk.Value = 0 Then
    RENK_VER
End If

'renk tuþlarýnýn yazý rengi
lblonrenk.ForeColor = vbBlack
lbltus.ForeColor = vbBlack
lblarkarenk.ForeColor = vbBlack

End Sub

Private Sub lblarkarenk_Click()
On Error Resume Next
CommonDialog1.Action = 3
lblarkarenk.BackColor = CommonDialog1.Color
lblarkarenk.Tag = CommonDialog1.Color
End Sub

Private Sub lblonrenk_Click()
On Error Resume Next
CommonDialog1.Action = 3
lblonrenk.BackColor = CommonDialog1.Color
lblonrenk.Tag = CommonDialog1.Color
End Sub

Private Sub lbltus_Click()
On Error Resume Next
CommonDialog1.Action = 3
lbltus.BackColor = CommonDialog1.Color
lbltus.Tag = CommonDialog1.Color
End Sub

Private Sub opttl_Click()
On Error Resume Next
rstucret.MoveFirst
rstucret.Edit
rstucret!parabirimi = 0
rstucret.Update
End Sub

Private Sub optytl_Click()
On Error Resume Next
rstucret.MoveFirst
rstucret.Edit
rstucret!parabirimi = -1
rstucret.Update
End Sub

Private Sub txtbasucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
txtbasucret = Format(txtbasucret, "#00,0")
txtbasucret.SelStart = Len(txtbasucret)
End If
End Sub

Private Sub txtbirim_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
txtbirim = Format(txtbirim, "#00,0")
txtbirim.SelStart = Len(txtbirim)
End If
End Sub

Private Sub txtmsayisi_Change()
On Error Resume Next
If txtmsayisi > 50 Then txtmsayisi = 50
End Sub

Private Sub txtsaatucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
txtsaatucret = Format(txtsaatucret, "#00,0")
txtsaatucret.SelStart = Len(txtsaatucret)
End If
End Sub

Private Sub txtsaatucret2_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
txtsaatucret2 = Format(txtsaatucret2, "#00,0")
txtsaatucret2.SelStart = Len(txtsaatucret2)
End If
End Sub

Private Sub updbasucret_Change()
On Error Resume Next
txtbasucret = Val(updbasucret.Value) * CDbl(txtbirim)
End Sub

Private Sub updbirim_Change()
On Error Resume Next
txtbirim = updbirim.Value
End Sub

Private Sub updmsayisi_Change()
On Error Resume Next
txtmsayisi = updmsayisi.Value
End Sub

Private Sub updsaatucret_Change()
On Error Resume Next
txtsaatucret = Val(updsaatucret.Value) * CDbl(txtbirim)
End Sub

Private Sub updsaatucret2_Change()
On Error Resume Next
txtsaatucret2 = Val(updsaatucret2.Value) * CDbl(txtbirim)
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
Next
End With
'****************************************************************************

End Sub

