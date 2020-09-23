VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmuye 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Üyeler::."
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmuye.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frasifre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8295
      Left            =   4080
      TabIndex        =   45
      Top             =   8160
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00BFA3C9&
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   1920
         TabIndex        =   46
         Top             =   3240
         Width           =   2535
         Begin VB.CommandButton cmdgiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtsifre 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   0
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
            TabIndex        =   48
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   47
            Top             =   0
            Width           =   2535
         End
      End
   End
   Begin VB.CommandButton cmdcikar 
      BackColor       =   &H006CFBD3&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmuye.frx":144A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "KONTÖR ÇIKAR"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdekle 
      BackColor       =   &H006CFBD3&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      MaskColor       =   &H8000000E&
      MouseIcon       =   "frmuye.frx":1754
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "KONTÖR EKLE"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtukontor2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3600
      TabIndex        =   59
      Text            =   "0"
      ToolTipText     =   "KONTOR (DK)"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtfiyat2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3600
      TabIndex        =   57
      ToolTipText     =   "SAAT ÜCRETÝNE GÖRE TEKABÜL EDEN FÝYAT (EN SON SATILAN KONTÖR ÝÞLEMÝNDEN)"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox txtusure2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   56
      ToolTipText     =   "KONTORÜN KARÞILIK GELDÝÐÝ SURE"
      Top             =   5520
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   8235
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3351
            MinWidth        =   3351
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "19.02.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1094
            MinWidth        =   1094
            TextSave        =   "14:18"
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
   Begin VB.ComboBox cmbay 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmuye.frx":1A5E
      Left            =   840
      List            =   "frmuye.frx":1A86
      TabIndex        =   50
      Text            =   "Ay Seçiniz"
      ToolTipText     =   "ÝÞLEM YAPILACAK AY'I SEÇÝNÝZ"
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txttopmiktar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      ToolTipText     =   "TÜM ÜYELERDEN PEÞÝN ALINAN ÜCRETLER TOPLAMI (BU AY ÝÇÝN)"
      Top             =   7680
      Width           =   2415
   End
   Begin VB.CheckBox chkkayit 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4200
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Frame frasorgu 
      BackColor       =   &H00404080&
      Caption         =   "SORGULA"
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
      Height          =   1815
      Left            =   2640
      TabIndex        =   38
      Top             =   6240
      Width           =   3495
      Begin VB.TextBox txtsukontor 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "KALAN KONTÖRÜ"
         Top             =   960
         Width           =   1935
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
         Left            =   1080
         MaskColor       =   &H00C0E0FF&
         MouseIcon       =   "frmuye.frx":1B05
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "ÜYE SORGULA"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtsuad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         ToolTipText     =   "ÜYE ADI"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtsusifre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         ToolTipText     =   "ÜYE ÞÝFRESÝ"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblsudurum 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H006CFBD3&
         Height          =   495
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "DURUM"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "DK."
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
         Left            =   3120
         TabIndex        =   42
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "KONTOR"
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
         TabIndex        =   41
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
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
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ÞÝFRE"
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
         TabIndex        =   39
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraubilgi 
      BackColor       =   &H00404080&
      Enabled         =   0   'False
      Height          =   4455
      Left            =   2640
      TabIndex        =   27
      ToolTipText     =   "MÜÞTERÝ BÝLGÝLERÝ"
      Top             =   360
      Width           =   3495
      Begin VB.TextBox txttopfiyat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "BU ÜYEYE ÞÝMDÝYE KADAR SATILAN KONTÖR TOPLAMI"
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtusifre2 
         Height          =   285
         Left            =   1080
         TabIndex        =   54
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker txttarih 
         Height          =   300
         Left            =   1080
         TabIndex        =   53
         ToolTipText     =   "ÝÞLEM TARÝHÝ"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24510465
         CurrentDate     =   38388
      End
      Begin VB.TextBox txtusure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "KONTORÜN KARÞILIK GELDÝÐÝ SURE"
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdosifre 
         Appearance      =   0  'Flat
         BackColor       =   &H006CFBD3&
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         MaskColor       =   &H00C0E0FF&
         MouseIcon       =   "frmuye.frx":1E0F
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "OTOMATÝK ÞÝFRE VER"
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox txtfiyat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         ToolTipText     =   "SAAT ÜCRETÝNE GÖRE TEKABÜL EDEN FÝYAT (EN SON SATILAN KONTÖR ÝÞLEMÝNDEN)"
         Top             =   3720
         Width           =   2295
      End
      Begin MSComCtl2.UpDown updkontor 
         Height          =   300
         Left            =   2400
         TabIndex        =   35
         Top             =   3360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Max             =   99999999
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtukontor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "KONTOR (DK)"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtuad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   14
         ToolTipText     =   "ÜYEYE VERÝLEN KISA AD (GÝRÝÞLER BU ADLA OLACAKTIR)"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtusifre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   15
         ToolTipText     =   "ÜYEYE VERÝLEN ÞÝFRE"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtadsoyad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "ÜYE ADI SOYADI"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txttel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "ÜYE TELEFONU"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtadres 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         ToolTipText     =   "ÜYE ADRESÝ"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtaciklama 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "ÜYE AÇIKLAMA"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TOPLAM"
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
         TabIndex        =   55
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih"
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
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "FÝYAT"
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
         TabIndex        =   37
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "KONTOR"
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
         TabIndex        =   34
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
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
         TabIndex        =   33
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ÞÝFRE"
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
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ad Soyad"
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
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel"
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
         TabIndex        =   30
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Adres"
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
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Açýklama"
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
         TabIndex        =   28
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.TextBox txtara 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "ARANACAK ÜYE ÝSMÝNÝ GÝRÝNÝZ"
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox lstuye 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   5490
      ItemData        =   "frmuye.frx":2119
      Left            =   120
      List            =   "frmuye.frx":211B
      TabIndex        =   4
      ToolTipText     =   "ÜYE SEÇÝNÝZ"
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton cmdyeni 
      BackColor       =   &H006CFBD3&
      Caption         =   "Yeni"
      Height          =   495
      Left            =   2640
      MaskColor       =   &H8000000E&
      MouseIcon       =   "frmuye.frx":211D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "YENÝ KAYIT"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmddegistir 
      BackColor       =   &H006CFBD3&
      Caption         =   "Düzelt"
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmuye.frx":2427
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "DEÐÝÞTÝR"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdiptal 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ýptal"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmuye.frx":2731
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "ÝPTAL"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdsil 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      Caption         =   "Sil"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmuye.frx":2A3B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "SÝL"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdara 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ara"
      Height          =   300
      Left            =   2160
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmuye.frx":2D45
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "ÜYE ARA"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdkaydet 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kaydet"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MouseIcon       =   "frmuye.frx":304F
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin MSComCtl2.UpDown updkontor2 
      Height          =   300
      Left            =   4680
      TabIndex        =   58
      Top             =   5520
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99999999
      Enabled         =   -1  'True
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "KONTOR"
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
      Left            =   2640
      TabIndex        =   61
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "FÝYAT"
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
      Left            =   2640
      TabIndex        =   60
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tarih"
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
      TabIndex        =   51
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TOPLAM MÝKTAR"
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
      Left            =   480
      TabIndex        =   49
      Top             =   7320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ÜYELER                           ÜYE  BÝLGÝLERÝ"
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
      TabIndex        =   26
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmuye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rstucret As Recordset
Dim rstuye As Recordset
Dim dtkafe As Database

Private Sub cmbay_Click()
On Error Resume Next
With rstuye
    .MoveFirst
    txttopmiktar = 0
    For i = 1 To .RecordCount
        If Val(Mid(!tarih, 4, 2)) = Val(Mid(cmbay.Text, 1, 2)) Then
            txttopmiktar = CDbl(txttopmiktar) + CDbl(!fiyat)
        End If
        .MoveNext
    Next i
End With

End Sub

Private Sub cmdara_Click()
On Error Resume Next
With rstuye
    .MoveFirst
    lstuye.Clear
    
    If txtara <> "" Then
        txtara = UCase(txtara)
        For i = 1 To .RecordCount
            If Left(txtara, Len(txtara)) = Left(!adsoyad, Len(txtara)) Then
                lstuye.AddItem ((lstuye.ListCount + 1) & "-" & !adsoyad)
            End If
            .MoveNext
        Next i
    Else
        For i = 1 To .RecordCount
            lstuye.AddItem ((lstuye.ListCount + 1) & "-" & !adsoyad)
            .MoveNext
        Next i
    End If
End With

txtara.SetFocus
lstuaye.ListIndex = 0
txtara = ""

End Sub

Private Sub cmdcikar_Click()
On Error Resume Next

With rstuye
    If txtfiyat2 <> "" And !KONTOR >= txtfiyat2 Then
    sec = lstuye.ListIndex
        .Edit
        !KONTOR = CDbl(!KONTOR) - CDbl(txtukontor2)
        !topfiyat = CDbl(txttopfiyat) - CDbl(txtfiyat2)
        !tarih = Date
        .Update
    End If
End With

lstuye.ListIndex = sec
lstuye_Click

End Sub

Private Sub cmddegistir_Click()
On Error Resume Next

cmdyeni.Enabled = False
cmddegistir.Enabled = False
cmdsil.Enabled = False
cmdara.Enabled = False
cmdkaydet.Enabled = True
cmdiptal.Enabled = True

fraubilgi.Enabled = True
lstuye.Enabled = False
txtara.Enabled = False
cmdrapor.Enabled = False
txtad.SetFocus

'þifreler gizlenmesi açýlacak
txtuad.PasswordChar = ""
txtusifre.PasswordChar = ""
End Sub

Private Sub cmdekle_Click()
On Error Resume Next

With rstuye
    If txtfiyat2 <> "" Then
    sec = lstuye.ListIndex
        .Edit
        !KONTOR = CDbl(txtukontor2) + CDbl(!KONTOR)
        !topfiyat = CDbl(txttopfiyat) + CDbl(txtfiyat2)
        !tarih = Date
        .Update
    End If
End With

lstuye.ListIndex = sec
lstuye_Click
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next

If txtsifre = rstucret!sifre Then
    frasifre.Visible = False
End If

End Sub

Private Sub cmdiptal_Click()
On Error Resume Next

'***gösterim refresh******************
cmdyeni.Enabled = True
cmddegistir.Enabled = True
cmdsil.Enabled = True
cmdara.Enabled = True
cmdkaydet.Enabled = False
cmdiptal.Enabled = False

fraubilgi.Enabled = False
lstuye.Enabled = True
txtara.Enabled = True
cmdrapor.Enabled = True
cmdyeni.SetFocus

If chkkayit.Value = 1 Then lstuye.ListIndex = 0

chkkayit.Value = 0

'þifreler gizlenmesi açýlacak
txtuad.PasswordChar = "*"
txtusifre.PasswordChar = "*"

End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next

With rstuye
    If txtadsoyad <> "" And txtuad <> "" And txtusifre <> "" And txtukontor <> "" Then
        If chkkayit.Value = 1 Then
            .AddNew
                ![adsoyad] = UCase(txtadsoyad)
                ![tel] = txttel
                ![adres] = txtadres
                ![aciklama] = txtaciklama
                !AD = txtuad
                !sifre = txtusifre
                !KONTOR = txtukontor
                !fiyat = txtfiyat
                !topfiyat = CDbl(txttopfiyat) + CDbl(txtfiyat)
                !tarih = txttarih
            .Update
        Else
            .Edit
                ![adsoyad] = UCase(txtadsoyad)
                ![tel] = txttel
                ![adres] = txtadres
                ![aciklama] = txtaciklama
                !AD = txtuad
                !sifre = txtusifre
                !KONTOR = txtukontor
                !fiyat = txtfiyat
                !topfiyat = CDbl(txttopfiyat) + CDbl(txtfiyat)
                !tarih = txttarih
            .Update
        End If
               
        '***gösterim refresh******************
        lstuye.Clear
        .MoveFirst
        For i = 1 To .RecordCount
            lstuye.AddItem (lstuye.ListCount + 1 & "-" & !adsoyad)
            .MoveNext
        Next i
        
        'geçerli ayýn seçilmesi
        cmbay_Click
        
        cmdyeni.Enabled = True
        cmddegistir.Enabled = True
        cmdsil.Enabled = True
        cmdara.Enabled = True
        cmdkaydet.Enabled = False
        cmdiptal.Enabled = False
        
        fraubilgi.Enabled = False
        lstuye.Enabled = True
        txtara.Enabled = True
        cmdrapor.Enabled = True
        cmdyeni.SetFocus
        
        chkkayit.Value = 0
        
        'þifreler gizlenmesi açýlacak
        txtuad.PasswordChar = "*"
        txtusifre.PasswordChar = "*"
        
    Else
        MsgBox "Eksik bilgi girdiniz !!!", vbCritical
    End If
End With

End Sub

Private Sub cmdosifre_Click()
On Error Resume Next
txtusifre = Val(Timer * (Val(Mid(Time, 7, 2) - Mid(Time, 1, 2)) + 2) \ 2) & Mid(Date, 1, 2) & Mid(Time, 4, 2) & Chr(Val(Mid(Time, 1, 2)) + 64) & Val(Mid(Time, 7, 2) * Timer)

End Sub

Private Sub cmdsil_Click()
On Error Resume Next

With rstuye
    cevap = MsgBox("Kaydý silmek istiyor musunuz?", vbCritical + vbYesNo)
    If cevap = vbYes Then
        .Delete
        lstuye.Clear
        .MoveFirst
        For i = 1 To .RecordCount
            lstuye.AddItem (lstuye.ListCount + 1 & "-" & !adsoyad)
            .MoveNext
        Next i
        
        'geçerli ayýn seçilmesi
        cmbay_Click
        
        '***temizlik
        txtadsoyad = ""
        txttel = ""
        txtadres = ""
        txtaciklama = ""
        txtuad = ""
        txtusifre = ""
        txtukontor = 0
        txtfiyat = ""
        txttopfiyat = 0
        .MoveFirst
        txtadsoyad = !adsoyad
        txttel = !tel
        txtadres = !adres
        txtaciklama = !aciklama
        txtuad = !AD
        txtusifre = !sifre
        txtukontor = !KONTOR
        txtfiyat = !fiyat
        txttarih = !tarih
    End If
End With

End Sub

Private Sub cmdsorgula_Click()
On Error Resume Next
With rstuye
    If txtsuad <> "" And txtsusifre <> "" Then
        .MoveFirst
        For i = 1 To .RecordCount
            If txtsuad = !AD And txtsusifre = !sifre Then
                lblsudurum = "VAR"
                txtsukontor = !KONTOR
                Exit For
            Else
                lblsudurum = "YOK"
                txtsukontor = 0
            End If
            .MoveNext
        Next i
    Else
        lblsudurum = "Eksik Giriþ"
    End If
    lblsudurum.ForeColor = vbRed
End With

End Sub

Private Sub cmdyeni_Click()
On Error Resume Next
'***
cmdyeni.Enabled = False
cmddegistir.Enabled = False
cmdsil.Enabled = False
cmdara.Enabled = False
cmdkaydet.Enabled = True
cmdiptal.Enabled = True

fraubilgi.Enabled = True
lstuye.Enabled = False
txtara.Enabled = False
cmdrapor.Enabled = False

txtadsoyad.SetFocus

chkkayit.Value = 1

'***temizlik
txtadsoyad = ""
txttel = ""
txtadres = ""
txtaciklama = ""
txtuad = ""
txtusifre = ""
txtukontor = 0
txtfiyat = 0
txttopfiyat = 0
txttarih = Date

'þifreler gizlenmesi açýlacak
txtuad.PasswordChar = ""
txtusifre.PasswordChar = ""

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
On Error Resume Next
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")
Set rstuye = dtkafe.OpenRecordset("uyeler")

rstucret.MoveFirst

'listeye üyelerin listelenmesi
With rstuye
    .MoveFirst
    For i = 1 To .RecordCount
        lstuye.AddItem (lstuye.ListCount + 1 & "-" & !adsoyad)
        .MoveNext
    Next i
End With

'geçerli ayýn seçilmesi
cmbay = cmbay.List(Val(Mid(Date, 3, 2)) + 1)
cmbay_Click

updkontor.Value = txtukontor
lstuye.ListIndex = 0

'þifre sorulmasý
If rstucret!konuye = 1 Then
    frasifre.Move 0, 0
    frasifre.Visible = True
End If

RENK_VER

End Sub



Private Sub lstuye_Click()
On Error Resume Next
With rstuye
    .MoveFirst
    For i = 1 To .RecordCount
        If Mid(lstuye.Text, InStr(1, lstuye.Text, "-") + 1) = !adsoyad Then
            txtadsoyad = !adsoyad
            txttel = !tel
            txtadres = !adres
            txtaciklama = !aciklama
            txtuad = !AD
            txtusifre = !sifre
            txtukontor = !KONTOR
            txtfiyat = !fiyat
            txttopfiyat = !topfiyat
            txttarih = !tarih
            Exit For
        End If
        .MoveNext
    Next i
End With


End Sub

Private Sub txtadsoyad_LostFocus()
    On Error Resume Next
    txtadsoyad = UCase(txtadsoyad)
End Sub

Private Sub txtara_Change()
On Error Resume Next
If KeyAscii = 13 Then cmdara.Value = True
End Sub



Private Sub txtfiyat2_Change()
On Error Resume Next
If txtfiyat = "" Then txtfiyat = 0
End Sub

Private Sub txtsuad_Click()
txtsuad = ""
lblsudurum = "Durum"
lblsudurum.ForeColor = Label12.ForeColor
End Sub

Private Sub txtsusifre_Click()
txtsusifre = ""
lblsudurum = "Durum"
lblsudurum.ForeColor = Label12.ForeColor
End Sub

Private Sub txttopmiktar_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txttopmiktar = Format(txttopmiktar, "#00,0")
Else
    txttopmiktar = Format(txttopmiktar, "#0.00")
End If
End Sub



Private Sub txtukontor_Change()
On Error Resume Next
If txtukontor = "" Then txtukontor = 0

updkontor.Value = txtukontor

If fraubilgi.Enabled = True Then
    If rstucret!parabirimi = 0 Then
        txtfiyat = Format(CDbl(rstucret!ucret / 60) * txtukontor, "#00,0")
    Else
        txtfiyat = Format(CDbl((rstucret!ucret / 100) / 60) * txtukontor, "#0.00")
    End If
End If

If txtukontor / 60 > 0 Then
    If (txtukontor \ 60) < 10 Then
        HH = "0" & (txtukontor \ 60)
    Else
        HH = (txtukontor \ 60)
    End If
    
    If txtukontor - ((txtukontor \ 60) * 60) < 10 Then
        MM = "0" & txtukontor - ((txtukontor \ 60) * 60)
    Else
        MM = txtukontor - ((txtukontor \ 60) * 60)
    End If
    
    txtusure = HH & ":" & MM
Else
    If txtukontor - ((txtukontor \ 60) * 60) < 10 Then
        MM = "0" & txtukontor - ((txtukontor \ 60) * 60)
    Else
        MM = txtukontor - ((txtukontor \ 60) * 60)
    End If
    
    txtusure = "00" & ":" & MM
End If


End Sub

Private Sub txtfiyat_LostFocus()
On Error Resume Next
If txtfiyat = "" Then txtfiyat = 0
cmdkaydet.SetFocus
End Sub


Private Sub txtukontor2_Change()
On Error Resume Next
If txtukontor2 = "" Then txtukontor2 = 0

updkontor2.Value = txtukontor2

    If rstucret!parabirimi = 0 Then
        txtfiyat2 = Format(CDbl(rstucret!ucret / 60) * txtukontor2, "#00,0")
    Else
        txtfiyat2 = Format(CDbl((rstucret!ucret / 100) / 60) * txtukontor2, "#0.00")
    End If
    
If txtukontor2 / 60 > 0 Then
    If (txtukontor2 \ 60) < 10 Then
        HH = "0" & (txtukontor2 \ 60)
    Else
        HH = (txtukontor2 \ 60)
    End If
    
    If txtukontor2 - ((txtukontor2 \ 60) * 60) < 10 Then
        MM = "0" & txtukontor2 - ((txtukontor2 \ 60) * 60)
    Else
        MM = txtukontor2 - ((txtukontor2 \ 60) * 60)
    End If
    
    txtusure2 = HH & ":" & MM
Else
    If txtukontor - ((txtukontor2 \ 60) * 60) < 10 Then
        MM = "0" & txtukontor2 - ((txtukontor2 \ 60) * 60)
    Else
        MM = txtukontor2 - ((txtukontor2 \ 60) * 60)
    End If
    
    txtusure2 = "00" & ":" & MM
End If

End Sub

Private Sub txtusifre_Change()
On Error Resume Next
txtusifre2 = ""
For i = 1 To Len(txtusifre)
    j = i + 65
    harf = Asc(Mid(txtusifre, i, 1)) + 3
    If harf < 100 Then harf = "0" & harf
    txtusifre2 = txtusifre2 & Chr(j) & harf
Next i
txtusifre2 = txtusifre2 & "E" & Len(txtusifre)
End Sub

Private Sub updkontor_Change()
On Error Resume Next
txtukontor = updkontor.Value
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
    If TypeOf C Is Frame Then C.ForeColor = !onrenk
Next
End With
'****************************************************************************
Label26.ForeColor = vbWhite
End Sub

Private Sub updkontor2_Change()
On Error Resume Next
txtukontor2 = updkontor2.Value
End Sub
