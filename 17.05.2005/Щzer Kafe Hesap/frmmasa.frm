VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmasa 
   Appearance      =   0  'Flat
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Masa::."
   ClientHeight    =   8295
   ClientLeft      =   8025
   ClientTop       =   1440
   ClientWidth     =   3345
   ControlBox      =   0   'False
   Icon            =   "frmmasa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   3345
   Begin VB.CheckBox chkuye 
      Caption         =   "uye"
      Height          =   255
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtducret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtparabirimi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   2520
      Locked          =   -1  'True
      MouseIcon       =   "frmmasa.frx":144A
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   1080
      Width           =   735
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            TextSave        =   "10:56"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frasifre 
      BackColor       =   &H00BFA3C9&
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdgiris 
         BackColor       =   &H006CFBD3&
         Caption         =   "*GÝRÝÞ* )>)>)>"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtsifre 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   360
         Width           =   1815
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
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
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
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
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
         TabIndex        =   15
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame framusteri 
      BackColor       =   &H00404080&
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdmusteri 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   3000
         MouseIcon       =   "frmmasa.frx":1754
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "MÜÞTERÝNÝN ÞUANKÝ TOPLAM BORCU "
         Top             =   3240
         Width           =   255
      End
      Begin VB.TextBox txtad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F9FF&
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
         Height          =   375
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "SEÇÝLEN MÜÞTERÝ"
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtara 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "ARANACAK MÜÞTERÝ ÝSMÝNÝ GÝRÝNÝZ"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdara 
         BackColor       =   &H006CFBD3&
         Caption         =   "Ara"
         Height          =   300
         Left            =   2880
         MaskColor       =   &H00C0E0FF&
         MouseIcon       =   "frmmasa.frx":1A76
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "MÜÞTERÝ ARA"
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox lstmusteri 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F9FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         ItemData        =   "frmmasa.frx":1D80
         Left            =   120
         List            =   "frmmasa.frx":1D82
         TabIndex        =   4
         ToolTipText     =   "MÜÞTERÝ SEÇÝNÝZ"
         Top             =   600
         Width           =   3135
      End
      Begin VB.CommandButton cmdtamam 
         BackColor       =   &H006CFBD3&
         Caption         =   "TAMAM"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmmasa.frx":1D84
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "BORC EKLE"
         Top             =   3240
         Width           =   735
      End
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   6
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   70
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   5
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   69
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   4
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   68
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   3
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   67
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   66
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   1
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   65
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   7
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   64
      ToolTipText     =   "DEÐÝÞTÝRMEK ÝÇÝN ÇÝFT TIKLAYINIZ"
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   1
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "00"
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "00"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   3
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "00"
      Top             =   5160
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   4
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "00"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   5
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "00"
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   6
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "00"
      Top             =   6240
      Width           =   375
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   7
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "00"
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox txteu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   840
      TabIndex        =   56
      Text            =   "0"
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox txtadet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   8
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "1"
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "Ek Ucret"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "Ek Ucret"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "Ek Ucret"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "Ek Ucret"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "Ek Ucret"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "Ek Ucret"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox chkeu 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "Ek Ucret"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdeudegistir 
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
      Left            =   120
      MouseIcon       =   "frmmasa.frx":208E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "EK UCRETLERÝ DEÐÝÞTÝR"
      Top             =   7320
      Width           =   3135
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
      Height          =   285
      Left            =   3000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "EK UCRET ÇIKAR"
      Top             =   6960
      Width           =   255
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
      Height          =   285
      Left            =   2640
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "EK UCRET EKLE"
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox txtaciklama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txttopucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtacilismm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmdcikis 
      BackColor       =   &H006CFBD3&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MouseIcon       =   "frmmasa.frx":2398
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "KAYDET KAPAT"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdeucret 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ek Ücretler A"
      Height          =   495
      Left            =   2160
      MouseIcon       =   "frmmasa.frx":26A2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "EK üCRETLER"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txthh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "00"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdveresiyekapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "V.Kapat"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "frmmasa.frx":29AC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "VERESÝYE HESAP KAPAT"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox txtmasanot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   270
      Left            =   600
      TabIndex        =   30
      Top             =   2280
      Width           =   2655
   End
   Begin VB.OptionButton chksucret 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404080&
      Caption         =   "Sýn.Ücret"
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
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton chkssure 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404080&
      Caption         =   "Sýn.Saat"
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
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtekucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmduzatsure 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   ">"
      Height          =   300
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "SINIR SÜRE UZAT"
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmduzat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   ">"
      Height          =   300
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "SINIR ÜCRET UZAT"
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdhesapac 
      BackColor       =   &H006CFBD3&
      Caption         =   "Hesap Aç"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmmasa.frx":2CB6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "HESAP AÇ"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtsure 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
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
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdhesapkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "N.Kapat"
      Height          =   375
      Left            =   1080
      MouseIcon       =   "frmmasa.frx":2FC0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "NAKÝT HESAP KAPAT"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtsucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1560
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox chkrapkaydet 
      BackColor       =   &H00404080&
      Caption         =   "Rapora Kaydet"
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
      TabIndex        =   19
      ToolTipText     =   "RAPORA KAYDET"
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txtacilishh 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkucret2 
      BackColor       =   &H00404080&
      Caption         =   "Alternatif Ücret"
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
      TabIndex        =   10
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtmm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "00"
      Top             =   1560
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   0
   End
   Begin MSComCtl2.UpDown updsucret 
      Height          =   300
      Left            =   3000
      TabIndex        =   34
      ToolTipText     =   "SINIR ÜCRET BELÝRLE"
      Top             =   1920
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   10000
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updhh 
      Height          =   300
      Left            =   2400
      TabIndex        =   35
      ToolTipText     =   "SINIR SURE BELÝRLE"
      Top             =   1560
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   24
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updmm 
      Height          =   300
      Left            =   3000
      TabIndex        =   36
      ToolTipText     =   "SINIR DAKÝKA BELÝRLE"
      Top             =   1560
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   60
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updacilishh 
      Height          =   300
      Left            =   360
      TabIndex        =   37
      ToolTipText     =   "ACILIS SAATÝNÝ DEÐÝÞTÝR"
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   24
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updacilismm 
      Height          =   300
      Left            =   840
      TabIndex        =   38
      ToolTipText     =   "ACÝLÝS DAKÝKASINI DEÐÝÞTÝR"
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   60
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtacilis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   39
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   7
      Left            =   1560
      TabIndex        =   71
      Top             =   6600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   6
      Left            =   1560
      TabIndex        =   72
      Top             =   6240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   5
      Left            =   1560
      TabIndex        =   73
      Top             =   5880
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   4
      Left            =   1560
      TabIndex        =   74
      Top             =   5520
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   75
      Top             =   5160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   76
      Top             =   4800
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown updadet 
      Height          =   300
      Index           =   1
      Left            =   1560
      TabIndex        =   77
      Top             =   4440
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   99
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Miktar"
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
      TabIndex        =   78
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  Açýlýþ          Süre         Ücret"
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
      TabIndex        =   42
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Not"
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ek Ücret"
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
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   1815
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuhesapac 
         Caption         =   "Hesap Aç"
      End
      Begin VB.Menu mnunkapat 
         Caption         =   "Nakit Kapat"
      End
      Begin VB.Menu mnuvkapat 
         Caption         =   "Veresiye Kapat"
      End
      Begin VB.Menu mnuekucret 
         Caption         =   "Ek Ucret Ekle"
      End
   End
   Begin VB.Menu mnudk 
      Caption         =   "mnudk"
      Visible         =   0   'False
      Begin VB.Menu mnu15 
         Caption         =   "15"
      End
      Begin VB.Menu mnu30 
         Caption         =   "30"
      End
      Begin VB.Menu mnu45 
         Caption         =   "45"
      End
   End
End
Attribute VB_Name = "frmmasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'ses apisi
Dim txtsayi
'***
Dim basla As Date
Dim baslangic As Date
Dim bitis As Date
Dim sure As Currency
Dim ucret As Currency
'***
Dim rstrapor As Recordset
Dim rstmusteri As Recordset
Dim rstmrapor As Recordset
Dim rstmasa As Recordset
Dim rstucret As Recordset
Dim dtkafe As Database



Private Sub chkrapkaydet_Click()
On Error Resume Next
'-----rapor kontrol------
If chkrapkaydet.Value = 0 Then
    '---
    If rstucret!konrapor = 1 Then
        '---
        If frasifre.Visible = False Then
            chkrapkaydet.Value = 1
        End If
        '---
        frasifre.Visible = True
        chkrapkaydet.Enabled = False
        txtsifre = ""
        '---
    End If
    '---
End If
'******************
End Sub



Private Sub chksucret_Click()
On Error Resume Next
If txtsucret <> "" Then
    '---------------------------
    If chksucret.Value = 1 Then
    'sýnýrlanan ucreti dakikaya çevir
        topmm = txtsucret \ (rstucret!ucret \ 60)
        '***
        For i = 1 To 24
        '---
            If (topmm) >= 60 Then
                tophh = tophh + 1
                topmm = topmm - 60
            End If
        '---
        Next i
        '********************
        If topmm < 10 Then topmm = "0" & topmm
        '---
        If tophh < 10 Then tophh = "0" & tophh
        '---
        If tophh = 0 Then tophh = "00"
        '---
        If topmm = 0 Then topmm = "00"
        '************************
        txthh = tophh
        txtmm = topmm
        '***************************
    End If
    '--------
Else
    chksucret.Value = 0
    MsgBox "Sýnýrlama Ücreti Girin !", vbInformation
End If
'---
End Sub

Private Sub chkucret2_Click()
On Error Resume Next

With rstmasa
    .Edit
    If chkucret2.Value = 1 Then
        !secucret2 = "1"
    Else
        !secucret2 = "0"
    End If
    .Update
End With

End Sub

Private Sub cmdara_Click()
On Error Resume Next
'***
cevap = txtara
If cevap <> "" Then
     With rstmusteri
        rstmusteri.MoveFirst
        lstmusteri.Clear
        For i = 1 To rstmusteri.RecordCount
            txtara = UCase(txtara)
            If Left(txtara, Len(txtara)) = Left(rstmusteri!AD, Len(txtara)) Then
                AA = i & " - " & ![AD]
                lstmusteri.AddItem (AA)
                lstmusteri.ListIndex = 0
            End If
            rstmusteri.MoveNext
        Next i
    End With
Else
    With rstmusteri
        rstmusteri.MoveFirst
        lstmusteri.Clear
        For i = 1 To rstmusteri.RecordCount
            AA = i & " - " & ![AD]
            lstmusteri.AddItem (AA)
            lstmusteri.ListIndex = 0
            rstmusteri.MoveNext
        Next i
    End With
End If
'***
txtara.SetFocus
SendKeys "{HOME}+{END}"
End Sub

Private Sub cmdcikar_Click()
On Error Resume Next
If CDbl(txteu(8)) <= CDbl(txtekucret) Then
    txtekucret = CDbl(txtekucret) - CDbl(txteu(8))
Else
    MsgBox "Geçersiz deðer !!!", vbInformation
End If
End Sub

Private Sub cmdcikis_Click()
On Error Resume Next
'****
i = txtsayi
frmana.a(i) = txtacilis
frmana.s(i) = txtsure

If txttopucret <> "" Then
    frmana.u(i) = CDbl(txtucret) + CDbl(txtekucret)
Else
    frmana.u(i) = txttopucret
End If

If txtucret <> "" And txtsure <> "" Then
    frmana.a(i) = txtacilishh & ":" & txtacilismm
End If

'****Deðiþiklikler kaydediliyor
If txtucret <> "" Then
    rstmasa.Index = "indexkod"
    rstmasa.Seek "=", txtsayi
    '***
    rstmasa.Edit
    'masaya not eklemek için
    rstmasa![masanot] = txtmasanot
    'ek ucret eklemek için
    rstmasa!eucret = CDbl(txtekucret)
    rstmasa.Update
End If
'***************************


'------sure uzatmak için--------
If txtucret <> "" And chkssure = True Then cmduzatsure_Click
'-------------------------------

'------ucret uzatmak için--------
If txttopucret <> "" And chksucret = True Then cmduzat_Click
'-------------------------------

'aciklamaya kayýt yapýlýyor
rstmasa.Index = "indexkod"
rstmasa.Seek "=", frmana.txtno
'***
rstmasa.Edit
rstmasa![aciklama] = txtadet(1) & "," & txtadet(2) & "," & txtadet(3) & "," & txtadet(4) & "," & txtadet(5) & "," & txtadet(6) & "," & txtadet(7) & "," & txteu(8)
rstmasa!uye = chkuye.Value
rstmasa.Update

Unload frmmasa

End Sub

Private Sub cmdekle_Click()
On Error Resume Next
txtekucret = CDbl(txtekucret) + CDbl(txteu(8))

End Sub

Private Sub cmdeucret_Click()
If txtacilis <> "" Then
    If cmdeucret.Caption = "Ek Ücretler A" Then
        frmmasa.Height = 8655
        framusteri.Visible = False
        cmdeucret.Caption = "Ek Ücretler K"
        '---------------------------------------
        'MADE BY MEHMET ALTINEL
        '---------------------------------------
    Else
        frmmasa.Height = 5055
        cmdeucret.Caption = "Ek Ücretler A"
    End If
Else
    MsgBox "Önce hesap açmalýsýnýz !!!", vbInformation
End If
End Sub

Private Sub cmdeudegistir_Click()
On Error Resume Next
'--------------------
If cmdeudegistir.Caption = "Deðiþtir" Then
    cmdeudegistir.Caption = "Kaydet"
    cmdeukaydet.Visible = True
    
    For i = 1 To 7
        chkeu(i).Locked = False
        txteu(i).Locked = False
    Next i
    
    chkeu(1).SetFocus
    SendKeys "{HOME}" + "{END}"
    '******
    fraengel.Enabled = False
    '***
Else
    cmdeudegistir.Caption = "Deðiþtir"
    
    rstucret.Edit
    rstucret![euisim1] = chkeu(1)
    rstucret![euisim2] = chkeu(2)
    rstucret![euisim3] = chkeu(3)
    rstucret![euisim4] = chkeu(4)
    rstucret![euisim5] = chkeu(5)
    rstucret![euisim6] = chkeu(6)
    rstucret![euisim7] = chkeu(7)
    rstucret.Update
    
    rstucret.Edit
    rstucret![eu1] = txteu(1)
    rstucret![eu2] = txteu(2)
    rstucret![eu3] = txteu(3)
    rstucret![eu4] = txteu(4)
    rstucret![eu5] = txteu(5)
    rstucret![eu6] = txteu(6)
    rstucret![eu7] = txteu(7)
    rstucret.Update
    
    For i = 1 To 7
        txteu(i).Locked = True
        chkeu(i).Locked = True
    Next i
    fraengel.Enabled = True
End If
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
'----
If txtsifre = rstucret!sifre Then
    chkrapkaydet.Value = 0
    frasifre.Visible = False
    chkrapkaydet.Enabled = True
Else
    MsgBox "Yanlýþ Þifre Girdiniz!!!", vbCritical
End If
End Sub


Private Sub cmdhesapac_Click()
On Error Resume Next
'kontrol
If txtsure = "" Then
    '***
    Timer1.Enabled = True
    Timer1.Interval = 1
    '***
    txtacilis = txtacilishh & ":" & txtacilismm
    '***
    i = txtsayi
    frmana.a(i) = txtacilis
    frmana.s(i) = txtsure
    frmana.u(i) = txtucret
        
    '****masaya not eklemek için
    If txtmasanot <> "" Then
        '***
        rstmasa.Index = "indexkod"
        rstmasa.Seek "=", txtsayi
        '***
        rstmasa.Edit
        rstmasa![masanot] = txtmasanot
        rstmasa.Update
    End If
    
    '***********ucret sýnýrlamasý için*******
        If chksucret.Value = True Then
            '***
            rstmasa.Index = "indexkod"
            rstmasa.Seek "=", txtsayi
            '***
            rstmasa.Edit
            rstmasa![sucret] = txtsucret
            rstmasa.Update
        Else
            '***
            rstmasa.Index = "indexkod"
            rstmasa.Seek "=", Val(txtsayi)
            '***
            rstmasa.Edit
            rstmasa![notsucret] = "0"
            rstmasa.Update
        End If

    '********sure sýnýrlamasý için***********
        If chkssure.Value = True Then
            '***
            rstmasa.Index = "indexkod"
            rstmasa.Seek "=", txtsayi
            '***
            rstmasa.Edit
            rstmasa![ssure] = txthh & ":" & txtmm
            rstmasa.Update
        Else
            '***
            rstmasa.Index = "indexkod"
            rstmasa.Seek "=", Val(txtsayi)
            '***
            rstmasa.Edit
            rstmasa![notssure] = "0"
            rstmasa.Update
        End If
    '************************************************
    '*******masa aktifleþtirmek (hesaba göre) için***
    If rstucret!hesapac = 1 Then
        With frmclient
            If rstucret!client = 1 Then
                .chkotodurum.Value = 0
                .optm(txtsayi).Value = True
                .chkotodurum.Value = 1
                .cmdkulac2.Value = True
            End If
        End With
    End If
    '*************************************************
Else
    MsgBox "Öncelikle Hesabý Kapatmalýsýnýz!!!", vbCritical
End If
'***
End Sub

Private Sub cmdhesapkapat_Click()
On Error Resume Next

If txtucret <> "" Then
    KSINYAL 'kasa sesi
End If

'***
Timer1.Enabled = False
Timer1.Interval = 0
'***

'uyeler için
If chkuye.Value = 0 Then
'raporlama için kayýt
    If chkrapkaydet.Value = 1 And txtacilis <> "" And txtucret <> "" And txtsure <> "" Then
        rstrapor.MoveFirst
        '-----------------------------------------
        rstrapor.AddNew
        rstrapor![acilis] = txtacilis
        rstrapor![sure] = txtsure
        rstrapor![ucret] = CDbl(txtucret) + CDbl(txtekucret)
        rstrapor![tarih] = Date
        rstrapor![aciklama] = txtaciklama & " (" & txtekucret & ")"
        rstrapor![bitis] = Format(Time, "hh:mm")
        rstrapor![masa] = Mid(Me.Caption, 4, Len(Me.Caption) - 6)
        rstrapor.Update
        '------------------------------------
    End If
End If

'***
txtacilis = ""
txtsure = ""
txtucret = ""
txttopucret = ""
txtmasanot = ""
txthh = "00"
txtmm = "00"
txtsucret = ""
txtekucret = ""
txtmasanot = ""
txtaciklama = ""
'***
'ek ucretleri sýfýrlama
For i = 1 To 7
    updadet(i).Value = 0
Next i

rstmasa.Index = "indexkod"
rstmasa.Seek "=", txtsayi
rstmasa.Edit
rstmasa![sucret] = ""
rstmasa![ssure] = ""
rstmasa![masanot] = ""
rstmasa![notsucret] = "0"
rstmasa![notssure] = "0"
rstmasa![eucret] = "0"
rstmasa![secucret2] = "0"
rstmasa![aciklama] = ""
chkuye.Value = 0
rstmasa!uye = chkuye.Value
rstmasa.Update
'***
i = txtsayi
frmana.a(i).ForeColor = vbBlack
frmana.s(i).ForeColor = vbBlack
frmana.u(i).ForeColor = vbBlack
'---
frmana.a(i).BackColor = vbWhite
frmana.s(i).BackColor = vbWhite
frmana.u(i).BackColor = vbWhite
'---

    '*******masa aktifleþtirmek (hesaba göre) için***
    If rstucret!hesapac = 1 Then
        With frmclient
            If rstucret!client = 1 Then
                .chkotodurum.Value = 0
                .optm(txtsayi).Value = True
                .chkotodurum.Value = 1
                
                .cmdkulkapat2.Value = True
                
            End If
        End With
    End If
'uyelerin resimlerinin yokolmasý
frmana.cmdmasa(txtsayi).Picture = frmana.cmdcikis.Picture
frmana.cmdmasa(txtsayi).Caption = rstmasa!masaad

End Sub

Private Sub cmdmusteri_Click()
On Error Resume Next
With rstmusteri
    .Index = "indexad"
    .Seek "=", txtad
    If .NoMatch = False Then
        MsgBox ![AD] & " isismli müþterinin þu anki borcu " & ![borc]
    End If
End With
End Sub

Private Sub cmdtamam_Click()
On Error Resume Next
'***
frmmasa.Height = 5055
framusteri.Visible = False
'***
With rstmusteri
    '---
    .Index = "indexad"
    .Seek "=", txtad
    '---
    If .NoMatch = False Then
        .Edit
        If rstucret!parabirimi = 0 Then
        ![borc] = Val(CLng(!borc)) + Val(CLng(txtucret))
            Else
       ![borc] = CDbl(!borc) + CDbl(txtucret)
        End If
        .Update
        
        '***raporlama için******
        With rstmrapor
            .AddNew
            !AD = txtad
            !tarih = Date
            !islem = "Borc...."
            !miktar = txtucret
            .Update
        End With

        MsgBox txtad & " isimli müþteriye " & txtucret & " Borc eklenmiþtir", vbInformation
    Else
        MsgBox "Müþteri bulunamadý !!!", vbInformation
        Exit Sub
    End If
    '---
End With
'***********************
cmdhesapkapat_Click
'****
End Sub

Private Sub cmduzat_Click()
On Error Resume Next
chksucret.Value = True
'---
If chksucret.Value = True Then
    '***
    rstmasa.Index = "indexkod"
    rstmasa.Seek "=", txtsayi
    '***
    
    rstmasa.Edit
    If rstucret!parabirimi = 0 Then
        rstmasa![sucret] = Val(CLng(txtsucret))
    Else
        rstmasa![sucret] = CDbl(txtsucret)
    End If
    
    rstmasa![notsucret] = "0"
    rstmasa![notssure] = "0"
    rstmasa![ssure] = ""
    txthh = "00"
    txtmm = "00"
    rstmasa.Update
    
    '*********sýnýrlanan masaya renk deðiþimi formun loadýnda************
    i = txtsayi
    With frmana
        .a(i).BackColor = vbWhite
        .s(i).BackColor = vbWhite
        .u(i).BackColor = vbWhite
        '---
        .a(i).ForeColor = vbBlack
        .s(i).ForeColor = vbBlack
        .u(i).ForeColor = vbBlack
    End With
    '*******************************************************************
    
    'sýnýrlanan hesap dolunca cliente gönderme
    If rstucret!otokapat = 1 Then
        With frmclient
        If rstucret!client = 1 Then
                .Timer2.Enabled = False
                If .chkotodurum.Value = 1 Then
                    .chkotodurum.Value = 0
                    .optm(i).Value = True
                    .cmdkulac2.Value = True
                    .chkotodurum.Value = 1
                Else
                    .optm(i).Value = True
                    .cmdkulac2.Value = True
                End If
                .Timer2.Enabled = True
        End If
        End With
    End If


End If

End Sub

Private Sub cmduzatsure_Click()
On Error Resume Next
chkssure.Value = True
'---
If chkssure.Value = True Then
    '***
    rstmasa.Index = "indexkod"
    rstmasa.Seek "=", txtsayi
    '***
    rstmasa.Edit
    rstmasa![ssure] = txthh & ":" & txtmm
    rstmasa![notssure] = "0"
    
    rstmasa![notsucret] = "0"
    rstmasa![sucret] = ""
    txtsucret = "0"
    rstmasa.Update
   

    '*********sýnýrlanan masaya renk deðiþimi formun loadýnda************
    i = txtsayi
    With frmana
        .a(i).BackColor = vbWhite
        .s(i).BackColor = vbWhite
        .u(i).BackColor = vbWhite
        '---
        .a(i).ForeColor = vbBlack
        .s(i).ForeColor = vbBlack
        .u(i).ForeColor = vbBlack
    End With
    '*******************************************************************
    
    'sýnýrlanan hesap dolunca cliente gönderme
    If rstucret!otokapat = 1 Then
        With frmclient
        If rstucret!client = 1 Then
                .Timer2.Enabled = False
                If .chkotodurum.Value = 1 Then
                    .chkotodurum.Value = 0
                    .optm(i).Value = True
                    .cmdkulac2.Value = True
                    .chkotodurum.Value = 1
                Else
                    .optm(i).Value = True
                    .cmdkulac2.Value = True
                End If
                .Timer2.Enabled = True
        End If
        End With
    End If
    
    'MsgBox "Masa " & txtsayi & " için hesap " & rstmasa!ssure & " 'ye  uzatýlmýþtýr"
    
End If
'---
End Sub

Private Sub cmdveresiyekapat_Click()
On Error Resume Next
'****
If txtacilis <> "" And txtucret <> "" And txtsure <> "" Then
    If txtucret <> 0 Then
        frmmasa.Height = 8655
        framusteri.Visible = True
        
        '***listeleniyor*****
        With rstmusteri
            Do Until .EOF
                lstmusteri.AddItem (lstmusteri.ListCount + 1 & " - " & ![AD])
                .MoveNext
            Loop
        End With
        
    Else
        MsgBox "Ücret deðeri geçersiz !!!", vbInformation
    End If
Else
    MsgBox "Önce hesap açmalýsýnýz !!!", vbInformation
End If
'***
txtara.SetFocus
End Sub

Private Sub cmdx_Click()
frasifre.Visible = False
chkrapkaydet.Enabled = True
chkrapkaydet.Value = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
On Error Resume Next
'-----görünüm------
Me.Top = frmana.Top
Me.Left = frmana.Left
'---------------------

'***
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")
Set rstmasa = dtkafe.OpenRecordset("masalar")
Set rstrapor = dtkafe.OpenRecordset("raporlar")
Set rstmusteri = dtkafe.OpenRecordset("musteriler")
Set rstmrapor = dtkafe.OpenRecordset("mrapor")
'---------------------

rstucret.MoveFirst
'*****Ek Ücretler*****
chkeu(1) = rstucret![euisim1]
chkeu(2) = rstucret![euisim2]
chkeu(3) = rstucret![euisim3]
chkeu(4) = rstucret![euisim4]
chkeu(5) = rstucret![euisim5]
chkeu(6) = rstucret![euisim6]
chkeu(7) = rstucret![euisim7]
'-------------------------
txteu(1) = rstucret![eu1]
txteu(2) = rstucret![eu2]
txteu(3) = rstucret![eu3]
txteu(4) = rstucret![eu4]
txteu(5) = rstucret![eu5]
txteu(6) = rstucret![eu6]
txteu(7) = rstucret![eu7]
'*************************


'---
rstmasa.Index = "indexkod"
rstmasa.Seek "=", frmana.txtno
txtmasanot = rstmasa!masanot
txtsucret = rstmasa![sucret]
chkuye.Value = rstmasa!uye

'ekucretler veriliyor
updadet(1).Value = Val(Mid(rstmasa!aciklama, 1, 2))
updadet(2).Value = Val(Mid(rstmasa!aciklama, 4, 2))
updadet(3).Value = Val(Mid(rstmasa!aciklama, 7, 2))
updadet(4).Value = Val(Mid(rstmasa!aciklama, 10, 2))
updadet(5).Value = Val(Mid(rstmasa!aciklama, 13, 2))
updadet(6).Value = Val(Mid(rstmasa!aciklama, 16, 2))
updadet(7).Value = Val(Mid(rstmasa!aciklama, 19, 2))
txtekucret = ""

'---------------------------------
If rstmasa!secucret2 = "1" Then
    chkucret2.Value = 1
Else
    chkucret2.Value = 0
End If
'----------------------------------

If rstucret!parabirimi = 0 Then
    txtekucret = Format(rstmasa![eucret], "#00,0")
Else
    txtekucret = Format(rstmasa![eucret], "#0.00")
End If

If rstmasa![ssure] <> "" Then
    txthh = Mid(rstmasa![ssure], 1, 2)
    txtmm = Mid(rstmasa![ssure], 4, 2)
End If

'-------------form baþlýðý------------------------------
AD = frmana.cmdmasa(frmana.txtno).Caption
frmmasa.Caption = ".::" & AD & "::."

'*****************************************************
Timer1.Enabled = False
Timer1.Interval = 0
'***
'----
txtsayi = frmana.txtno
i = txtsayi
txtacilis = frmana.a(i)
txtsure = frmana.s(i)
txtucret = frmana.u(i)
'******************
Me.Height = 5055

'********baþlangýç saati**********
If frmana.a(i) = "" Then
    baslangic = Format(Now(), "hh:mm")
    txtacilishh = Mid(baslangic, 1, 2)
    txtacilismm = Mid(baslangic, 4, 2)

    updacilishh.Value = Val(txtacilishh)
    updacilismm.Value = Val(txtacilismm)
Else
    txtacilishh = Mid(txtacilis, 1, 2)
    txtacilismm = Mid(txtacilis, 4, 2)
    
    '****sýnýrlama uzatma****
    If rstmasa!sucret = 1 Then
    txtsucret = CLng(frmana.u(txtsayi))
    updsucret.Value = Val(txtsucret) \ Val(rstucret![birim])
    updhh.Value = Val(txthh)
    updmm.Value = Val(txtmm)
    End If
    '************************
End If

If txtucret <> "" And txtsure <> "" Then
    Timer1.Enabled = True
    Timer1.Interval = 1
End If

'---------------------------
updacilishh.Value = Val(txtacilishh)
updacilismm.Value = Val(txtacilismm)

updhh.Value = Val(txthh)
updmm.Value = Val(txtmm)


If rstucret!parabirimi = 0 Then
    updsucret.Value = txtsucret \ Val(rstucret![birim])
Else
    updsucret.Value = CDbl(txtsucret) / CDbl(CDbl(rstucret![birim]) / 100)
End If

chksucret.Value = False
chkssure.Value = False
'---------------------------

RENK_VER

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
txtaciklama.Height = 495
txtducret.Visible = False
End Sub

Private Sub lstmusteri_Click()
On Error Resume Next
AA = InStr(1, lstmusteri.Text, "-") + 2
txtad = Mid(lstmusteri.Text, AA)
End Sub

Private Sub mnu15_Click()
On Error Resume Next
updmm.Value = 15
End Sub

Private Sub mnu30_Click()
On Error Resume Next
updmm.Value = 30

End Sub

Private Sub mnu45_Click()
On Error Resume Next
updmm.Value = 45
End Sub

Private Sub mnuekucret_Click()
On Error Resume Next
frmana.cmdmasa(frmana.txtno) = Click
Me.Show
cmdeucret_Click
End Sub

Private Sub mnuhesapac_Click()
frmana.cmdmasa(frmana.txtno) = Click
Me.Show
cmdhesapac_Click
cmdcikis_Click
End Sub

Private Sub mnunkapat_Click()
frmana.cmdmasa(frmana.txtno) = Click
Me.Show
If txtmasanot <> "" Then
    MsgBox "NOT:" + vbCrLf & txtmasanot, vbInformation
    cmdhesapkapat_Click
    cmdcikis_Click
Else
    cmdhesapkapat_Click
    cmdcikis_Click
End If
End Sub

Private Sub mnuvkapat_Click()
frmana.cmdmasa(frmana.txtno) = Click
Me.Show
cmdveresiyekapat_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'-----------------------------------------
ahh1 = Val(Mid(txtacilis, 1, 2))
amm = Val(Mid(txtacilis, 4, 2))
'----
bhh1 = Val(Mid(Format(Time, "hh:mm"), 1, 2))
bmm = Val(Mid(Format(Time, "hh:mm"), 4, 2))
'-----------------------------------------

If ahh1 > bhh1 Then
    bhh1 = Val(bhh1) + 24
    shh = Val(bhh1 - ahh1)
    smm = Val(bmm - amm)
    
    If smm < 0 Then
        smm = smm + 60
        shh = Val(shh) - 1
    End If
    
Else
    ahh = ahh1 * 60 + amm
    bhh = bhh1 * 60 + bmm
    shh = (bhh - ahh) \ 60
    smm = (bhh - ahh) - (60 * shh)
End If

'*************************ucret için************************
With rstucret
    '***alternatif ücret seçimi(ucret2)--------
    Dim ucret
    If chkucret2.Value = 1 Then
        ucret = !ucret2
    Else
        ucret = !ucret
    End If
    '-----------------------------------------
    '--deðerleri deðiþkenlere atýyoruz
    Dim lngUCRET, lngBASUCRET, lngBIRIM, lngKAFEUCRET, lngEKUCRET, lngATILANSIFIR, lngKURUS
    lngATILANSIFIR = 1000000
    lngKURUS = 100
    lngUCRET = Val(ucret) * lngATILANSIFIR
    lngBIRIM = Val(!birim) * lngATILANSIFIR
    lngBASUCRET = Val(!basucret) * lngATILANSIFIR
    lngKAFEUCRET = Val(rstmasa!ucret) * lngATILANSIFIR
    lngEKUCRET = Val(rstmasa!eucret) * lngATILANSIFIR
    '-------------------------------------------------------------------

                                   
    If !parabirimi = 0 Then 'parabirimi TL ise
         txtucret = 0
         txtucret = (((((shh * 60) + smm) * (Val(ucret) \ 60)) \ Val(!birim)) * Val(!birim))
            
           If (shh * 60 + smm) * (Val(ucret) \ 60) < Val(!basucret) Then
                txtucret = Val(!basucret)
            Else
                txtucret = ((((shh * 60) + smm) * (Val(ucret) \ 60) \ Val(!birim)) * Val(!birim))
                'yukarý yuvarlama
                If rstucret!yyuvarla = 1 Then
                If ((shh * 60) + smm) - (((shh * 60) + smm) \ (Val(!birim) / 10000)) * (Val(!birim) / 10000) >= 3 Then
                    txtucret = ((((shh * 60) + smm) * (Val(ucret) \ 60) \ Val(!birim)) * Val(!birim)) + Val(!birim)
                End If
                End If
                
            End If
            txtucret = Format(txtucret, "#00,0")
            txtucret = (CDbl(txtucret) \ CDbl(!birim)) * CDbl(!birim)
                                   
            txtucret = Val(CLng(txtucret))
            txtucret = Format(txtucret, "#00,0")
            
    Else 'ytl ise
        txtucret = 0
        If (shh * 60 + smm) * (Val(lngUCRET) \ 60) < Val(lngBASUCRET) Then
            txtucret = lngBASUCRET / lngATILANSIFIR / lngKURUS
        Else
            txtucret = ((((((shh * 60 + smm) * Val((lngUCRET) \ 60)) \ Val(lngBIRIM)) * Val(lngBIRIM)) + Val(lngEKUCRET)) / lngATILANSIFIR) / lngKURUS
            'yukarý yuvarlama
            If rstucret!yyuvarla = 1 Then
            If ((shh * 60) + smm) - (((shh * 60) + smm) \ (Val(lngBIRIM) / 1000000)) * (Val(lngBIRIM) / 1000000) >= 3 Then
               txtucret = (((((((shh * 60 + smm) * Val((lngUCRET) \ 60)) \ Val(lngBIRIM)) * Val(lngBIRIM)) + Val(lngEKUCRET)) + Val(lngBIRIM)) / lngATILANSIFIR) / lngKURUS
            End If
            
            End If
        End If
            txtucret = Format(txtucret, "#0.00")
            txtucret = ((CDbl(txtucret) * 100) \ (CDbl(lngBIRIM) / 1000000)) * (CDbl(lngBIRIM) / 1000000) / 100
            
            txtucret = CDbl(txtucret)
            txtucret = Format(txtucret, "#0.00")
    End If

End With
'*********************************************************
''------------------------------------
If shh < 10 Then shh = "0" & shh
'------------------------------------
If smm < 10 Then smm = "0" & smm
'---------------------------
txtsure = shh & ":" & smm
'---------------------------

'toplam ucret gösterimi
txttopucret = CDbl(txtucret) + CDbl(txtekucret)

End Sub

Private Sub txtaciklama_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If txtaciklama <> "" Then txtaciklama.Height = 975
txtaciklama.ToolTipText = txtaciklama
End Sub

Private Sub txtadet_Change(Index As Integer)
'On Error Resume Next
If txtadet(Index) < 10 Then txtadet(Index) = "0" & Val(txtadet(Index))

'aciklama gösterimi
txtaciklama = ""
For i = 1 To 7
    If txtadet(i) > 0 Then
        If Not rstucret!parabirimi = 1 Then
            If txtaciklama = "" Then
                txtaciklama = Val(txtadet(i)) & "x" & chkeu(i) & "(" & Format(txteu(i) * txtadet(i), "#0.00") & ") "
            Else
                txtaciklama = txtaciklama & ", " & Val(txtadet(i)) & "x" & chkeu(i) & "(" & Format(txteu(i) * txtadet(i), "#0.00") & ") "
            End If
        Else
            If txtaciklama = "" Then
                txtaciklama = Val(txtadet(i)) & "x" & chkeu(i) & "(" & Format(txteu(i) * txtadet(i), "#00,0") & ") "
            Else
                txtaciklama = txtaciklama & ", " & Val(txtadet(i)) & "x" & chkeu(i) & "(" & Format(txteu(i) * txtadet(i), "#00,0") & ") "
            End If
        End If
    End If
Next i

End Sub

Private Sub txtara_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdara_Click
End Sub

Private Sub txtekucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtekucret = Format(txtekucret, "#00,0")
Else
    txtekucret = Format(txtekucret, "#0.00")
End If
End Sub

Private Sub txteu_Change(Index As Integer)
If rstucret!parabirimi = 0 Then
    txteu(Index) = Format(txteu(Index), "#00,0")
    txteu(Index).SelStart = Len(txteu(Index))
End If
End Sub

Private Sub txteuozelmiktar_Change()
If rstucret!parabirimi = 0 Then
    txteuozelmiktar = Format(txteuozelmiktar, "#00,0")
    txteuozelmiktar.SelStart = Len(txteuozelmiktar)
End If
End Sub

Private Sub txteu_GotFocus(Index As Integer)
On Error Resume Next
If Index = 8 Then txteu(8) = ""
End Sub





Private Sub txtparabirimi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If txtparabirimi <> "" Then
    txtducret = ""
    txtducret.Visible = True
    If txtparabirimi = "YTL" Then
        txtducret = CDbl(txttopucret) * 1000000
        txtducret = Format(txtducret, "#00,0") & " TL"
    Else
        txtducret = CDbl(txttopucret) / 1000000
        txtducret = Format(txtducret, "#0.00") & " YTL"
    End If
End If

End Sub

Private Sub txtsifre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgiris_Click
End If
End Sub

Private Sub txtsucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtsucret = Format(txtsucret, "#00,0")
Else
    txtsucret = Format(txtsucret, "#0.00")
End If

End Sub

Private Sub txttopucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then 'parabirimi TL ise
    txttopucret = Format(txttopucret, "#00,0")
    txtparabirimi = "TL"
Else
    txttopucret = Format(txttopucret, "#0.00")
    txtparabirimi = "YTL"
End If

End Sub

Private Sub updacilishh_Change()
On Error Resume Next
If Not rstucret!konasaat = 1 Then
    '---
    If updacilishh.Value = 24 Then updacilishh.Value = 0
    If updacilishh.Value = -1 Then updacilishh.Value = 23
    txtacilishh = updacilishh.Value
    If txtacilishh < 10 Then txtacilishh = "0" & txtacilishh
    '---
End If
End Sub

Private Sub updacilismm_Change()
On Error Resume Next
If Not rstucret!konasaat = 1 Then
    '---
    If updacilismm.Value = 60 Then updacilismm.Value = 0
    If updacilismm.Value = -1 Then updacilismm.Value = 59
    
    txtacilismm = updacilismm.Value
    If txtacilismm < 10 Then txtacilismm = "0" & txtacilismm
    '---
End If
End Sub

Private Sub updadet_Change(Index As Integer)
'On Error Resume Next
txtadet(Index) = updadet(Index).Value

txtekucret = txteu(1) * txtadet(1)
For i = 1 To 6
    txtekucret = txtekucret + txteu(i + 1) * txtadet(i + 1)
Next i

End Sub

Private Sub updhh_Change()
On Error Resume Next
updhh.Max = 24
'***
If updhh.Value = 24 Then updhh.Value = 0
If updhh.Value = -1 Then updhh.Value = 23
    
If updhh.Value < 10 Then
    txthh = "0" & updhh.Value
Else
    txthh = updhh.Value
End If
'***
chkssure.Value = 1
End Sub

Private Sub updmm_Change()
On Error Resume Next
updmm.Max = 60
'***
If updmm.Value = 60 Then updmm.Value = 0
If updmm.Value = -1 Then updmm.Value = 59

If updmm.Value < 10 Then
    txtmm = "0" & updmm.Value
Else
    txtmm = updmm.Value
End If
'***
chkssure.Value = 1
End Sub

Private Sub updmm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
    PopupMenu frmmasa.mnudk
End If
End Sub

Private Sub updsucret_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtsucret = Val(updsucret.Value) * Val(rstucret![birim])
Else
    txtsucret = Format((Val(updsucret.Value) * Val(rstucret![birim])) / 100, "#0.00")
End If
'---
chksucret_Click
'----
chksucret.Value = 1
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
Label5.ForeColor = vbWhite
End Sub
Private Sub KSINYAL()
sndPlaySound (App.Path & "\kasa.wav"), 0
End Sub
