VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmkasa 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: $$$ Kasa $$$ ::."
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmkasa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frasifre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   4800
      TabIndex        =   83
      Top             =   5400
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame2 
         BackColor       =   &H00BFA3C9&
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   1680
         TabIndex        =   84
         Top             =   1680
         Width           =   2535
         Begin VB.TextBox txtsifre 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdgiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   720
            Width           =   1215
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
            TabIndex        =   87
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
            TabIndex        =   86
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.Frame frakasamenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame fraekgelir 
         Appearance      =   0  'Flat
         BackColor       =   &H00BFA3C9&
         Caption         =   ","
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   1200
         TabIndex        =   70
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         Begin VB.CommandButton cmdcikis 
            Appearance      =   0  'Flat
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
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "EK GELÝR BÖLÜMÜNÜ  KAPAT"
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtegaciklama 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            ToolTipText     =   "EK GELÝR HAKKINDA AÇIKLAMA"
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtegmiktar 
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
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Text            =   "0"
            ToolTipText     =   "EK GELÝR MÝKTARI"
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmdegkaydet 
            BackColor       =   &H006CFBD3&
            Caption         =   "K  A  Y  D E  T  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   3000
            MouseIcon       =   "frmkasa.frx":144A
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "KAYDET"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label23 
            BackColor       =   &H00808080&
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "                EK GELÝR RAPORU"
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
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   3495
         End
      End
      Begin VB.Frame frakayit 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   5160
         TabIndex        =   68
         Top             =   720
         Width           =   735
         Begin VB.CommandButton cmdekgelir 
            Appearance      =   0  'Flat
            BackColor       =   &H006CFBD3&
            Caption         =   "Ek Gelir"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            MouseIcon       =   "frmkasa.frx":1754
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   94
            ToolTipText     =   "EK GELÝR BÖLÜMÜ"
            Top             =   1800
            Width           =   735
         End
         Begin VB.CommandButton cmdgsil 
            Appearance      =   0  'Flat
            BackColor       =   &H006CFBD3&
            Caption         =   "Sil"
            Height          =   495
            Left            =   0
            MaskColor       =   &H00C0E0FF&
            MouseIcon       =   "frmkasa.frx":1A5E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "SÝL"
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton cmdgkaydet 
            BackColor       =   &H006CFBD3&
            Caption         =   "Kaydet"
            Enabled         =   0   'False
            Height          =   495
            Left            =   0
            MaskColor       =   &H00C0E0FF&
            MouseIcon       =   "frmkasa.frx":1D68
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "KAYDET"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton cmdgdegistir 
            BackColor       =   &H006CFBD3&
            Caption         =   "Deðiþtir"
            Height          =   495
            Left            =   0
            MaskColor       =   &H00C0E0FF&
            MouseIcon       =   "frmkasa.frx":2072
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "DEÐÝÞTÝR"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         TabIndex        =   61
         Top             =   3360
         Width           =   3135
         Begin VB.TextBox txtveresiye 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   67
            ToolTipText     =   "TÜM VERESÝYELERÝN TOPLAMI"
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtntop 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   66
            ToolTipText     =   "BU AY ÝÇÝN TOPLAM  NAKÝT PARA (ÜYELER+MÜÞTERÝLER+MASALAR+EK GELÝR)"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtgtop 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   65
            ToolTipText     =   "BU AY ÝÇÝN GÝDERLER TOPLAMI"
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label19 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Veresiyeler"
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
            TabIndex        =   64
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label18 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Nakit Top"
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
            TabIndex        =   63
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Gider Top."
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
            TabIndex        =   62
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.TextBox txtnbakiye 
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
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   60
         ToolTipText     =   "NET KAR(VERESÝYELER HARÝÇ)"
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Frame frapanel 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   0
         TabIndex        =   51
         Top             =   360
         Width           =   5895
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   7
            Left            =   4080
            TabIndex        =   32
            Top             =   2520
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   6
            Left            =   4080
            TabIndex        =   29
            Top             =   2160
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   5
            Left            =   4080
            TabIndex        =   26
            Top             =   1800
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   4
            Left            =   4080
            TabIndex        =   23
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   3
            Left            =   4080
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   2
            Left            =   4080
            TabIndex        =   17
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkode 
            BackColor       =   &H00404080&
            Caption         =   "Ödendi"
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
            Index           =   1
            Left            =   4080
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   30
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   960
            TabIndex        =   27
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   960
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   960
            TabIndex        =   21
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   960
            TabIndex        =   18
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   960
            TabIndex        =   15
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   16
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   3
            Left            =   2760
            TabIndex        =   19
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   4
            Left            =   2760
            TabIndex        =   22
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   5
            Left            =   2760
            TabIndex        =   25
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   6
            Left            =   2760
            TabIndex        =   28
            Top             =   2160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Index           =   7
            Left            =   2760
            TabIndex        =   31
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16707296
            Format          =   24641537
            CurrentDate     =   38253
         End
         Begin VB.Label Label21 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Bedeli                      Ödeme tarihi"
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
            Left            =   960
            TabIndex        =   69
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label9 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Kira"
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
            TabIndex        =   58
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Elektrik"
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
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Su"
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
            TabIndex        =   56
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label12 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Ýnternet"
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
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefon"
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
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Eleman"
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
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   "Diðer"
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
            Top             =   2520
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbay 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmkasa.frx":237C
         Left            =   4320
         List            =   "frmkasa.frx":23A4
         TabIndex        =   11
         Text            =   "Ay Seçiniz"
         ToolTipText     =   "ÝÞLEM YAPILACAK AY'I SEÇÝNÝZ"
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblyorum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yorum"
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
         Height          =   255
         Left            =   3360
         TabIndex        =   82
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nakit Bakiye"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H006CFBD3&
         Height          =   255
         Left            =   3360
         TabIndex        =   59
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "AYLIK KASA ÝÞLEMLERÝ"
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
         TabIndex        =   50
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdkaydet 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kaydet"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      MouseIcon       =   "frmkasa.frx":2423
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdara 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ara"
      Height          =   300
      Left            =   2160
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmkasa.frx":272D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "MÜÞTERÝ ARA"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdsil 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      Caption         =   "Sil"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmkasa.frx":2A37
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "SÝL"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdiptal 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ýptal"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmkasa.frx":2D41
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "ÝPTAL"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmddegistir 
      BackColor       =   &H006CFBD3&
      Caption         =   "Düzelt"
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00C0E0FF&
      MouseIcon       =   "frmkasa.frx":304B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   90
      ToolTipText     =   "DEÐÝÞTÝR"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdyeni 
      BackColor       =   &H006CFBD3&
      Caption         =   "Yeni"
      Height          =   495
      Left            =   2640
      MaskColor       =   &H8000000E&
      MouseIcon       =   "frmkasa.frx":3355
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   89
      ToolTipText     =   "YENÝ KAYIT"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdraporsil 
      BackColor       =   &H000000FF&
      Caption         =   "Tüm Raporu Sil"
      Height          =   375
      Left            =   7560
      MouseIcon       =   "frmkasa.frx":365F
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "TÜM RAPORLARI SÝL(TAVSÝYE EDÝLMEZ)"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdtumgoster 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tümünü Göster"
      Height          =   375
      Left            =   6240
      MouseIcon       =   "frmkasa.frx":3969
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "TÜM ÝÞLEMLERÝN RAPORLARI"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ListBox lstmrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H00C00000&
      Height          =   4320
      ItemData        =   "frmkasa.frx":3C73
      Left            =   6240
      List            =   "frmkasa.frx":3C75
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   78
      ToolTipText     =   "TARÝHE GÖRE MÜÞTERÝ ÝÞLEMLERÝ"
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdkasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "KASA ÝÞLEMLERÝ"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmkasa.frx":3C77
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CheckBox chkkayit 
      Caption         =   "Check1"
      Height          =   315
      Left            =   2760
      TabIndex        =   48
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtmiktar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "MÝKTAR GÝRÝNÝZ"
      Top             =   3120
      Width           =   2175
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
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ÖDEME YAP"
      Top             =   3120
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
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "BORC EKLE"
      Top             =   3120
      Width           =   375
   End
   Begin VB.Frame frambilgi 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1815
      Left            =   2760
      TabIndex        =   37
      ToolTipText     =   "MÜÞTERÝ BÝLGÝLERÝ"
      Top             =   600
      Width           =   3255
      Begin VB.TextBox txtaciklama 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtadres 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txttel 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   0
         Width           =   2295
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
         Left            =   0
         TabIndex        =   41
         Top             =   1320
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
         Left            =   0
         TabIndex        =   40
         Top             =   720
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
         Left            =   0
         TabIndex        =   39
         Top             =   360
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
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.ListBox lstmusteri 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   3540
      ItemData        =   "frmkasa.frx":3F81
      Left            =   120
      List            =   "frmkasa.frx":3F83
      TabIndex        =   2
      ToolTipText     =   "MÜÞTERÝ SEÇÝNÝZ"
      Top             =   840
      Width           =   2415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   5310
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   7408
            MinWidth        =   7408
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "19.02.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1376
            MinWidth        =   1376
            TextSave        =   "14:22"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdrapor 
      BackColor       =   &H006CFBD3&
      Caption         =   "Rapor >"
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
      Left            =   5160
      MouseIcon       =   "frmkasa.frx":3F85
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "ÝÞLEMLER VE GENEL RAPOR"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtara 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "ARANACAK MÜÞTERÝ ÝSMÝNÝ GÝRÝNÝZ"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblsonodeme 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   95
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Tarih       Müþteri                  Ýþlem       Miktar"
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
      Height          =   255
      Left            =   6240
      TabIndex        =   79
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblkasa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      TabIndex        =   47
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label lblborc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   46
      ToolTipText     =   "SEÇÝLÝ MÜÞTERÝNÝN BORCU"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Top.Veresiye"
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
      Left            =   2040
      TabIndex        =   45
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Son Ödeme"
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
      Left            =   2760
      TabIndex        =   44
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Top.Borcu"
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
      Left            =   2760
      TabIndex        =   43
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MÜÞTERÝLER                   MÜÞTERÝ BÝLGÝLERÝ"
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
      TabIndex        =   36
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmkasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstmusteri As Recordset
Dim rstliste As Recordset
Dim rstmrapor As Recordset
Dim rstkasa As Recordset
Dim rstucret As Recordset
Dim rstuye As Recordset

Dim dtkafe As Database
Dim ntop As Double
Dim vodeme As Double
Dim utoplam As Double

Private Sub cmbay_Click()
On Error Resume Next
'***
With rstkasa
.Index = "indexay"
.Seek "=", Val(Mid(cmbay.Text, 1, 2))
 txtg(1) = !kira
 txtg(2) = !elektrik
 txtg(3) = !su
 txtg(4) = !internet
 txtg(5) = !telefon
 txtg(6) = !eleman
 txtg(7) = !diger
 '----
  DTPicker1(1) = Mid(!tkira, 1, 10)
  DTPicker1(2) = Mid(!telektrik, 1, 10)
  DTPicker1(3) = Mid(!tsu, 1, 10)
  DTPicker1(4) = Mid(!tinternet, 1, 10)
  DTPicker1(5) = Mid(!ttelefon, 1, 10)
  DTPicker1(6) = Mid(!teleman, 1, 10)
  DTPicker1(7) = Mid(!tdiger, 1, 10)
 '---
  chkode(1).Value = Val(Mid(!tkira, 12))
  chkode(2).Value = Val(Mid(!telektrik, 12))
  chkode(3).Value = Val(Mid(!tsu, 12))
  chkode(4).Value = Val(Mid(!tinternet, 12))
  chkode(5).Value = Val(Mid(!ttelefon, 12))
  chkode(6).Value = Val(Mid(!teleman, 12))
  chkode(7).Value = Val(Mid(!tdiger, 12))
 '---


If rstucret!parabirimi = 0 Then 'tl ise
        '*************toplamlar bölümü*************
    '---gider toplam------
    .Index = "indexay"
    .Seek "=", CDbl(Mid(cmbay.Text, 1, 2))
        txtgtop = "0"
        txtgtop = Format(CDbl(CDbl(!kira)) + CDbl(CDbl(!elektrik)) + CDbl(CDbl(!su)) + CDbl(CDbl(!internet)) + CDbl(CDbl(!telefon)) + CDbl(CDbl(!eleman)) + CDbl(CDbl(!diger)), "#00,0")
    
    '---veresiyeler-------
    rstmusteri.MoveFirst
    txtveresiye = 0
    Do Until rstmusteri.EOF
        txtveresiye = CDbl(CDbl(txtveresiye)) + CDbl(rstmusteri!borc)
        rstmusteri.MoveNext
    Loop
    '----
    txtveresiye = Format(txtveresiye, "#00,0")
    
    '---Nakit toplamý-----
    rstliste.MoveFirst
    ntop = "0"
    Do Until rstliste.EOF
        If CLng(Mid(cmbay.Text, 1, 2)) = CLng(Mid(rstliste!tarih, 4, 2)) Then
            ntop = CDbl(ntop) + CDbl(rstliste!ucret)
        End If
        rstliste.MoveNext
    Loop
    
'    If Not CDbl(ntop) <= CDbl(txtveresiye) Then
'        ntop = Format(CDbl(ntop) - CDbl(txtveresiye) + CDbl(rstkasa![ekgelir]), "#00,0")
'    End If
    
    'üyelerden gelenlerle toplanacak
    With rstuye
    .MoveFirst
    utoplam = 0
    For i = 1 To .RecordCount
        If Val(Mid(!tarih, 4, 2)) = Val(Mid(cmbay.Text, 1, 2)) Then
            utoplam = CDbl(utoplam) + CDbl(!fiyat)
        End If
        .MoveNext
    Next i
    End With
    
    ntop = CDbl(ntop) + CDbl(utoplam)
    
    'veresiyelerden ödenenlerle toplanacak
    With rstmrapor
    vodeme = "0"
    .MoveFirst
    For i = 1 To .RecordCount
        If Val(Mid(!tarih, 4, 2)) = Val(Mid(cmbay.Text, 1, 2)) Then
            If !islem = "Ödeme" Then
                vodeme = CDbl(vodeme) + CDbl(!miktar)
            End If
        End If
        .MoveNext
    Next i
    End With
    
    ntop = CDbl(ntop) + CDbl(vodeme)
    
    txtntop = CDbl(ntop) + CDbl(rstkasa![ekgelir])
    txtntop = Format(txtntop, "#00,0")
    
    '-----nakit bakiye-------------
    txtnbakiye = Format(CDbl(txtntop) - CDbl(txtgtop), "#00,0")
    
    '-----toplam bakiye -----------
    lblkasa = Format(CDbl(txtntop) + CDbl(txtveresiye) - CDbl(txtgtop), "#00,0")
    '*******************************************
  
    If CDbl(txtnbakiye) < 0 Then lblyorum = "Zararardasýnýz :("
    If CDbl(txtnbakiye) > 0 Then lblyorum = "Kardasýnýz :)"
    '***
Else
        '*************toplamlar bölümü*************
    '---gider toplam------
    .Index = "indexay"
    .Seek "=", CDbl(Mid(cmbay.Text, 1, 2))
        txtgtop = "0"
        txtgtop = Format(CDbl(CDbl(!kira)) + CDbl(CDbl(!elektrik)) + CDbl(CDbl(!su)) + CDbl(CDbl(!internet)) + CDbl(CDbl(!telefon)) + CDbl(CDbl(!eleman)) + CDbl(CDbl(!diger)), "#0.00")
    
    '---veresiyeler-------
    rstmusteri.MoveFirst
    txtveresiye = 0
    Do Until rstmusteri.EOF
        txtveresiye = CDbl(CDbl(txtveresiye)) + CDbl(rstmusteri!borc)
        rstmusteri.MoveNext
    Loop
    '----
    txtveresiye = Format(txtveresiye, "#0.00")
    
    '---Nakit toplamý-----
    rstliste.MoveFirst
    ntop = "0"
    Do Until rstliste.EOF
        If CLng(Mid(cmbay.Text, 1, 2)) = CLng(Mid(rstliste!tarih, 4, 2)) Then
            ntop = CDbl(ntop) + CDbl(rstliste!ucret)
        End If
        rstliste.MoveNext
    Loop
    
'    If Not CDbl(ntop) <= CDbl(txtveresiye) Then
'        ntop = Format(CDbl(ntop) - CDbl(txtveresiye) + CDbl(rstkasa![ekgelir]), "#0.00")
'    End If
    
    'üyelerden gelenlerle toplanacak
    With rstuye
    .MoveFirst
    utoplam = 0
    For i = 1 To .RecordCount
        If Val(Mid(!tarih, 4, 2)) = Val(Mid(cmbay.Text, 1, 2)) Then
            utoplam = CDbl(utoplam) + CDbl(!fiyat)
        End If
        .MoveNext
    Next i
    End With
    
    ntop = CDbl(ntop) + CDbl(utoplam)
    
    'veresiyelerden ödenenlerle toplanacak
    With rstmrapor
    vodeme = "0"
    .MoveFirst
    For i = 1 To .RecordCount
        If Val(Mid(!tarih, 4, 2)) = Val(Mid(cmbay.Text, 1, 2)) Then
            If !islem = "Ödeme" Then
                vodeme = CDbl(vodeme) + CDbl(!miktar)
            End If
        End If
        .MoveNext
    Next i
    End With
    
    ntop = CDbl(ntop) + CDbl(vodeme)
    
    txtntop = CDbl(ntop) + CDbl(rstkasa![ekgelir])
    txtntop = Format(txtntop, "#0.00")
    
    '-----nakit bakiye-------------
    txtnbakiye = Format(CDbl(txtntop) - CDbl(txtgtop), "#0.00")
    
    '-----toplam bakiye -----------
    lblkasa = Format(CDbl(txtntop) + CDbl(txtveresiye) - CDbl(txtgtop), "#0.00")
    '*******************************************
  
    If CDbl(txtnbakiye) < 0 Then lblyorum = "Zararardasýnýz :("
    If CDbl(txtnbakiye) > 0 Then lblyorum = "Kardasýnýz :)"
    '***
End If

End With
End Sub


Private Sub cmdara_Click()
On Error Resume Next
'*******listeleme iþlemi******
rstmusteri.MoveFirst
lstmusteri.Clear
If txtara <> "" Then
    txtara = UCase(txtara)
    For i = 1 To rstmusteri.RecordCount
        If Left(txtara, Len(txtara)) = Left(rstmusteri!AD, Len(txtara)) Then
            lstmusteri.AddItem ((lstmusteri.ListCount + 1) & "-" & rstmusteri!AD)
        End If
        rstmusteri.MoveNext
    Next i
Else
    For i = 1 To rstmusteri.RecordCount
        lstmusteri.AddItem ((lstmusteri.ListCount + 1) & "-" & rstmusteri!AD)
        rstmusteri.MoveNext
    Next i
End If
txtara.SetFocus
lstmusteri.ListIndex = 0
txtara = ""
End Sub

Private Sub cmdcikar_Click()
On Error Resume Next
'***
With rstmusteri
'---
.Index = "indexad"
.Seek "=", txtad

    If rstucret!parabirimi = 0 Then
        If txtmiktar <> "" And txtmiktar <= Val(CLng(Format(![borc], "#00,0"))) And IsNumeric(txtmiktar) = True Then
        .Edit
            ![borc] = Val(CLng(!borc)) - Val(CLng(txtmiktar))
            ![sotarih] = Format(Date, "dd.mm.yyyy")
            ![sonodeme] = Val(CLng(txtmiktar))
            .Update
            '---
            lblborc = (Format(CLng(![borc]), "#00,0"))
            If ![sonodeme] <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "###,###")
            '***raporlama için******************
            With rstmrapor
            .AddNew
            !AD = txtad
            !tarih = Date
            !islem = "Ödeme"
            !miktar = Format(txtmiktar, "#00,0")
            .Update
            End With
            '*************************************
        Else
            MsgBox "Geçersiz deðer girdiniz !!!", vbInformation
        End If
    
    Else
        If txtmiktar <> "" And IsNumeric(txtmiktar) = True Then
        .Edit
            ![borc] = CDbl(!borc) - CDbl(txtmiktar)
            ![sotarih] = Format(Date, "dd.mm.yyyy")
            ![sonodeme] = CDbl(CDbl(txtmiktar))
        .Update
            lblborc = (Format(CDbl(![borc]), "#0.00"))
            If ![sonodeme] <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
                '***raporlama için******************
                With rstmrapor
                    .AddNew
                        !AD = txtad
                        !tarih = Date
                        !islem = "Ödeme"
                        !miktar = Format(txtmiktar, "#0.00")
                    .Update
                End With
                '*************************************
        Else
            MsgBox "Geçersiz deðer girdiniz !!!", vbInformation
        End If
    End If
    
End With
'---------
txtmiktar = ""
'--------
lstmusteri_Click
'***********************
End Sub

Private Sub cmdcikis_Click()
On Error Resume Next
fraekgelir.Visible = False
End Sub

Private Sub cmddegistir_Click()
On Error Resume Next
'***
cmdyeni.Enabled = False
cmddegistir.Enabled = False
cmdsil.Enabled = False
cmdara.Enabled = False
cmdkaydet.Enabled = True
cmdiptal.Enabled = True
'---
cmdekle.Enabled = False
cmdcikar.Enabled = False
cmdkasa.Enabled = False
'---
frambilgi.Enabled = True
lstmusteri.Enabled = False
txtara.Enabled = False
cmdrapor.Enabled = False
txtad.SetFocus
'***
End Sub



Private Sub cmdegkaydet_Click()
On Error Resume Next
If txtegmiktar = "" Then txtegmiktar = "0"

With rstkasa
    .Index = "indexay"
    .Seek "=", Val(Mid(cmbay, 1, 2))
    .Edit
    
    If rstucret!parabirimi = 0 Then
        !ekgelir = Val(CLng(txtegmiktar))
    Else
        !ekgelir = CDbl(CDbl(txtegmiktar))
    End If
    
    !egaciklama = txtegaciklama
    .Update
End With
'----
fraekgelir.Visible = False
'----
End Sub

Private Sub cmdekgelir_Click()
On Error Resume Next
'----
fraekgelir.Visible = True
'----
If rstucret!parabirimi = 0 Then
    txtegmiktar = Format(rstkasa![ekgelir], "#00,0")
    txtegaciklama = rstkasa![egaciklama]
Else
    txtegmiktar = Format(rstkasa![ekgelir], "#0.00")
    txtegaciklama = rstkasa![egaciklama]
End If
'---
txtegmiktar.SetFocus
End Sub

Private Sub cmdekle_Click()
On Error Resume Next
If txtmiktar <> "" And IsNumeric(txtmiktar) = True Then
    With rstmusteri
        .Index = "indexad"
        .Seek "=", txtad
        If rstucret!parabirimi = 0 Then
                .Edit
                ![borc] = Val(CLng(!borc)) + Val(CLng(txtmiktar))
                .Update
                '---
                lblborc = (Format(CLng(![borc]), "#00,0"))
                 
                '***raporlama için******
                With rstmrapor
                .AddNew
                !AD = txtad
                !tarih = Date
                !islem = "Borc...."
                !miktar = Format(txtmiktar, "#00,0")
                .Update
                End With
        Else
                .Edit
                ![borc] = CDbl(CDbl(!borc)) + CDbl(CDbl(txtmiktar))
                .Update
                '---
                lblborc = (Format(CDbl(![borc]), "#0.00"))
                 
                '***raporlama için******
                With rstmrapor
                .AddNew
                !AD = txtad
                !tarih = Date
                !islem = "Borc...."
                !miktar = Format(txtmiktar, "#0.00")
                .Update
                End With
        End If
    End With
        
    lstmusteri_Click
Else
    MsgBox "Geçersiz deðer girdiniz !!!", vbInformation
End If
'---------
txtmiktar = ""
'--------
End Sub

Private Sub cmdgdegistir_Click()
On Error Resume Next
'***
cmdgdegistir.Enabled = False
cmdgsil.Enabled = False
cmdekgelir.Enabled = True
cmdgkaydet.Enabled = True
'--
frapanel.Enabled = True
cmbay.Enabled = False
'---
txtg(1).SetFocus
'***
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
'---
If txtsifre = rstucret!sifre Then
frasifre.Visible = False
Else
MsgBox "Yanlýþ Þifre Girdiniz!!!", vbCritical
End If
End Sub

Private Sub cmdgkaydet_Click()
On Error Resume Next
'***
With rstkasa
    .Index = "indexay"
    .Seek "=", Val(Mid(cmbay.Text, 1, 2))
       '--sýfýrlama için--
       For i = 1 To 7
            If txtg(i) = "" Then
            txtg(i) = "0"
            End If
       Next i
       '---------------------
        If .NoMatch = False Then
        .Edit
       '----
            If rstucret!parabirimi = 0 Then
                !kira = Val(CLng(txtg(1)))
                !elektrik = Val(CLng(txtg(2)))
                !su = Val(CLng(txtg(3)))
                !internet = Val(CLng(txtg(4)))
                !telefon = Val(CLng(txtg(5)))
                !eleman = Val(CLng(txtg(6)))
                !diger = Val(CLng(txtg(7)))
                !ekgelir = Val(CLng(txtegmiktar))
                !egaciklama = txtegaciklama
            Else
                !kira = CDbl(CDbl(txtg(1)))
                !elektrik = CDbl(CDbl(txtg(2)))
                !su = CDbl(CDbl(txtg(3)))
                !internet = CDbl(CDbl(txtg(4)))
                !telefon = CDbl(CDbl(txtg(5)))
                !eleman = CDbl(CDbl(txtg(6)))
                !diger = CDbl(CDbl(txtg(7)))
                !ekgelir = CDbl(CDbl(txtegmiktar))
                !egaciklama = txtegaciklama
            End If
        
        
        '----
            !tkira = DTPicker1(1) & "-" & chkode(1).Value
            !telektrik = DTPicker1(2) & "-" & chkode(2).Value
            !tsu = DTPicker1(3) & "-" & chkode(3).Value
            !tinternet = DTPicker1(4) & "-" & chkode(4).Value
            !ttelefon = DTPicker1(5) & "-" & chkode(5).Value
            !teleman = DTPicker1(6) & "-" & chkode(6).Value
            !tdiger = DTPicker1(7) & "-" & chkode(7).Value
        '----
        .Update
        End If
End With
'***
    cmdgdegistir.Enabled = True
    cmdgsil.Enabled = True
    cmdekgelir.Enabled = False
    cmdgkaydet.Enabled = False
    '--
    frapanel.Enabled = False
    cmbay.Enabled = True
    '---
    cmdgdegistir.SetFocus
'***
    cmbay_Click


End Sub

Private Sub cmdgsil_Click()
'On Error Resume Next
'***
cevap = MsgBox(cmbay.Text & " Ayýnýn iþlemlerini silmek istiyor musunuz?", vbYesNo + vbCritical)
If cevap = vbYes Then
'---
With rstkasa
    .Index = "indexay"
    .Seek "=", Val(Mid(cmbay.Text, 1, 2))
       '--sýfýrlama için--
       For i = 1 To 7
            txtg(i) = "0"
            DTPicker1(i) = Date
            chkode(i).Value = 0
       Next i
       '---------------------
        If .NoMatch = False Then
        .Edit
       If rstucret!parabirimi = 0 Then
                !kira = Val(CLng(txtg(1)))
                !elektrik = Val(CLng(txtg(2)))
                !su = Val(CLng(txtg(3)))
                !internet = Val(CLng(txtg(4)))
                !telefon = Val(CLng(txtg(5)))
                !eleman = Val(CLng(txtg(6)))
                !diger = Val(CLng(txtg(7)))
                !ekgelir = Val(CLng(txtegmiktar))
                !egaciklama = txtegaciklama
            Else
                !kira = CDbl(CDbl(txtg(1)))
                !elektrik = CDbl(CDbl(txtg(2)))
                !su = CDbl(CDbl(txtg(3)))
                !internet = CDbl(CDbl(txtg(4)))
                !telefon = CDbl(CDbl(txtg(5)))
                !eleman = CDbl(CDbl(txtg(6)))
                !diger = CDbl(CDbl(txtg(7)))
                !ekgelir = CDbl(CDbl(txtegmiktar))
                !egaciklama = txtegaciklama
            End If
            
            !tkira = DTPicker1(1) & "-" & chkode(1).Value
            !telektrik = DTPicker1(2) & "-" & chkode(2).Value
            !tsu = DTPicker1(3) & "-" & chkode(3).Value
            !tinternet = DTPicker1(4) & "-" & chkode(4).Value
            !ttelefon = DTPicker1(5) & "-" & chkode(5).Value
            !teleman = DTPicker1(6) & "-" & chkode(6).Value
            !tdiger = DTPicker1(7) & "-" & chkode(7).Value
        '----
        .Update
        End If
End With
'***
cmbay_Click
'***
End If
'----
End Sub

Private Sub cmdiptal_Click()
On Error Resume Next
'***gösterim refresh******************
With rstmusteri
'----
txtad = ![AD]
txttel = ![tel]
txtadres = ![adres]
txtaciklama = ![aciklama]
'---
If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
'---
If rstucret!parabirimi = 0 Then
    If !borc <> "" Then
    lblborc = (Format(CLng(![borc]), "#00,0"))
    Else
    lblborc = "0"
    End If
Else
    If !borc <> "" Then
    lblborc = (Format(CDbl(![borc]), "#0.00"))
    Else
    lblborc = "0"
    End If
End If

End With
'**************************************
'***
cmdyeni.Enabled = True
cmddegistir.Enabled = True
cmdsil.Enabled = True
cmdara.Enabled = True
cmdkaydet.Enabled = False
cmdiptal.Enabled = False
'---
cmdekle.Enabled = True
cmdcikar.Enabled = True
cmdkasa.Enabled = True
'---
frambilgi.Enabled = False
lstmusteri.Enabled = True
txtara.Enabled = True
cmdrapor.Enabled = True
cmdyeni.SetFocus
'---
chkkayit.Value = 0
'***
End Sub

Private Sub cmdkasa_Click()
On Error Resume Next
'***
If cmdkasa.Caption = "KASA ÝÞLEMLERÝ" Then
    '---
    cmdkasa.Caption = "MÜSTERÝ ÝÞLEMLERÝ"
    Label8.Caption = "TOP.BAKÝYE"
    lblkasa = ""
    '---
    frakasamenu.Top = 120
    frakasamenu.Left = 120
    frakasamenu.Visible = True
    '---
    '---load yordamý gibi----ilk iþlemleri yaptýrýyoruz----
    cmbay.ListIndex = Val(Mid(Date, 4, 2)) - 1
    cmbay_Click
    '-------------------------------------------------------------------------
Else
    '---
    cmdkasa.Caption = "KASA ÝÞLEMLERÝ"
    Label8.Caption = "Top.Veresiye"
    '-----veresiyeler------------
    lblkasa = 0
    rstmusteri.MoveFirst
    
    If rstucret!parabirimi = 0 Then
        Do Until rstmusteri.EOF
        lblkasa = CLng(Val(lblkasa)) + Val(rstmusteri!borc)
        rstmusteri.MoveNext
        Loop
        lblkasa = Format(lblkasa, "#00,0")
    Else
        Do Until rstmusteri.EOF
        lblkasa = CDbl(CDbl(lblkasa)) + CDbl(rstmusteri!borc)
        rstmusteri.MoveNext
        Loop
        lblkasa = Format(lblkasa, "#0.00")
    End If
    
    '***
    frakasamenu.Visible = False
    '---
End If
'***
End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next
'***
With rstmusteri
'---
If txtad <> "" Then
    If chkkayit.Value = 1 Then
        .MoveFirst
        Var = False
        For i = 1 To .RecordCount
            If !AD = UCase(txtad) Then
                Var = True
            End If
            .MoveNext
        Next i
        
        If Var = False Then
            .AddNew
            ![AD] = UCase(txtad)
            ![tel] = txttel
            ![adres] = txtadres
            ![aciklama] = txtaciklama
            ![borc] = "0"
            .Update
        Else
            MsgBox "Bu isimde müþteri var baþka isim seçiniz !!!", vbInformation
        End If
    Else
        .Edit
        ![AD] = UCase(txtad)
        ![tel] = txttel
        ![adres] = txtadres
        ![aciklama] = txtaciklama
        .Update
    End If
Else
MsgBox "Ad Soyad girmediniz iþlem iptal edildi", vbCritical
End If

'***gösterim refresh******************
lstmusteri.Clear
.MoveFirst
Do Until .EOF
lstmusteri.AddItem (lstmusteri.ListCount + 1 & "-" & ![AD])
.MoveNext
Loop
'---
txtad = ![AD]
txttel = ![tel]
txtadres = ![adres]
txtaciklama = ![aciklama]


If rstucret!parabirimi = 0 Then
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#00,0")
    If !borc <> "" Then
        lblborc = (Format(CLng(![borc]), "#00,0"))
    Else
        lblborc = "0"
    End If
Else
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
    If !borc <> "" Then
        lblborc = (Format(CDbl(![borc]), "#0.00"))
    Else
        lblborc = "0"
    End If
End If

'**************************************
End With
'***-*-*-*-*-*-*-*-**--*-*-**--****-*--*-***-**-*-**
cmdyeni.Enabled = True
cmddegistir.Enabled = True
cmdsil.Enabled = True
cmdara.Enabled = True
cmdkaydet.Enabled = False
cmdiptal.Enabled = False
'---
cmdekle.Enabled = True
cmdcikar.Enabled = True
cmdkasa.Enabled = True
'---
frambilgi.Enabled = False
lstmusteri.Enabled = True
txtara.Enabled = True
cmdrapor.Enabled = True
cmdyeni.SetFocus
'---
chkkayit.Value = 0
'***
End Sub

Private Sub cmdrapor_Click()
'---
If cmdrapor.Caption = "Rapor >" Then
cmdrapor.Caption = "Rapor <"
Me.Width = 11700
cmdkasa.Enabled = False
Else
Me.Width = 6225
cmdrapor.Caption = "Rapor >"
cmdkasa.Enabled = True
End If
'---
End Sub

Private Sub cmdraporsil_Click()
On Error Resume Next
'*****
cevap = MsgBox("Tüm Raporu silmek istiyor musunuz?" & vbCrLf & "NOT: Eðer raporu silerseniz hesaplar kasadan düþülecektir(Tavsiye Edilmez)", vbYesNo + vbCritical)
If cevap = vbYes Then
'---
With rstmrapor
'--
.MoveFirst
Do Until .EOF
.Delete
.MoveNext
Loop
'---
lstmrapor.Clear
'--
End With
'---
End If
'****
End Sub

Private Sub cmdsil_Click()
On Error Resume Next
'***
With rstmusteri
    cevap = MsgBox("Kaydý silmek istiyor musunuz?", vbCritical + vbYesNo)
    If cevap = vbYes Then
    .Delete
'---
    lstmusteri.Clear
.MoveFirst
Do Until .EOF
    lstmusteri.AddItem (lstmusteri.ListCount + 1 & "-" & ![AD])
    .MoveNext
Loop
'---
txtad = ""
txttel = ""
txtadres = ""
txtaciklama = ""
lblborc = ""
lblsonodeme = ""
'---
txtad = ![AD]
txttel = ![tel]
txtadres = ![adres]
txtaciklama = ![aciklama]

If rstucret!parabirimi = 0 Then
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#00,0")
    If !borc <> "" Then
        lblborc = (Format(CLng(![borc]), "#00,0"))
    Else
        lblborc = "0"
    End If
Else
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
    If !borc <> "" Then
        lblborc = (Format(CDbl(![borc]), "#0.00"))
    Else
        lblborc = "0"
    End If
End If

End If
'***
End With
End Sub

Private Sub cmdtumgoster_Click()
On Error Resume Next
With rstmrapor
'------------------------------
.MoveFirst
lstmrapor.Clear
'----
Do Until .EOF
lstmrapor.AddItem (!tarih & "...." & !AD & "......." & !islem & "......." & !miktar)
.MoveNext
Loop
'------------------------------
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
'---
cmdekle.Enabled = False
cmdcikar.Enabled = False
cmdkasa.Enabled = False
'---
frambilgi.Enabled = True
lstmusteri.Enabled = False
txtara.Enabled = False
cmdrapor.Enabled = False
txtad.SetFocus
'---
chkkayit.Value = 1
'***temizlik
txtad = ""
txttel = ""
txtadres = ""
txtaciklama = ""
txtmiktar = ""
lblsonodeme = ""
lblborc = ""
'****************
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
SendKeys vbTab
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Set dtkafe = OpenDatabase(App.Path & "\datakafe")
Set rstmusteri = dtkafe.OpenRecordset("musteriler")
Set rstkasa = dtkafe.OpenRecordset("kasa")
Set rstucret = dtkafe.OpenRecordset("ucretler")
Set rstliste = dtkafe.OpenRecordset("raporlar")
Set rstmrapor = dtkafe.OpenRecordset("mrapor")
Set rstuye = dtkafe.OpenRecordset("uyeler")
'---
'***
rstmusteri.MoveFirst
txtara.SetFocus

rstucret.MoveFirst
If rstucret!konkasa = 1 Then frasifre.Visible = True
'---
With rstmusteri
txtad = ![AD]
txttel = ![tel]
txtadres = ![adres]
txtaciklama = ![aciklama]

If rstucret!parabirimi = 0 Then
    '---
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
    '---
    If !borc <> "" Then
    lblborc = (Format(CLng(![borc]), "#00,0"))
    Else
    lblborc = "0"
    End If
    '---
    .MoveFirst
    For i = 1 To .RecordCount
    lstmusteri.AddItem (lstmusteri.ListCount + 1 & "-" & ![AD])
    .MoveNext
    Next i
    '***
  
    '-----veresiyelere------------
    rstmusteri.MoveFirst
    Do Until rstmusteri.EOF
    lblkasa = CLng(Val(lblkasa)) + Val(rstmusteri!borc)
    rstmusteri.MoveNext
    Loop
    lblkasa = Format(lblkasa, "#00,0")
    '***********************************************
Else
    '---
    If !sonodeme <> "" Then lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
        If !borc <> "" Then
            lblborc = (Format(CDbl(![borc]), "#0.00"))
        Else
            lblborc = "0"
        End If
    '---
    Do Until .EOF
    lstmusteri.AddItem (lstmusteri.ListCount + 1 & "-" & ![AD])
    .MoveNext
    Loop
    '***
    
    '-----veresiyelere------------
    rstmusteri.MoveFirst
    lblkasa = 0
    Do Until rstmusteri.EOF
        lblkasa = CDbl(lblkasa) + CDbl(Format(rstmusteri!borc, "#0.00"))
        rstmusteri.MoveNext
    Loop
    
    lblkasa = Format(lblkasa, "#0.00")
    '***********************************************
End If
'********
End With
'-----------------------------------------------------------------------------
Me.Width = 6225
Me.Height = 6060
'***
frasifre.Move 0, 0
'**
LISTELE

RENK_VER

End Sub

Private Sub lstmrapor_Click()
On Error Resume Next
lstmrapor.ToolTipText = lstmrapor.Text
End Sub

Private Sub lstmusteri_Click()
On Error Resume Next
'***
With rstmusteri
.Index = "indexad"
.Seek "=", Mid(lstmusteri.Text, InStr(1, lstmusteri.Text, "-") + 1)
txtad = ![AD]
txttel = ![tel]
txtadres = ![adres]
txtaciklama = ![aciklama]


If rstucret!parabirimi = 0 Then
    '---
    If !sonodeme <> "" Then
    lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#00,0")
    Else
    lblsonodeme = "Hiç Ödeme Yapmadý"
    End If
    '---
    If !borc <> "" Then
    lblborc = (Format(CLng(![borc]), "#00,0"))
    Else
    lblborc = "0"
    End If
    '---
Else
    If !sonodeme <> "" Then
    lblsonodeme = ![sotarih] & " - " & Format(![sonodeme], "#0.00")
    Else
    lblsonodeme = "Hiç Ödeme Yapmadý"
    End If
    '---
    If !borc <> "" Then
    lblborc = (Format(CDbl(![borc]), "#0.00"))
    Else
    lblborc = "0"
    End If
End If
'---
End With
'***
LISTELE
End Sub

Private Sub txtara_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdara.Value = True
End Sub

Private Sub txtaciklama_LostFocus()
On Error Resume Next
cmdkaydet.SetFocus
End Sub

Private Sub txtegmiktar_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtegmiktar = Format(txtegmiktar, "#00,0")
    txtegmiktar.SelStart = Len(txtegmiktar)
End If
End Sub

Private Sub txtg_Change(Index As Integer)
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtg(Index) = Format(txtg(Index), "#00,0")
    txtg(Index).SelStart = Len(txtg(Index))
End If
End Sub

Private Sub txtmiktar_Change()
On Error Resume Next
If rstucret!parabirimi = 0 Then
    txtmiktar = Format(txtmiktar, "#00,0")
    txtmiktar.SelStart = Len(txtmiktar)
End If
End Sub

Private Sub LISTELE()
'On Error Resume Next
'*******müþteri raporlamasý**********************************************
With rstmrapor
'------------------------------
.MoveFirst
lstmrapor.Clear
'----
Do Until .EOF
If !AD = txtad Then
lstmrapor.AddItem (!tarih & "...." & !AD & "......." & !islem & "......." & !miktar)
End If
.MoveNext
Loop
'------------------------------
End With
'************************************************************************
End Sub

Private Sub txtsifre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdgiris_Click
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
Label22.ForeColor = vbWhite
Label26.ForeColor = vbWhite
End Sub
