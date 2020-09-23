VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmana 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Ana Menü (Hesap) V"
   ClientHeight    =   8550
   ClientLeft      =   2130
   ClientTop       =   1470
   ClientWidth     =   10755
   Icon            =   "frmana.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10755
   Begin VB.Frame fraversiyon 
      BackColor       =   &H00FFE7E3&
      Height          =   3735
      Left            =   120
      TabIndex        =   231
      Top             =   120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdvyukselt 
         BackColor       =   &H006CFBD3&
         Caption         =   "Versiyon Yükselt"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdgcikis 
         BackColor       =   &H006CFBD3&
         Caption         =   "X"
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   233
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtgbilgi 
         Appearance      =   0  'Flat
         Height          =   2775
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   232
         Top             =   360
         Width           =   3135
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   0
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Label lblgbaslik 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ".::YENÝ VERSÝYON::."
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
         TabIndex        =   235
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdresim 
      Height          =   255
      Left            =   1200
      Picture         =   "frmana.frx":144A
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmduye 
      BackColor       =   &H006CFBD3&
      Caption         =   "Üyeler"
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
      Left            =   2400
      MouseIcon       =   "frmana.frx":1C00
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   229
      ToolTipText     =   "MÜÞTERÝLER VE HESAPLAR"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdsohbet 
      BackColor       =   &H006CFBD3&
      Caption         =   "Browser"
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
      Left            =   1800
      MouseIcon       =   "frmana.frx":2042
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "WEB BROWSER"
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdnot 
      BackColor       =   &H006CFBD3&
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
      Height          =   375
      Left            =   2400
      MouseIcon       =   "frmana.frx":2484
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "NOT VE UYARILAR"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdrapor 
      BackColor       =   &H006CFBD3&
      Caption         =   "Rapor"
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
      Left            =   960
      MouseIcon       =   "frmana.frx":28C6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "GÜNLÜK VE AYLIK RAPOR"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdprohakkinda 
      BackColor       =   &H006CFBD3&
      Caption         =   "Hakkýnda"
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
      MouseIcon       =   "frmana.frx":2D08
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "PROGRAM HAKKINDA"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdayar 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ayarlar"
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
      Left            =   2760
      MouseIcon       =   "frmana.frx":301A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "PROGRAM AYARLARI"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdkasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kasa "
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
      Left            =   1320
      MouseIcon       =   "frmana.frx":3324
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "MÜÞTERÝLER VE HESAPLAR"
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame frasifre 
      BackColor       =   &H00BFA3C9&
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   600
      TabIndex        =   70
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton cmdgiris 
         BackColor       =   &H006CFBD3&
         Caption         =   "*GÝRÝÞ* )>)>)>"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtsifre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   72
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdx 
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
         Height          =   240
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   71
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
         Index           =   3
         Left            =   120
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 34"
      Height          =   255
      Index           =   34
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 35"
      Height          =   255
      Index           =   35
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 36"
      Height          =   255
      Index           =   36
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "MASA SEÇÝNÝZ"
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 37"
      Height          =   255
      Index           =   37
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 38"
      Height          =   255
      Index           =   38
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 39"
      Height          =   255
      Index           =   39
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 40"
      Height          =   255
      Index           =   40
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 41"
      Height          =   255
      Index           =   41
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "MASA SEÇÝNÝZ"
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 42"
      Height          =   255
      Index           =   42
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 43"
      Height          =   255
      Index           =   43
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 44"
      Height          =   255
      Index           =   44
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 45"
      Height          =   255
      Index           =   45
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 46"
      Height          =   255
      Index           =   46
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 47"
      Height          =   255
      Index           =   47
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 48"
      Height          =   255
      Index           =   48
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 49"
      Height          =   255
      Index           =   49
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 50"
      Height          =   255
      Index           =   50
      Left            =   7320
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 16"
      Height          =   255
      Index           =   16
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 17"
      Height          =   255
      Index           =   17
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 18"
      Height          =   255
      Index           =   18
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 19"
      Height          =   255
      Index           =   19
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 20"
      Height          =   255
      Index           =   20
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 21"
      Height          =   255
      Index           =   21
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 22"
      Height          =   255
      Index           =   22
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 23"
      Height          =   255
      Index           =   23
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 24"
      Height          =   255
      Index           =   24
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 25"
      Height          =   255
      Index           =   25
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 26"
      Height          =   255
      Index           =   26
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "MASA SEÇÝNÝZ"
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 27"
      Height          =   255
      Index           =   27
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 28"
      Height          =   255
      Index           =   28
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 29"
      Height          =   255
      Index           =   29
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 30"
      Height          =   255
      Index           =   30
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 31"
      Height          =   255
      Index           =   31
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "MASA SEÇÝNÝZ"
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 32"
      Height          =   255
      Index           =   32
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 33"
      Height          =   255
      Index           =   33
      Left            =   3720
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 11"
      Height          =   255
      Index           =   11
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 10"
      Height          =   255
      Index           =   10
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 09"
      Height          =   255
      Index           =   9
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 08"
      Height          =   255
      Index           =   8
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 07"
      Height          =   255
      Index           =   7
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 06"
      Height          =   255
      Index           =   6
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 05"
      Height          =   255
      Index           =   5
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 04"
      Height          =   255
      Index           =   4
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 03"
      Height          =   255
      Index           =   3
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 02"
      Height          =   255
      Index           =   2
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 01"
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmana.frx":3766
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "MASA SEÇÝNÝZ"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 12"
      Height          =   255
      Index           =   12
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 13"
      Height          =   255
      Index           =   13
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 14"
      Height          =   255
      Index           =   14
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdmasa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa 15"
      Height          =   255
      Index           =   15
      Left            =   120
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Messenger 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdclient 
      BackColor       =   &H006CFBD3&
      Caption         =   "Client"
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
      MouseIcon       =   "frmana.frx":3A70
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "CLIENTLERÝ AÇ"
      Top             =   6480
      Width           =   735
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
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmana.frx":3D7A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "PROGRAMDAN ÇIK"
      Top             =   6000
      Width           =   495
   End
   Begin MSComCtl2.UpDown upd2 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "2.MASA SEÇ"
      Top             =   6000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upd1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "1.MASA SEÇ"
      Top             =   6000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtno 
      Height          =   285
      Left            =   360
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   0
   End
   Begin VB.TextBox txtm2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtm1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton cmdaktar 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      Caption         =   "Masa Aktar"
      Height          =   375
      Left            =   720
      MouseIcon       =   "frmana.frx":41BC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "MASALARI AKTAR"
      Top             =   6000
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   62
      Top             =   8175
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "23.04.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "13:18"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Mehmet ALTINEL & Türker ÖZER"
            TextSave        =   "Mehmet ALTINEL & Türker ÖZER"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   735
      Left            =   3600
      TabIndex        =   67
      Top             =   7440
      Width           =   7215
      Begin VB.Label lblsite 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4800
         MouseIcon       =   "frmana.frx":44C6
         MousePointer    =   99  'Custom
         TabIndex        =   227
         ToolTipText     =   "TIKLAYIN SÝTEMÝZÝ ZÝYARET EDÝN"
         Top             =   360
         Width           =   3255
      End
   End
   Begin SHDocVwCtl.WebBrowser lblguncel1 
      Height          =   975
      Left            =   0
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   7440
      Width           =   5595
      ExtentX         =   9869
      ExtentY         =   1720
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   4680
      TabIndex        =   236
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblbuyukfiyat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "25.000.000"
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
      Height          =   495
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   228
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " Masalar     Açýlýþ    Süre      Ücret"
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
      Index           =   0
      Left            =   120
      TabIndex        =   61
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   34
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   226
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   35
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   225
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   9000
      TabIndex        =   224
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   8280
      TabIndex        =   223
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   36
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   222
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   9000
      TabIndex        =   221
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   8280
      TabIndex        =   220
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   9000
      TabIndex        =   219
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   8280
      TabIndex        =   218
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   37
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   217
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   9000
      TabIndex        =   216
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   8280
      TabIndex        =   215
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   38
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   214
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   39
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   213
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   9000
      TabIndex        =   212
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   8280
      TabIndex        =   211
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   40
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   210
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   9000
      TabIndex        =   209
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   8280
      TabIndex        =   208
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   41
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   207
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   9000
      TabIndex        =   206
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   8280
      TabIndex        =   205
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   9000
      TabIndex        =   204
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   8280
      TabIndex        =   203
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   42
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   202
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   9000
      TabIndex        =   201
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   8280
      TabIndex        =   200
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   43
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   199
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   9000
      TabIndex        =   198
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   8280
      TabIndex        =   197
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   44
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   196
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   9000
      TabIndex        =   195
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   8280
      TabIndex        =   194
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   45
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   193
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   9000
      TabIndex        =   192
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   8280
      TabIndex        =   191
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   46
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   190
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   47
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   189
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   48
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   188
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   49
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   187
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   50
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   186
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   9000
      TabIndex        =   185
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   9000
      TabIndex        =   184
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   9000
      TabIndex        =   183
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   9000
      TabIndex        =   182
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   9000
      TabIndex        =   181
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   8280
      TabIndex        =   180
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   8280
      TabIndex        =   179
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   48
      Left            =   8280
      TabIndex        =   178
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   49
      Left            =   8280
      TabIndex        =   177
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   50
      Left            =   8280
      TabIndex        =   176
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "EÐER DAHA FAZLA MAKÝNANIZ VARSA BÝZE BÝLDÝRÝNÝZ"
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
      Height          =   375
      Left            =   7320
      TabIndex        =   175
      Top             =   6600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   4680
      TabIndex        =   174
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   4680
      TabIndex        =   173
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   4680
      TabIndex        =   172
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   4680
      TabIndex        =   171
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   4680
      TabIndex        =   170
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   5400
      TabIndex        =   169
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   5400
      TabIndex        =   168
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   5400
      TabIndex        =   167
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   5400
      TabIndex        =   166
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   5400
      TabIndex        =   165
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   5400
      TabIndex        =   164
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   16
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   163
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   17
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   162
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   18
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   161
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   19
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   160
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   20
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   159
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   21
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   158
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   4680
      TabIndex        =   157
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   5400
      TabIndex        =   156
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   22
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   155
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   4680
      TabIndex        =   154
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   5400
      TabIndex        =   153
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   23
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   152
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   4680
      TabIndex        =   151
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   5400
      TabIndex        =   150
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   24
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   149
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   4680
      TabIndex        =   148
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   5400
      TabIndex        =   147
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   25
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   146
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   4680
      TabIndex        =   145
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   5400
      TabIndex        =   144
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   4680
      TabIndex        =   143
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   5400
      TabIndex        =   142
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   26
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   141
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   4680
      TabIndex        =   140
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   5400
      TabIndex        =   139
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   27
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   138
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   4680
      TabIndex        =   137
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   5400
      TabIndex        =   136
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   28
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   135
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   29
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   134
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   4680
      TabIndex        =   133
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   5400
      TabIndex        =   132
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   30
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   131
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   4680
      TabIndex        =   130
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   5400
      TabIndex        =   129
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   4680
      TabIndex        =   128
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   5400
      TabIndex        =   127
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   31
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   126
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   4680
      TabIndex        =   125
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   5400
      TabIndex        =   124
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   32
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   123
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   33
      Left            =   6120
      MousePointer    =   99  'Custom
      TabIndex        =   122
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   11
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   121
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   10
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   120
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   9
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   119
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   8
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   118
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   7
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   117
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   6
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   116
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   5
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   115
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   4
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   114
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   3
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   113
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   2
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   112
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   1
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   111
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   110
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   109
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   108
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   107
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   106
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   105
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   104
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   103
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   102
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   101
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   100
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1080
      TabIndex        =   99
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   98
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   97
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   96
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   95
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   94
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   93
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   92
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   91
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   90
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   89
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1080
      TabIndex        =   88
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1080
      TabIndex        =   87
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   1080
      TabIndex        =   86
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1800
      TabIndex        =   85
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   1800
      TabIndex        =   84
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   1800
      TabIndex        =   83
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   12
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   82
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   13
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   81
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   14
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   80
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label u 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Index           =   15
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   79
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label s 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1800
      TabIndex        =   78
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1080
      TabIndex        =   77
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " Masalar     Açýlýþ     Süre      Ücret"
      DragMode        =   1  'Automatic
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
      Left            =   7320
      TabIndex        =   69
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " Masalar     Açýlýþ    Süre      Ücret"
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
      Left            =   3720
      TabIndex        =   68
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnu 
      Caption         =   "Mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuana 
         Caption         =   "Ana Menu"
      End
      Begin VB.Menu mnukasa 
         Caption         =   "Kasa"
      End
      Begin VB.Menu mnurapo 
         Caption         =   "Rapor"
      End
      Begin VB.Menu mnuclient 
         Caption         =   "Client"
      End
      Begin VB.Menu ayrac 
         Caption         =   "-"
      End
      Begin VB.Menu mnunot 
         Caption         =   "Not"
      End
      Begin VB.Menu mnubrowser 
         Caption         =   "Browser"
      End
      Begin VB.Menu ayrac2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhakkinda 
         Caption         =   "Hakkýnda"
      End
      Begin VB.Menu mnucikis 
         Caption         =   "Çýkýþ"
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "mnu2"
      Visible         =   0   'False
      Begin VB.Menu mnuad 
         Caption         =   "Adlandýr"
      End
   End
End
Attribute VB_Name = "frmana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'ses apisi
'***
Dim basla As Date
Dim bitis As Date
Dim sure As Currency
Dim ucret As Currency
Dim MS As Long
'***
Dim rstucret As Recordset
Dim rstnot As Recordset
Dim rstuyar As Recordset
Dim rstkafe As Recordset
Dim dtkafe As Database

'***********************************************************************************************************
'***********bu kýsm msn gibi simge halinde görünmesi için***************************************************

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
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
'*************************************************************************************************************************
Dim VERSIYON As String
Dim masano As String 'masa isimlendirme için
Private Sub chkaktar1_Click()
On Error Resume Next
'***
If chkaktar.Value = 1 Then
txtno = 1
Else
txtno = ""
End If
'***
End Sub
Private Sub a_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
On Error Resume Next
    txtm2 = Index
    txtm1 = Source.Index
    cmdaktar_Click
End Sub

Private Sub a_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    txtno = Index
    
    If a(Index) = "" Then
        frmmasa.mnunkapat.Enabled = False
        frmmasa.mnuvkapat.Enabled = False
        frmmasa.mnuekucret.Enabled = False
    Else
        frmmasa.mnuhesapac.Enabled = False
        frmmasa.mnunkapat.Enabled = True
        frmmasa.mnuvkapat.Enabled = True
        frmmasa.mnuekucret.Enabled = True
    End If
    PopupMenu frmmasa.mnu
    
End If

End Sub

Private Sub a_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
a(Index).ToolTipText = "SÜRÜKLE VE AKTAR"
a(Index).MousePointer = a(1).MousePointer
a(Index).MouseIcon = a(1).MouseIcon
a(Index).BorderStyle = 0
End Sub

Private Sub cmdaktar_Click()
On Error Resume Next
'------------------------------------
rstucret.MoveFirst
MS = rstucret![msayisi]
'-----------------
If txtm1 <> "" And txtm2 <> "" Then
'*************************************
i = txtm1
j = txtm2
'***
yrda = a(i)
yrds = s(i)
yrdu = u(i)
'***
a(i) = a(j)
s(i) = s(j)
u(i) = u(j)
'***
a(j) = yrda
s(j) = yrds
u(j) = yrdu
'**************************************
With rstkafe
    rstkafe.Index = "indexkod"
    rstkafe.Seek "=", i
    SU1 = rstkafe![sucret]
    SS1 = rstkafe![ssure]
    SNOT1 = rstkafe!masanot
    SEU1 = !eucret
    SNOTSS1 = !notssure
    SNOTSU1 = !notsucret
    ALTUCRET1 = !secucret2
    ACIKLAMA1 = !aciklama
    '---
    rstkafe.Index = "indexkod"
    rstkafe.Seek "=", j
    SU2 = rstkafe![sucret]
    SS2 = rstkafe![ssure]
    SNOT2 = rstkafe!masanot
    SEU2 = !eucret
    SNOTSS2 = !notssure
    SNOTSU2 = !notsucret
    ALTUCRET2 = !secucret2
    ACIKLAMA2 = !aciklama
    '***
    rstkafe.Index = "indexkod"
    rstkafe.Seek "=", i
    rstkafe.Edit
    rstkafe![sucret] = SU2
    rstkafe![ssure] = SS2
    rstkafe!masanot = SNOT2
    !eucret = SEU2
    !notssure = SNOTSS2
    !notsucret = SNOTSU2
    !secucret2 = ALTUCRET2
    !aciklama = ACIKLAMA2
    rstkafe.Update
    '---
    rstkafe.Index = "indexkod"
    rstkafe.Seek "=", j
    rstkafe.Edit
    rstkafe![sucret] = SU1
    rstkafe![ssure] = SS1
    rstkafe!masanot = SNOT1
    !eucret = SEU1
    !notssure = SNOTSS1
    !notsucret = SNOTSU1
    !secucret2 = ALTUCRET1
    !aciklama = ACIKLAMA1
    rstkafe.Update
End With
'**************************************
KAYDET
'***
End If
'**************************************
For j = 1 To MS
a(j).BackColor = vbWhite
s(j).BackColor = vbWhite
u(j).BackColor = vbWhite
Next j
'***
For j = 1 To MS
a(j).ForeColor = vbBlack
s(j).ForeColor = vbBlack
u(j).ForeColor = vbBlack
Next j
'**************************************

txtm1 = ""
txtm2 = ""
upd1.Value = 1
upd2.Value = 1
txtm1 = ""
txtm2 = ""
'***
End Sub

Private Sub cmdayar_Click()
On Error Resume Next
frasifre.Visible = True
txtsifre = ""
txtsifre.SetFocus
End Sub

Private Sub cmdcikis_Click()
On Error Resume Next
'***
Timer1.Interval = 0
'***
'***yedek alýnmasý incelemesi*****
If rstucret!yedek = 1 Then
    cevap = MsgBox("Yedek almak istiyor musunuz?", vbYesNo + vbInformation)
    If cevap = vbYes Then
        Shell App.Path & "\ÖKH Yedekle.exe", vbNormalFocus
        End
    Else
        End
    End If
Else
    End
End If
'*********************************
End Sub

Private Sub cmdclient_Click()
frmclient.Show
End Sub

Private Sub cmdgcikis_Click()
On Error Resume Next
fraversiyon.Visible = False
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
    If txtsifre = rstucret!sifre Then
    frasifre.Visible = False
    frmayar.Show
    Else
    MsgBox "Yanlýþ þifre girdiniz !!!", vbCritical
    End If
'---
End Sub

Private Sub cmdkasa_Click()
'***
frmkasa.Show
'***
End Sub

Private Sub cmdmasa_Click(Index As Integer)
'***
txtno = Index
frmmasa.Show
Unload frmmasa
frmmasa.Show
'***
End Sub

Private Sub cmdmasa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
masano = Index
PopupMenu mnu2
End If
End Sub

Private Sub cmdmasa_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdmasa(Index).ToolTipText = "MASA SEÇÝNÝZ"
cmdmasa(Index).MouseIcon = cmdmasa(1).MouseIcon
End Sub


Private Sub cmdnot_Click()
frmnot.Show
End Sub

Private Sub cmdprohakkinda_Click()
'***
frmhakkinda.Show
'***
End Sub

Private Sub cmdrapor_Click()
frmrapor.Show
End Sub

Private Sub cmdsohbet_Click()
'***
frmsohbet.Show
'***
End Sub

Private Sub cmduye_Click()
On Error Resume Next
frmuye.Show
End Sub

Private Sub cmdvyukselt_Click()
On Error Resume Next
cevap = MsgBox("Yeni versiyonu yüklemek istiyor musunuz?" + vbCrLf + "Not: Programýnýzýn veri tabanýnýn yedeðini alýnýz. (datakafe.mdb)", vbYesNo + vbInformation)
If cevap = vbYes Then
    Shell App.Path & "\ÖKH Güncelle.exe", vbNormalFocus 'güncellemeyi yapacak olan program çalýþtýrýlýyor
    End                                           'program kapatýlýyor ki güncellensin
End If

End Sub

Private Sub cmdx_Click()
frasifre.Visible = False
End Sub

Private Sub Form_Activate()
mnu.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'-------------------------
Unload frmilk
Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

    
    If KeyCode = vbKeyF1 Then
    If ShiftDown And CtrlDown And AltDown Then
        '-----site bilgisi----------
        rstucret.MoveFirst
        cevap = InputBox("Programcý Þifresini Girin(MED)", ".::Programcý Þifresi::.")
        If cevap = "/***/" Then
        cevap2 = InputBox("Web Adresini Giriniz", ".::Web Adresi Deðiþtir::.", rstucret![webadresi])
        rstucret.Edit
        rstucret![webadresi] = cevap2
        rstucret.Update
        MsgBox "Web Adresi Deðiþtirildi :)"
        End If
    End If
    End If
'****************************************************************
If KeyCode = vbKeyF11 Then
If ShiftDown And AltDown And CtrlDown Then
frmayar.Show
End If
End If
'****************************************************************
End Sub

Private Sub Form_Load()
On Error Resume Next

'*********************icon load**************************
Tray.cbSize = Len(Tray)
    Tray.hwnd = Messenger.hwnd
    Tray.uID = 1&
    Tray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Tray.uCallbackMessage = WM_MOUSEMOVE
    Tray.hIcon = Me.Icon
    Tray.szTip = "Özer Kafe Hesap" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, Tray
'*********************************************************

'---------------çalýþýp çalýþmadýðýna bakma çalýþýyorsa açma-----------
If App.PrevInstance Then
    MsgBox "Program zaten çalýþýyor !!!", vbInformation
    End
End If
'----------------------------------------------
'***
VERSIYON = App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = Me.Caption & " " & VERSIYON & "::."
'***
Timer1.Interval = 1000
'***
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstkafe = dtkafe.OpenRecordset("masalar")
Set rstuyar = dtkafe.OpenRecordset("uyarilar")
Set rstucret = dtkafe.OpenRecordset("ucretler")
Set rstnot = dtkafe.OpenRecordset("notlar")
'**MASA SAYISI ÝÇÝN****
rstucret.MoveFirst
MS = rstucret![msayisi]
'***********************
'---MASA GÖRÜNÜMLERÝ---
If MS = 50 Then Label2.Visible = True
'---visible olaylarý---
For i = 1 To MS
cmdmasa(i).Visible = True
a(i).Visible = True
s(i).Visible = True
u(i).Visible = True
Next i
'---geniþlik---
Me.Width = 3675
If MS >= 16 Then Me.Width = 7320
If MS >= 34 Then Me.Width = 10900
'-------------

'-----görünüm------
Me.Top = 0
Me.Left = Screen.Width - Me.Width
'---------------------

'***********************
'***bilgiler yükleniyor****
For i = 1 To MS
a(i) = rstkafe![acilis1]
s(i) = rstkafe![sure1]
u(i) = rstkafe![ucret1]

'masa adlarý için
If rstkafe!masaad <> "" Then
    cmdmasa(i).Caption = rstkafe![masaad]
End If

'uyeler için masa tuþuna resim
If rstkafe!uye = 1 Then
    cmdmasa(i).Picture = cmdresim.Picture
    cmdmasa(i).Caption = ""
End If

rstkafe.MoveNext
Next i
'****

'*****ilk açýlýþta clientin kapalý oluþu**********
With rstucret
.MoveFirst
.Edit
!client = 0
.Update
End With
'*******************************************
'web sayfasýna giriþ
rstucret.MoveFirst
lblguncel1.Navigate (rstucret!webadresi)

'otomatik clientlere baðlan
rstucret.MoveFirst
If rstucret!otobaglan = "1" Then
    cmdclient.Value = True
    frmclient.cmdgizle.Value = True
End If

RENK_VER
GUNCELLE

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblbuyukfiyat.Visible = False

For i = 1 To 50
    a(i).BorderStyle = 1
Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel = True
'***
KAYDET
'***
Me.Hide
End Sub



Private Sub lblguncel1_DownloadComplete()
On Error Resume Next
Dim V As String
V = lblguncel1.LocationName
If Mid(V, 1, 8) = "versiyon" Then
    If Mid(V, 9) <> VERSIYON Then
'        cevap = MsgBox("Yeni bir versiyon var  " & Mid(V, 9) & "  Yüklemek istermisiniz?", vbYesNo + vbInformation)
        If cevap = vbYes Then
            
        End If
    End If
End If
End Sub

Private Sub lblsite_Click()
On Error Resume Next
Shell "start http://www.ozerkafe.com", vbHide
End Sub

Private Sub lblsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblsite.ForeColor = vbYellow
End Sub

Private Sub Messenger_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Me.PopupMenu mnu
End Sub
Private Sub Messenger_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Rec  As Boolean, Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    If Rec = False Then
        Rec = True
        Select Case Msg
            Case WM_RBUTTONUP:
                Me.PopupMenu mnu
            Case WM_LBUTTONDBLCLK
                mnuana_Click
        End Select
        Rec = False
    End If
End Sub



Private Sub mnuad_Click()
On Error Resume Next
cevap = InputBox("Masa ismini giriniz", "Masa Ýsimlendirme")
    rstkafe.Index = "indexkod"
    rstkafe.Seek "=", masano
    rstkafe.Edit
    rstkafe![masaad] = cevap
    rstkafe.Update
    If rstkafe!masaad <> "" Then
        cmdmasa(masano).Caption = rstkafe!masaad
    End If
End Sub

Private Sub mnuana_Click()
On Error Resume Next
Me.Show
End Sub

Private Sub mnubrowser_Click()
On Error Resume Next
cmdsohbet_Click
End Sub

Private Sub mnucikis_Click()
On Error Resume Next
cmdcikis_Click
End Sub

Private Sub mnuclient_Click()
On Error Resume Next
cmdclient_Click
End Sub

Private Sub mnuhakkinda_Click()
On Error Resume Next
cmdprohakkinda_Click
End Sub

Private Sub mnukasa_Click()
On Error Resume Next
cmdkasa_Click
End Sub

Private Sub mnunot_Click()
On Error Resume Next
cmdnot_Click
End Sub

Private Sub mnurapo_Click()
On Error Resume Next
cmdrapor_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'*****************************************************************
lblsite.ForeColor = vbWhite
'not uyarýsý için
rstuyar.MoveFirst
    Do Until rstuyar.EOF
        If rstuyar![utarih] = CStr(Date) And Format(rstuyar![usaat], "hh:mm") = CStr(Format(Time(), "hh:mm")) Then
            SINYAL
            cevap = MsgBox(rstuyar![unot] + vbCrLf + vbCrLf & "UYARI SÝLÝNSÝN MÝ?(Evet)   YOKSA NOT BÖLÜMÜNE EKLENSÝN MÝ?(Hayýr)", vbInformation + vbYesNo)
            If cevap = vbYes Then
                rstuyar.Delete
            Else
                rstnot.MoveFirst
                rstnot.AddNew
                rstnot!baslik = rstuyar!ubaslik
                rstnot!Not = rstuyar!unot
                rstnot!tarih = rstuyar!utarih & " - " & CStr(Time())
                rstnot.Update
                rstuyar.Delete
            End If
        End If
        rstuyar.MoveNext
    Loop
'*****************************************************************
'sýnýrlama uyarýsý için
    rstucret.MoveFirst
    MS = rstucret![msayisi]
        For i = 1 To MS
            '***
            If s(i) <> "" Then
                '---
                rstkafe.Index = "indexkod"
                rstkafe.Seek "=", i
                '---
                '*********sýnýrlanan masaya renk deðiþimi formun loadýnda************
                    If rstkafe![notsucret] <> "" And rstkafe![notsucret] = "1" Or rstkafe![notssure] <> "" And rstkafe![notssure] = "1" Then
                        a(i).BackColor = vbRed
                        s(i).BackColor = vbRed
                        u(i).BackColor = vbRed
                        '---
                        a(i).ForeColor = vbWhite
                        s(i).ForeColor = vbWhite
                        u(i).ForeColor = vbWhite
                        
                        'sýnýrlanan hesap dolunca cliente gönderme
                        '********************************************
                        If rstucret!otokapat = 1 Then
                            With frmclient
                            If rstucret!client = 1 Then
                                'If frmclient.Visible = False Then
                                    .Timer2.Enabled = False
                                    If .chkotodurum.Value = 1 Then
                                        .chkotodurum.Value = 0
                                        .optm(i).Value = True
                                        .cmdkulkapat2.Value = True
                                        .chkotodurum.Value = 1
                                    Else
                                        .optm(i).Value = True
                                        .cmdkulkapat2.Value = True
                                    End If
                                    .Timer2.Enabled = True
                                'End If
                            End If
                            End With
                        End If
                        '*********************************************
                        
                    End If
                    
                    
                    '******************sure sýnýrlama uyarýsý*********************************
                    RHH = CLng(Val(Mid(rstkafe![ssure], 1, 2)))
                    RMM = CLng(Val(Mid(rstkafe![ssure], 4, 2)))
                    SSHH = CLng(Val(Mid(s(i), 1, 2)))
                    SSMM = CLng(Val(Mid(s(i), 4, 2)))
                    '---
                    '****
                    If rstkafe![notssure] <> "1" And rstkafe![ssure] <> "" Then
                        If SSHH > RHH Or (SSHH = RHH And SSMM >= RMM) Then
                            '***
                           
                            SINYAL 'sinyal yordamý
                            'MsgBox "Masa " & i & " için uyarý !!!  " & s(i) & " dakika " & u(i) & " TL", vbCritical
                            'cmdmasa_Click (i)

                            '***
                            With rstkafe
                                .Edit
                                ![notsucret] = "0"
                                ![notssure] = "1"
                                .Update
                            End With
                            '***
                        End If
                    '---
                    End If
                    
                    '**************ucret sýnýrlama uyarýsý*******************
                    If rstkafe![notsucret] <> "1" And rstkafe![sucret] <> "" And u(i) <> "" Then
                        If rstucret!parabirimi = 0 Then
                            If CDbl(u(i)) >= CDbl(rstkafe![sucret]) Then
                                SINYAL 'sinyal yordamý
                                'MsgBox "Masa " & i & " için uyarý !!!  " & s(i) & " dakika " & u(i) & " TL", vbCritical
                                'cmdmasa_Click (i)
                                    With rstkafe
                                        .Edit
                                        ![notsucret] = "1"
                                        ![notssure] = "0"
                                        .Update
                                    End With
                            End If
                        Else
                            If CDbl(u(i)) >= CDbl(rstkafe![sucret]) Then
                                SINYAL 'sinyal yordamý
                                'MsgBox "Masa " & i & " için uyarý !!!  " & s(i) & " dakika " & u(i) & " YTL", vbCritical
                                'cmdmasa_Click (i)
                                    With rstkafe
                                        .Edit
                                        ![notsucret] = "1"
                                        ![notssure] = "0"
                                        .Update
                                    End With
                            End If
                        End If
                    End If
                End If
            Next i
'*******************yükleme baþlýyor*******************************************
            For i = 1 To MS
            '***
                If a(i) <> "" Then
                    '-----------------------------------------
                    ahh1 = Val(Mid(a(i), 1, 2))
                    amm = Val(Mid(a(i), 4, 2))
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
                            shh = Val(bhh - ahh) \ 60
                            smm = Val(bhh - ahh) - (60 * shh)
                        End If
'*************************ucret için******************************************
                        rstkafe.Index = "indexkod"
                        rstkafe.Seek "=", i
      
                        With rstucret
                        
                        '***alternatif ücret seçimi(ucret2)--------
                        Dim ucret
                        If rstkafe!secucret2 = 1 Then
                            ucret = !ucret2
                        Else
                            ucret = !ucret
                        End If
                        '-----------------------------------------
                        
                        '--deðerleri deðiþkenlere atýyoruz
                        Dim lngUCRET, lngBASUCRET, lngBIRIM, lngKAFEUCRET, lngEKUCRET, lngATILANSIFIR, lngKURUS
                        lngATILANSIFIR = 1000000
                        lngKURUS = 100
                        lngUCRET = CDbl(ucret) * lngATILANSIFIR
                        lngBIRIM = CDbl(!birim) * lngATILANSIFIR
                        lngBASUCRET = CDbl(!basucret) * lngATILANSIFIR
                        lngKAFEUCRET = CDbl(rstkafe!ucret) * lngATILANSIFIR
                        lngEKUCRET = CDbl(rstkafe!eucret) * lngATILANSIFIR
                            
                            If rstucret!parabirimi = 0 Then 'parabirimi TL ise
                                  lblbuyukfiyat.Width = 2055
                                 
                                   u(i) = 0
                                   u(i) = (((((shh * 60) + smm) * (Val(ucret) \ 60)) \ Val(!birim)) * Val(!birim))
                                   If (shh * 60 + smm) * (Val(ucret) \ 60) < Val(!basucret) Then
                                       u(i) = Val(!basucret)
                                   Else
                                       u(i) = ((((shh * 60) + smm) * (Val(ucret) \ 60) \ Val(!birim)) * Val(!birim))
                                       
                                       'yukarý yuvarlama
                                       If rstucret!yyuvarla = 1 Then
                                       If ((shh * 60) + smm) - (((shh * 60) + smm) \ (Val(!birim) / 10000)) * (Val(!birim) / 10000) >= 3 Then
                                           u(i) = ((((shh * 60) + smm) * (Val(ucret) \ 60) \ Val(!birim)) * Val(!birim)) + Val(!birim)
                                       End If
                                       End If
                                       
                                   End If
                                   u(i) = Val(CDbl(u(i))) + Val(rstkafe!eucret)
                                   u(i) = Format(u(i), "#00,0")
                                   u(i) = (CDbl(u(i)) \ CDbl(!birim)) * CDbl(!birim)
                                    
                                   u(i) = Format(u(i), "#00,0")
                            
                            Else 'parabirimi yeni türklirasý ise YTL
                                 Dim AAS
                                u(i) = 0
                                
                                If (shh * 60 + smm) * (Val(lngUCRET) \ 60) < Val(lngBASUCRET) Then
                                    u(i) = lngBASUCRET / lngATILANSIFIR / lngKURUS
                                   
                                Else
                                    
                                    u(i) = ((((((shh * 60 + smm) * Val((lngUCRET) \ 60)) \ Val(lngBIRIM)) * Val(lngBIRIM)) + Val(lngEKUCRET)) / lngATILANSIFIR) / lngKURUS
                                  
                                    'yukarý yuvarlama
                                    If rstucret!yyuvarla = 1 Then
                                    If ((shh * 60) + smm) - (((shh * 60) + smm) \ (Val(lngBIRIM) / 1000000)) * (Val(lngBIRIM) / 1000000) >= 3 Then
                                        u(i) = (((((((shh * 60 + smm) * (lngUCRET) \ 60)) \ Val(lngBIRIM)) * Val(lngBIRIM)) + Val(lngBIRIM)) / lngATILANSIFIR) / lngKURUS
                                    
                                    End If
                                    End If
                                     
                                End If
                                   u(i) = Format(u(i), "#0.00")
                                   u(i) = ((CDbl(u(i)) * 100) \ (CDbl(lngBIRIM) / 1000000)) * (CDbl(lngBIRIM) / 1000000) / 100
                    
                                    u(i) = CDbl(u(i)) + CDbl(rstkafe!eucret)
                                    
                                    u(i) = Format(u(i), "#0.00")
                            End If
                            
                        End With
                        '***********************************************************
                            '------------------------------------
                            If shh < 10 Then shh = "0" & shh
                            '------------------------------------
                            If smm < 10 Then smm = "0" & smm
                            '---------------------------
                                s(i) = shh & ":" & smm
                            '---------------------------
                            End If
                    '--------------
                    Next i
                '***
                KAYDET
            '***
End Sub

Private Sub txtm1_Change()
On Error Resume Next
i = txtm1
k = txtm2
'***
rstucret.MoveFirst
MS = rstucret![msayisi]
'---
For j = 1 To MS
a(j).BackColor = vbWhite
s(j).BackColor = vbWhite
u(j).BackColor = vbWhite
Next j
'***
a(i).BackColor = vbBlue
s(i).BackColor = vbBlue
u(i).BackColor = vbBlue
'***
a(k).BackColor = vbBlue
s(k).BackColor = vbBlue
u(k).BackColor = vbBlue
'***
For j = 1 To MS
a(j).ForeColor = vbBlack
s(j).ForeColor = vbBlack
u(j).ForeColor = vbBlack
Next j
'***
a(i).ForeColor = vbWhite
s(i).ForeColor = vbWhite
u(i).ForeColor = vbWhite
'***
a(k).ForeColor = vbWhite
s(k).ForeColor = vbWhite
u(k).ForeColor = vbWhite
'***
End Sub

Private Sub txtm2_Change()
On Error Resume Next
i = txtm2
k = txtm1
'***
rstucret.MoveFirst
MS = rstucret![msayisi]
'---
For j = 1 To MS
a(j).BackColor = vbWhite
s(j).BackColor = vbWhite
u(j).BackColor = vbWhite
Next j
'***
a(i).BackColor = vbBlue
s(i).BackColor = vbBlue
u(i).BackColor = vbBlue
'***
a(k).BackColor = vbBlue
s(k).BackColor = vbBlue
u(k).BackColor = vbBlue

'***
For j = 1 To MS
a(j).ForeColor = vbBlack
s(j).ForeColor = vbBlack
u(j).ForeColor = vbBlack
Next j
'***
a(i).ForeColor = vbWhite
s(i).ForeColor = vbWhite
u(i).ForeColor = vbWhite
'***
a(k).ForeColor = vbWhite
s(k).ForeColor = vbWhite
u(k).ForeColor = vbWhite
'***
End Sub



Private Sub txtno_Change()
On Error Resume Next
If txtno < 10 Then txtno = "0" & Val(txtno)
End Sub

Private Sub txtsifre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdgiris_Click
End Sub

Private Sub u_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If u(Index) <> "" Then
    lblbuyukfiyat.Move u(Index).Left + u(Index).Width - lblbuyukfiyat.Width, u(Index).Top
    lblbuyukfiyat = u(Index)
    lblbuyukfiyat.Visible = True
End If
End Sub

Private Sub upd1_Change()
upd1.Max = 1
rstucret.MoveFirst
MS = rstucret![msayisi]
'---
upd1.Min = MS
txtm1 = upd1.Value
End Sub

Private Sub upd2_Change()
upd2.Max = 1
rstucret.MoveFirst
MS = rstucret![msayisi]
'---
upd2.Min = MS
txtm2 = upd2.Value
End Sub
Sub KAYDET()
On Error Resume Next
'---
rstucret.MoveFirst
MS = rstucret![msayisi]
'***
rstkafe.MoveFirst
'***
For i = 1 To MS
rstkafe.Edit
rstkafe![acilis1] = a(i)
rstkafe![sure1] = s(i)
rstkafe![ucret1] = u(i)
rstkafe.Update
rstkafe.MoveNext
Next i
'***
End Sub

Private Sub SINYAL()
sndPlaySound (App.Path & "\uyari.wav"), 0
End Sub
Private Sub KSINYAL()
sndPlaySound (App.Path & "\kasa.wav"), 0
End Sub

Private Sub GUNCELLE()
On Error Resume Next
Dim Version As String, News As String
Dim Site As String

Site = rstucret!versite
Version = Inet1.OpenURL(Site & "Versiyon.txt")
    
Dim Uzunluk
Uzunluk = Len(Version)
    

If Uzunluk <> 0 And Uzunluk < 10 Then  'eðer versiyon bilgisine ulaþýlamýyorsa yada 404 hata sayfasý geliyorsa güncelleme iptal edilir.
    If Not Trim(Version) = "kilit" Then 'her ihtimale karþý kilitleme durumlarýnda
        If Trim(Version) > App.Major & "." & App.Minor & "." & App.Revision Then
            fraversiyon.Visible = True
            txtgbilgi = Replace(Inet1.OpenURL(Site & "Yenilikler.txt"), Chr(10), vbCrLf)
        End If
    Else
        MsgBox "Programýnýz Program Sahibi Tarafýndan Kilitlenmiþtir..."
        End
    End If
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

For i = 1 To MS
    a(i).ForeColor = vbBlack
    s(i).ForeColor = vbBlack
    u(i).ForeColor = vbBlack
    lblbuyukfiyat.ForeColor = &HFF0000
Next i

Label5.ForeColor = vbWhite
lblgbaslik.ForeColor = vbWhite

End Sub
