VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Clientler::."
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ControlBox      =   0   'False
   Icon            =   "frmclient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   7080
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Timer Timer2 
         Left            =   960
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Left            =   480
         Top             =   0
      End
      Begin VB.Timer Timer3 
         Left            =   1440
         Top             =   0
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   1
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   2
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   3
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   4
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   5
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   6
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   7
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   8
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   9
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   10
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   11
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   12
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   13
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   14
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   15
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   16
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   17
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   18
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   19
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   20
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   21
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   22
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   23
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   24
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   25
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   26
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   27
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   28
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   29
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   30
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   31
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   32
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   33
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   34
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   35
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   36
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   37
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   38
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   39
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   40
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   41
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   42
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   43
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   44
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   45
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   46
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   47
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   48
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   49
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock winsck 
         Index           =   50
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.CommandButton cmdhgonder 
      BackColor       =   &H006CFBD3&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8760
      MouseIcon       =   "frmclient.frx":144A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   111
      ToolTipText     =   "ÞÝMDÝ HESAP GÖNDER"
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdgizle 
      BackColor       =   &H006CFBD3&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8640
      MouseIcon       =   "frmclient.frx":1754
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "FORMU GÝZLE"
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdcik 
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
      Height          =   495
      Left            =   8640
      MouseIcon       =   "frmclient.frx":1A5E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "FORMU KAPAT BAÐLANTIYI KOPAR"
      Top             =   360
      Width           =   495
   End
   Begin VB.Frame frasifre 
      BackColor       =   &H00BFA3C9&
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   3480
      TabIndex        =   104
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
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
         TabIndex        =   81
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtsifresor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   600
         PasswordChar    =   "*"
         TabIndex        =   79
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdgiris 
         BackColor       =   &H006CFBD3&
         Caption         =   "*GÝRÝÞ* )>)>)>"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label14 
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
         TabIndex        =   106
         Top             =   0
         Width           =   2535
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
         TabIndex        =   105
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraayar 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      DragMode        =   1  'Automatic
      Height          =   5295
      Left            =   0
      TabIndex        =   99
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   120
         TabIndex        =   100
         Top             =   120
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
            TabIndex        =   116
            ToolTipText     =   "WINDOWS AÇILIÞINDA MASA ÜSTÜ ARAYÜZÜ ÇIKSIN"
            Top             =   3480
            Width           =   3615
         End
         Begin VB.CheckBox chkkontor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404080&
            Caption         =   "Kontörlü sistemi devrede"
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
            TabIndex        =   115
            ToolTipText     =   "EKRAN KORUYUCUDA FLASH ANÝMASYONU GÖSTERÝR"
            Top             =   3240
            Width           =   3375
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
            TabIndex        =   114
            ToolTipText     =   "EKRAN KORUYUCUDA FLASH ANÝMASYONU GÖSTERÝR"
            Top             =   2280
            Width           =   3375
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
            TabIndex        =   113
            ToolTipText     =   "MASA ÜSTÜNDE HESAP GÖSTERÝLMESÝNÝ ÝSTÝYORSANIZ SEÇÝNÝZ"
            Top             =   3000
            Width           =   3375
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
            TabIndex        =   112
            ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
            Top             =   2760
            Width           =   3015
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
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   110
            Text            =   "frmclient.frx":1D68
            Top             =   4440
            Width           =   3495
         End
         Begin VB.CommandButton cmdsite 
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
            Height          =   375
            Left            =   2760
            MouseIcon       =   "frmclient.frx":1DB5
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "YASAKLANAN SÝTELER"
            Top             =   3960
            Width           =   855
         End
         Begin VB.CommandButton cmdiptal 
            BackColor       =   &H006CFBD3&
            Caption         =   "Kapat"
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
            MouseIcon       =   "frmclient.frx":20BF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "KAPAT CLÝENTLER FORMUNA DÖN"
            Top             =   3960
            Width           =   855
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
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmclient.frx":23C9
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "AYARLARI KAYDET VE CLIENT'LERE GÖNDER"
            Top             =   3960
            Width           =   1575
         End
         Begin VB.CommandButton cmdsifregoster 
            BackColor       =   &H006CFBD3&
            Caption         =   "Göster"
            Height          =   300
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "ÞÝFREGÖSTER/GÝZLE"
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txtek 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "EKLEMEK ÝSTEDÝKLERÝNÝZ EKRAN KORUYUCUDA GÖRÜNÜR"
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox txtkafeadi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "KAFENÝZÝN ÝSMÝNÝ GÝRÝN"
            Top             =   960
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
            TabIndex        =   5
            ToolTipText     =   "BÝLGÝSAYAR KULLANIMA KAPALIYKEN MÜÞTERÝNÝN SERVERLA  MESAJLAÞMASINI ÝSTÝYORSANIZ SEÇÝN"
            Top             =   2520
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
            TabIndex        =   2
            ToolTipText     =   "WINDOWS AÇILIÞINDA EKRANI KULLANIMA KAPATMAK ÝSTÝYORSANIZ SEÇÝN"
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox txtsifre 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   0
            ToolTipText     =   "CLIENT PROGRAMINA ÞÝFRE VERÝN AYARLARA SÝZDEN BAÞKASI ULAÞAMASIN"
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label13 
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
            TabIndex        =   103
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label12 
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
            TabIndex        =   102
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label11 
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
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmdtumybaslat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Tüm Y Baþlat"
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmclient.frx":26D3
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdtumkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Tüm. Kapat"
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmclient.frx":29DD
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdkulkapat2 
      Caption         =   "Kulllanýma kapat"
      Height          =   375
      Left            =   1080
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdkulac2 
      Caption         =   "Kullanýma aç"
      Height          =   375
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkctrl 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Ctrl+Alt+Del Açýk"
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
      ToolTipText     =   "CLIENTTE CTRL+ALT+DEL AÇIK OLSUN"
      Top             =   5520
      Width           =   1815
   End
   Begin MSComctlLib.Slider sldses 
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      ToolTipText     =   "SES AYARLAMA"
      Top             =   5400
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      MousePointer    =   99
      Min             =   10
      Max             =   65535
      SelStart        =   32500
      Value           =   32500
      TextPosition    =   1
   End
   Begin VB.CommandButton cmdclientayarla 
      BackColor       =   &H006CFBD3&
      Caption         =   "Client Ayarla"
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frmclient.frx":2CE7
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "TÜM CLIENTLERÝ BURADAN AYARLAYIN ZAMANDAN KAZANIN"
      Top             =   960
      Width           =   855
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   5850
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "11:36"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "18.04.2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Mehmet ALTINEL   &   Türker ÖZER"
            TextSave        =   "Mehmet ALTINEL   &   Türker ÖZER"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstdurum 
      BackColor       =   &H00D5F9FF&
      Height          =   2400
      ItemData        =   "frmclient.frx":2FF1
      Left            =   6120
      List            =   "frmclient.frx":2FF3
      TabIndex        =   22
      ToolTipText     =   "KAPATILACAK PROGRAMI SEÇÝNÝZ"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton cmdprokapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kapat"
      Height          =   300
      Left            =   8520
      MouseIcon       =   "frmclient.frx":2FF5
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "SEÇÝLEN PROGRAMI KAPAT"
      Top             =   1920
      Width           =   615
   End
   Begin MSComCtl2.UpDown updhgonderdk 
      Height          =   300
      Left            =   8160
      TabIndex        =   95
      Top             =   4800
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      Max             =   60
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txthgonderdk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "01"
      ToolTipText     =   "SEÇÝLEN DAKÝKADA BÝR HESAP GÖNDER"
      Top             =   4800
      Width           =   255
   End
   Begin VB.CheckBox chkhgonder 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Hesaplarý Gönder         dk"
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
      TabIndex        =   23
      ToolTipText     =   "CLIENTLERE HESAP GÖNDER"
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton cmddgoster 
      BackColor       =   &H006CFBD3&
      Caption         =   "Göster"
      Height          =   300
      Left            =   7800
      MouseIcon       =   "frmclient.frx":32FF
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "CLÝENT ÜZERÝNDE ÇALIÞAN PROGRAMLARI GÖSTER"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtmsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      ToolTipText     =   "MESAJ YAZMA EKRANI"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtms 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   92
      Text            =   "0"
      ToolTipText     =   "TOPLAM MASA SAYISI"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtmno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   90
      ToolTipText     =   "CLIENT NO"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtport 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   86
      ToolTipText     =   "CLIENT BAÐLANMA PORTU"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   85
      ToolTipText     =   "SERVER ÝP"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 50"
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
      Index           =   50
      Left            =   4920
      TabIndex        =   78
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 49"
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
      Index           =   49
      Left            =   4920
      TabIndex        =   77
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 48"
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
      Index           =   48
      Left            =   4920
      TabIndex        =   76
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 47"
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
      Index           =   47
      Left            =   4920
      TabIndex        =   75
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 46"
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
      Index           =   46
      Left            =   4920
      TabIndex        =   74
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 45"
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
      Index           =   45
      Left            =   4920
      TabIndex        =   73
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 44"
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
      Index           =   44
      Left            =   4920
      TabIndex        =   72
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 43"
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
      Index           =   43
      Left            =   4920
      TabIndex        =   71
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 42"
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
      Index           =   42
      Left            =   4920
      TabIndex        =   70
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 41"
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
      Index           =   41
      Left            =   4920
      TabIndex        =   69
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 40"
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
      Index           =   40
      Left            =   3720
      TabIndex        =   68
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 39"
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
      Index           =   39
      Left            =   3720
      TabIndex        =   67
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 38"
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
      Index           =   38
      Left            =   3720
      TabIndex        =   66
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 37"
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
      Index           =   37
      Left            =   3720
      TabIndex        =   65
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 36"
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
      Index           =   36
      Left            =   3720
      TabIndex        =   64
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 35"
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
      Index           =   35
      Left            =   3720
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 34"
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
      Index           =   34
      Left            =   3720
      TabIndex        =   62
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 33"
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
      Index           =   33
      Left            =   3720
      TabIndex        =   61
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 32"
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
      Index           =   32
      Left            =   3720
      TabIndex        =   60
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 31"
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
      Index           =   31
      Left            =   3720
      TabIndex        =   59
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 30"
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
      Index           =   30
      Left            =   2520
      TabIndex        =   58
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 29"
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
      Index           =   29
      Left            =   2520
      TabIndex        =   57
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 28"
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
      Index           =   28
      Left            =   2520
      TabIndex        =   56
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 27"
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
      Index           =   27
      Left            =   2520
      TabIndex        =   55
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 26"
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
      Index           =   26
      Left            =   2520
      TabIndex        =   54
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 25"
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
      Index           =   25
      Left            =   2520
      TabIndex        =   53
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 24"
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
      Index           =   24
      Left            =   2520
      TabIndex        =   52
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 23"
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
      Index           =   23
      Left            =   2520
      TabIndex        =   51
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 22"
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
      Index           =   22
      Left            =   2520
      TabIndex        =   50
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 21"
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
      Index           =   21
      Left            =   2520
      TabIndex        =   49
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 20"
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
      Index           =   20
      Left            =   1320
      TabIndex        =   48
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 19"
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
      Index           =   19
      Left            =   1320
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 18"
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
      Index           =   18
      Left            =   1320
      TabIndex        =   45
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 17"
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
      Index           =   17
      Left            =   1320
      TabIndex        =   44
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 16"
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
      Index           =   16
      Left            =   1320
      TabIndex        =   43
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 15"
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
      Index           =   15
      Left            =   1320
      TabIndex        =   42
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 14"
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
      Index           =   14
      Left            =   1320
      TabIndex        =   41
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 13"
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
      Index           =   13
      Left            =   1320
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 12"
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
      Index           =   12
      Left            =   1320
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 11"
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
      Index           =   11
      Left            =   1320
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 10"
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
      Index           =   10
      Left            =   120
      TabIndex        =   37
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 9"
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
      Index           =   9
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 8"
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
      Index           =   8
      Left            =   120
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 7"
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
      Index           =   7
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 6"
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
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 5"
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
      Index           =   5
      Left            =   120
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 4"
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
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 3"
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
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 2"
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
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optm 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Masa 1"
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
      Index           =   1
      Left            =   120
      TabIndex        =   28
      ToolTipText     =   "ÝÞLEM YAPILACAK MASA"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdkulkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kullanýma Kapat"
      Height          =   495
      Left            =   6120
      MouseIcon       =   "frmclient.frx":3609
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "CLIENT KULLANIMA KAPAT"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdcdrom 
      BackColor       =   &H006CFBD3&
      Caption         =   "CD Rom A"
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frmclient.frx":3913
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "CLIENT CD-ROM AC/KAPAT"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdkulac 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kullanýma Aç"
      Height          =   495
      Left            =   6120
      MouseIcon       =   "frmclient.frx":3C1D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "CLIENT KULLANIMA AÇ"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdybaslat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Y. Baþlat"
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmclient.frx":3F27
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "CLIENT YENÝDEN BAÞLAT"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdkapat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Kapat"
      Height          =   255
      Left            =   7200
      MouseIcon       =   "frmclient.frx":4231
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "CLIENT KAPAT"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdgonder 
      BackColor       =   &H006CFBD3&
      Caption         =   "Gönder"
      Height          =   1095
      Left            =   4320
      MouseIcon       =   "frmclient.frx":453B
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "MESAJ GÖNDER"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtekran 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      ToolTipText     =   "MESAÞLAÞMA EKRANI"
      Top             =   360
      Width           =   4095
   End
   Begin VB.CheckBox chkotodurum 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "Oto. Durum Göster"
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
      Left            =   2040
      TabIndex        =   107
      ToolTipText     =   "MASA SEÇÝLDÝÐÝNDE OTOMATÝK DURUMUNU GÖSTER"
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DURUM"
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
      Left            =   6360
      TabIndex        =   108
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Ses Ayarý"
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
      TabIndex        =   96
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
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
      Left            =   4680
      TabIndex        =   94
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "< MESAJ >"
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
      TabIndex        =   82
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
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
      Left            =   6840
      TabIndex        =   89
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
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
      Left            =   7800
      TabIndex        =   88
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER IP"
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
      Left            =   2160
      TabIndex        =   87
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    M A S A L A R"
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
      TabIndex        =   84
      Top             =   1560
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "*KOMUTLAR*"
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
      Left            =   6000
      TabIndex        =   83
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx
Dim yy

Dim III As Integer

Dim mesaj As String
Dim dtkafe As Database
Dim rstkafe As Recordset
Dim rstucret As Recordset
Dim rstclient As Recordset
Dim rstuye As Recordset
Dim rstclientayar As Recordset

Private Sub chkctrl_Click()
On Error Resume Next
i = Val(txtmno)
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    If chkctrl.Value = 1 Then
        winsck(i).SendData ("*CTRLAC*")
    Else
        winsck(i).SendData ("*CTRLKAPAT*")
    End If
End If
'------
End Sub

Private Sub chkekran_Click()
On Error Resume Next
If chkekran.Value = 1 Then
    txtkafeadi.Enabled = True
    txtek.Enabled = True
    txtkafeadi.SetFocus
Else
    txtkafeadi.Enabled = False
    txtek.Enabled = False
End If
End Sub

Private Sub chkhgonder_Click()
On Error Resume Next
rstucret.MoveFirst
rstucret.Edit
    rstucret!hgonder = chkhgonder.Value
rstucret.Update
End Sub

Private Sub cmdayarkapat_Click()
fraayar.Visible = False
Me.Height = 6600
Me.Width = 8985
End Sub

Private Sub chkotodurum_Click()
On Error Resume Next
rstucret.MoveFirst
rstucret.Edit
    rstucret!otodurum = chkotodurum.Value
rstucret.Update
End Sub

Private Sub cmdcik_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdclientayarla_Click()
On Error Resume Next
'******client ayarlarýnýn yüklenmesi*********
With rstclientayar
    .MoveFirst
    txtsifre = !sifre
    chkekran.Value = !ekran
    txtkafeadi = !kafeadi
    txtek = !ek
    chkchat.Value = !chat
    chkflash.Value = !flash
    chkeyaz.Value = !eyaz
    chkhesap.Value = !hesap
    chkkontor.Value = !KONTOR
    chkarayuz.Value = !arayuz
End With
'********************************************
frasifre.Visible = True
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next
If txtsifresor = rstucret!sifre Then
    Me.Height = fraayar.Height + 350
    Me.Width = fraayar.Width + 75
    fraayar.Visible = True
    frasifre.Visible = False
End If
txtsifresor = ""
End Sub

Private Sub cmdgizle_Click()
On Error Resume Next
Me.Hide
End Sub

Private Sub cmdhgonder_Click()
On Error Resume Next
If chkotodurum.Value = 1 Then
    chkotodurum.Value = 0
    i = txtmno
    optm(i).Value = True
    If Not winsck(i).State <> sckConnected Then
        If frmana.a(i) <> "" Then
            winsck(i).SendData ("*U*" & frmana.u(i).Caption)
        Else
            winsck(i).SendData ("*U*" & "0")
        End If
    End If
    chkotodurum.Value = 1
Else
    i = txtmno
    optm(i).Value = True
    If Not winsck(i).State <> sckConnected Then
        If frmana.a(i) <> "" Then
            winsck(i).SendData ("*U*" & frmana.u(i).Caption)
        Else
            winsck(i).SendData ("*U*" & "0")
        End If
    End If
End If

End Sub

Private Sub cmdiptal_Click()
On Error Resume Next
fraayar.Visible = False
Me.Height = 6600
Me.Width = 9315
End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next
Timer2.Enabled = False
'******client ayarlarýnýn kaydedilmesi*********
With rstclientayar
    .MoveFirst
    .Edit
        !sifre = txtsifre
        !ekran = chkekran.Value
        !kafeadi = txtkafeadi
        !ek = txtek
        !chat = chkchat.Value
        !eyaz = chkeyaz.Value
        !hesap = chkhesap.Value
        !flash = chkflash.Value
        !KONTOR = chkkontor.Value
        !arayuz = chkarayuz.Value
    .Update
End With
'*******************************************

For i = 1 To txtms
    If winsck(i).State = sckConnected Then
        winsck(i).SendData ("AYAR" & txtsifre & "~" & chkekran.Value & "~" & txtkafeadi & "~" & txtek & "~" & chkchat.Value & "~" & chkeyaz.Value & "~" & chkhesap.Value & "~" & chkflash.Value & "~" & chkkontor.Value & "~" & chkarayuz.Value)
    End If
Next i
MsgBox "Ayarlarýnýz Kaydedildi", vbInformation
Timer2.Enabled = True

End Sub

Private Sub cmdkulac_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " kullanýma açýlsýn mý?", vbYesNo)
        If cevap = vbYes Then
            winsck(i).SendData ("*MASAAC*")
        End If
End If
'------

Timer2.Enabled = True
End Sub

Private Sub cmdcdrom_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
   If cmdcdrom.Caption = "CD Rom A" Then
    cevap = MsgBox("Masa " & i & " CD-ROM Açýlsýn mý?", vbYesNo)
        If cevap = vbYes Then
            winsck(i).SendData ("*CDROMAC*")
            cmdcdrom.Caption = "CD Rom K"
        End If
    Else
    cevap = MsgBox("Masa " & i & " CD-ROM Kapansýn mý?", vbYesNo + vbCritical)
        If cevap = vbYes Then
            winsck(i).SendData ("*CDROMKAPAT*")
            cmdcdrom.Caption = "CD Rom A"
        End If
    
    End If
End If
'------
Timer2.Enabled = True
End Sub

Private Sub cmddgoster_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If Not winsck(i).State <> sckConnected Then
    winsck(i).SendData ("*PGONDER*")
Else
    lstdurum.Clear
End If
'------

Timer2.Enabled = True

End Sub

Private Sub cmdkulac2_Click()
On Error Resume Next
i = txtmno
'------
    winsck(i).SendData ("*MASAAC*")
End Sub

Private Sub cmdkulkapat_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " kullanýma kapansýn mý?", vbYesNo)
        If cevap = vbYes Then
            winsck(i).SendData ("*MASAK*")
        End If
End If
'------

Timer2.Enabled = True
End Sub

Private Sub cmdgonder_Click()
On Error Resume Next
'****************************************
Timer2.Enabled = False
i = Val(txtmno)
If winsck(i).State <> sckConnected Then
MsgBox "Masa " & i & " kapalý !!!"
Else
txtekran.SelText = "<*> " & txtmsg + vbCrLf
winsck(i).SendData (txtmsg)
txtmsg = ""
End If
Timer2.Enabled = True
'****************************************
End Sub

Private Sub cmdkapat_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " kapatýlsýn mý?", vbYesNo + vbCritical)
        If cevap = vbYes Then
            winsck(i).SendData ("*KAPAT*")
        End If
End If
'------

Timer2.Enabled = True
End Sub

Private Sub cmdmasaustu_Click()
On Error Resume Next

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " Masa üstü görüntüsü alýnsýn mý?", vbYesNo + vbCritical)
        If cevap = vbYes Then
            winsck(i).SendData ("*MASAAL*")
        End If
End If
'------
End Sub

Private Sub cmdkulkapat2_Click()
On Error Resume Next
i = txtmno
'------
    winsck(i).SendData ("*MASAK*")
End Sub

Private Sub cmdprokapat_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " seçilen program kapatýlsýn mý?", vbYesNo)
        If cevap = vbYes Then
            winsck(i).SendData ("*PROKAPAT*" & lstdurum.ListIndex)
        End If
End If
'------
Timer2.Enabled = True

End Sub

Private Sub cmdsifregoster_Click()
On Error Resume Next
If cmdsifregoster.Caption = "Göster" Then
    txtsifre.PasswordChar = ""
    cmdsifregoster.Caption = "Gizle"
Else
    cmdsifregoster.Caption = "Göster"
    txtsifre.PasswordChar = "*"
End If
End Sub





Private Sub cmdsite_Click()
On Error Resume Next
MsgBox "BU BÖLÜM YAPIM AÞAMASINDADIR EN KISA ZAMANDA HÝZMETE GÝRECEKTÝR"
End Sub

Private Sub cmdtumkapat_Click()
On Error Resume Next
cevap = MsgBox("Tüm Bilgisayarlar kapatýlsýn mý?", vbYesNo + vbCritical)

If cevap = vbYes Then
Timer2.Enabled = False

If chkotodurum.Value = 1 Then
    chkotodurum.Value = 0
        For j = 1 To txtms
            optm(j).Value = True
                i = txtmno
                If Not winsck(i).State <> sckConnected Then
                    winsck(i).SendData ("*KAPAT*")
                End If
        Next j
    End If
    chkotodurum.Value = 1
Else
    For j = 1 To txtms
        optm(j).Value = True
            i = txtmno
            If Not winsck(i).State <> sckConnected Then
                winsck(i).SendData ("*KAPAT*")
            End If
    Next j
End If

Timer2.Enabled = True

End Sub

Private Sub cmdtumybaslat_Click()
On Error Resume Next
cevap = MsgBox("Tüm bilgisayarlar yeniden kapatýlsýn mý?", vbYesNo + vbCritical)

If cevap = vbYes Then
    Timer2.Enabled = False
    If chkotodurum.Value = 1 Then
        chkotodurum.Value = 0
        For j = 1 To txtms
            optm(j).Value = True
                i = txtmno
                If Not winsck(i).State <> sckConnected Then
                    winsck(i).SendData ("*YBASLAT*")
                End If
                
        Next j
        chkotodurum.Value = 1
    Else
        For j = 1 To txtms
             optm(j).Value = True
            i = txtmno
            If Not winsck(i).State <> sckConnected Then
              winsck(i).SendData ("*YBASLAT*")
             End If
        Next j
    End If
End If
Timer2.Enabled = True

End Sub

Private Sub cmdx_Click()
On Error Resume Next
frasifre.Visible = False
End Sub







Private Sub cmdybaslat_Click()
On Error Resume Next
Timer2.Enabled = False

i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    cevap = MsgBox("Masa " & i & " yeniden baþlatýlsýn mý?", vbYesNo + vbCritical)
        If cevap = vbYes Then
            winsck(i).SendData ("*YBASLAT*")
        End If
End If
'------

Timer2.Enabled = True
End Sub


Private Sub Form_Load()
On Error Resume Next

'********data iþlemleri**********************************
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstclient = dtkafe.OpenRecordset("client")
Set rstkafe = dtkafe.OpenRecordset("masalar")
Set rstucret = dtkafe.OpenRecordset("ucretler")
Set rstclientayar = dtkafe.OpenRecordset("clientayar")
Set rstuye = dtkafe.OpenRecordset("uyeler")
'*********************************************************
rstclient.MoveFirst
rstucret.MoveFirst

'***bazý sabitlerin datadan tüklenmesi****
txtad = winsck(1).LocalHostName
txtip = winsck(1).LocalIP
txtms = rstucret![msayisi]

If rstucret!hgonder = True Then chkhgonder.Value = 1
txthgonderdk = rstucret!hgonderdk
optm(1).Value = True
If rstucret!otodurum = True Then chkotodurum.Value = 1
'******************************************


For i = 1 To Val(txtms)
    '----görünüm-----------
    optm(i).Visible = True
    '----------------------
    '***tüm winsck'lar dinlemeye alýnýyor***
    winsck(i).LocalPort = rstclient!dport
    winsck(i).Listen
    rstclient.MoveNext
    '**************************************
Next i

'**********************************
Timer1.Interval = 2000
Timer2.Interval = Val(txthgonderdk) * 60000
Timer3.Interval = 1000
Timer3.Enabled = False
'**********************************

'********************
With rstucret
.MoveFirst
.Edit
!client = 1
.Update
End With
'********************

RENK_VER

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If fraayar.Visible = False Then
    cevap = MsgBox("Clientlerle baðlantýyý koparmak istiyor musunuz?", vbYesNo + vbCritical)
    If cevap = vbNo Then
    Cancel = True
    End If
    
    '*******masa aktif pasif yapmak için***********
    With rstucret
    .MoveFirst
    .Edit
    !client = 0
    .Update
    End With
    '**********************************************
Else
    cmdiptal_Click
    Cancel = True
End If
End Sub



Private Sub lstdurum_Click()
On Error Resume Next
lstdurum.ToolTipText = lstdurum.Text
End Sub

Private Sub optm_Click(Index As Integer)
On Error Resume Next
'*********data iþlemleri****************
'***************************************
With rstclient
.Index = "indexmasano"
.Seek "=", Index
If .NoMatch = False Then
txtmno = ![masano]
txtport = ![dport]
End If
End With
'***************************************
'***************************************
If chkotodurum.Value = 1 Then cmddgoster_Click
'***************************************
End Sub

Private Sub sldses_Change()
On Error Resume Next
i = txtmno
'------
If winsck(i).State <> sckConnected Then
    MsgBox "Masa " & i & " kapalý !!!"
Else
    winsck(i).SendData ("*SES*" & sldses.Value)
End If
'------


End Sub


Private Sub Timer1_Timer()
On Error Resume Next

rstclient.MoveFirst
For i = 1 To txtms
    
    If frmana.a(i) <> "" Then
    optm(i).BackColor = &H808080
    Else
    optm(i).BackColor = &H404040
    End If
    
    If winsck(i).State <> sckConnected Then
        winsck(i).Close
        winsck(i).LocalPort = rstclient!dport
        winsck(i).Listen
        optm(i).ForeColor = &HC0C0C0
    End If
rstclient.MoveNext
Next i
Timer2.Interval = Val(txthgonderdk) * 60000
End Sub



Private Sub Timer2_Timer()
On Error Resume Next
III = 0
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
III = III + 1
If Not III = txtms + 1 Then
    If chkotodurum.Value = 1 Then
        chkotodurum.Value = 0
        
        IMN = txtmno
        optm(III).Value = True
        cmdhgonder.Value = True
        optm(IMN).Value = True
        
        chkotodurum.Value = 1
    Else
        IMN = txtmno
        optm(III).Value = True
        cmdhgonder.Value = True
        optm(IMN).Value = True
    End If
Else
    Timer3.Enabled = False
End If
End Sub

Private Sub txthgonderdk_Change()
On Error Resume Next
If Val(txthgonderdk) < 10 Then
    txthgonderdk = "0" & Val(txthgonderdk)
End If

rstucret.MoveFirst
rstucret.Edit
 rstucret!hgonderdk = Val(txthgonderdk)
rstucret.Update
End Sub

Private Sub txtmsg_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgonder_Click
End Sub

Private Sub txtsifresor_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then cmdgiris_Click
End Sub

Private Sub updhgonderdk_Change()
On Error Resume Next
txthgonderdk = updhgonderdk.Value
End Sub

Private Sub winsck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
'****koþulsuz kabul et****
DoEvents
If winsck(Index).State <> sckClosed Then winsck(Index).Close
winsck(Index).Accept (requestID)
'*************************
End Sub

Private Sub winsck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
winsck(Index).GetData mesaj, vbString, bytesTotal

'hesabýn açýk olup olmadýðýnýn sorulmasý verilen cevap
If Left(mesaj, 6) = "*HAMI*" Then
    Timer2.Enabled = False
    If frmana.a(Index) <> "" Then
        winsck(Index).SendData ("*MASAAC*")
    End If
    Timer2.Enabled = Enabled
End If


'kullanýma açma isteði
If Left(mesaj, 5) = "*KAC*" Then
    Timer2.Enabled = False
    cevap = MsgBox("Masa " & Index & " kullanýma açma isteði !!!" & vbCrLf & "Kullanýma açýlmasýný istiyor musunuz?", vbYesNo + vbInformation)
    If cevap = vbYes Then
        frmana.cmdmasa(Index).Value = True
        If frmmasa.txtacilis = "" Then
            frmmasa.cmdhesapac.Value = True
        End If
        frmmasa.cmdcikis.Value = True
        
    End If
    Timer2.Enabled = Enabled
End If


'---çalýþan programlarý sýrala----
If Left(mesaj, 3) = "*P*" Then
    lstdurum.Clear
    bas = 1
    For i = 0 To 100
        bas1 = InStr(bas, mesaj, "*P*")
        son = InStr(bas1 + 1, mesaj, "*P*")
        lstdurum.AddItem (Mid(mesaj, bas1 + 3, son - bas1 - 3))
        bas = son
    Next i
End If

'baðlantý kontrolü
If mesaj = "*A*" Then optm(Index).ForeColor = vbGreen

'mesaj alýnmasý
If Left(mesaj, 3) = "*M*" Then txtekran.SelText = "<Masa " & Index & ">(" & Time & ") " & Mid(mesaj, 4) + vbCrLf

'üye kontrolü hesap açýlmasý ekran koruyucuda
If Left(mesaj, 3) = "*K*" Then
Timer2.Enabled = False

With rstuye
    If Mid(mesaj, 4) <> "" Then
        .MoveFirst
        For i = 1 To .RecordCount
            If Mid(mesaj, 4) = !AD & "~" & !sifre Then
                If !KONTOR > 0 Then
                    winsck(Index).SendData ("*O*" & !AD & "~" & !sifre & "~" & !KONTOR & "~" & Format(Time, "hh:mm"))
                    
                    frmana.cmdmasa(Index).Value = True
                    frmana.cmdmasa(Index).Picture = frmana.cmdresim.Picture
                    frmana.cmdmasa(Index).Caption = ""
                    frmmasa.cmdhesapac.Value = True
                    frmmasa.chkuye.Value = 1
                    frmmasa.cmdcikis.Value = True
                    
                    Exit For
                End If
            End If
            .MoveNext
        Next i
    End If
End With

Timer2.Enabled = True
End If

'üye kontrolü üyelik sisteminde
If Left(mesaj, 4) = "*KO*" Then
Timer2.Enabled = False

With rstuye
    If Mid(mesaj, 5) <> "" Then
        .MoveFirst
        For i = 1 To .RecordCount
            If Mid(mesaj, 5) = !AD & "~" & !sifre Then
                winsck(Index).SendData ("*OO*" & !KONTOR)
                Exit For
            End If
            .MoveNext
        Next i
    End If
End With

Timer2.Enabled = True
End If

'üye kontor saat kontrolü üyelik sisteminde
If Left(mesaj, 7) = "*SORGU*" Then
Timer2.Enabled = False

winsck(Index).SendData ("*S*" & frmana.a(Index) & "~" & frmana.s(Index))

Timer2.Enabled = True
End If

'üye þifre deðiþimi üyelik sisteminde
If Left(mesaj, 5) = "*SFR*" Then
Timer2.Enabled = False

With rstuye
    If Mid(mesaj, 5) <> "" Then
        .MoveFirst
        For i = 1 To .RecordCount
            Dim bul1, bul2
            bul1 = InStr(1, mesaj, "~")
            bul2 = InStr(bul1 + 1, mesaj, "~")
            
            If Mid(mesaj, 6, bul2 - 6) = !AD & "~" & !sifre Then
                winsck(Index).SendData ("*SFRD*")
                
                .Edit
                    !sifre = Mid(mesaj, bul2 + 1)
                .Update
                
                Exit For
            End If
            
            If i <> .RecordCount Then
                .MoveNext
            Else
                winsck(Index).SendData ("*SFRY*")
            End If
        Next i
    End If
End With
    
Timer2.Enabled = True
End If

'üye hesap kapatma üyelik sisteminde
If Left(mesaj, 4) = "*HK*" Then
Timer2.Enabled = False

With rstuye
    .MoveFirst
    For i = 1 To .RecordCount
        
        Dim bula1
        bula1 = InStr(1, mesaj, "~")
        
        'kullaným anýnda kontör biterse
        If Mid(mesaj, 5) = !AD & "-" & !sifre Then
            .Edit
                !KONTOR = 0
            .Update
            
                frmana.cmdmasa(Index).Value = True
                frmmasa.cmdhesapkapat.Value = True
                frmmasa.cmdcikis.Value = True
            Exit For
         End If
         
         'hesabýný kendi kapatýrsa
         If Mid(mesaj, 5, bula1 - 5) = !AD & "-" & !sifre Then
            .Edit
                !KONTOR = Mid(mesaj, bula1 + 1)
            .Update
            
                frmana.cmdmasa(Index).Value = True
                frmmasa.cmdhesapkapat.Value = True
                frmmasa.cmdcikis.Value = True
            Exit For
         End If
         
        .MoveNext
    Next i
End With
    
Timer2.Enabled = True
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
    If TypeOf C Is Frame Then C.BackColor = !arkarenk
Next
End With
'****************************************************************************

fraayar.BackColor = vbRed
Label14.ForeColor = vbWhite

End Sub



