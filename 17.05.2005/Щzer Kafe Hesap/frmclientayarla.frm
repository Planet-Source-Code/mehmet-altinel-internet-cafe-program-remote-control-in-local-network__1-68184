VERSION 5.00
Begin VB.Form frmclientayarla 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Client Ayarla::."
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "frmclientayarla.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtsifre 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   10
         ToolTipText     =   "CLIENT PROGRAMINA ��FRE VER�N AYARLARA S�ZDEN BA�KASI ULA�AMASIN"
         Top             =   120
         Width           =   2775
      End
      Begin VB.CheckBox chkekran 
         BackColor       =   &H00404080&
         Caption         =   "A��l��ta Ekran Koruyucu Gelsin"
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
         TabIndex        =   4
         ToolTipText     =   "WINDOWS A�ILI�INDA EKRANI KULLANIMA KAPATMAK �ST�YORSANIZ SE��N"
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkchat 
         BackColor       =   &H00404080&
         Caption         =   "Server ' la G�r��meye �zin ver"
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
         TabIndex        =   3
         ToolTipText     =   "B�LG�SAYAR KULLANIMA KAPALIYKEN M��TER�N�N SERVERLA  MESAJLA�MASINI �ST�YORSANIZ SE��N"
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox txtkafeadi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   ".::�zer �nternet Kafe::."
         ToolTipText     =   "KAFEN�Z�N �SM�N� G�R�N"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtek 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmclientayarla.frx":144A
         ToolTipText     =   "EKLEMEK �STED�KLER�N�Z EKRAN KORUYUCUDA G�R�N�R"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "�ifre"
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
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Kafenizin �smi G�r�ns�n "
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
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Sizin Eklemek �stedikleriniz"
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
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label cmdkaydet 
         Alignment       =   2  'Center
         BackColor       =   &H006CFBD3&
         Caption         =   "Ayarlar� Kaydet"
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
         Left            =   120
         MouseIcon       =   "frmclientayarla.frx":1486
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label cmdiptal 
         Alignment       =   2  'Center
         BackColor       =   &H006CFBD3&
         Caption         =   "�ptal"
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
         Left            =   1560
         MouseIcon       =   "frmclientayarla.frx":26F8
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2760
         Width           =   975
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
         Left            =   2640
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmclientayarla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

