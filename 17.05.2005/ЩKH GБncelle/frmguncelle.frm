VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmguncelle 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::ÖZER KAFE HESAP GÜNCELLEME::."
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmguncelle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   840
   End
   Begin MSComctlLib.ProgressBar prgislem 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdbaslat 
      BackColor       =   &H006CFBD3&
      Caption         =   "Baþlat"
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1350
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "ÖKH GUNCELLEME"
            TextSave        =   "ÖKH GUNCELLEME"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "13.03.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "12:56"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbldurum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Durum..."
      ForeColor       =   &H006CFBD3&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmguncelle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbkafe As Database
Dim rstucret As Recordset

Private Sub cmdbaslat_Click()
On Error Resume Next
Set dbkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dbkafe.OpenRecordset("ucretler")
rstucret.MoveFirst

Dim Site As String, Program As String 'site adresi ve programýn ismi tanýmlamýyor
Dim Mx() As Byte 'Mx tanýmlanýyor
'Dim Mx2() As Byte

cmdbaslat.Enabled = False

'Site = "http://freehost04.websamba.com/depomed/okh/Özer Kafe Hesap.exe"
Program = "Özer Kafe Hesap.exe"
Site = rstucret!versite & Program


lbldurum = "Güncelleme Baþladý..."
Timer1.Interval = 10

'exe download ediliyor
Mx() = Inet1.OpenURL(Site, 1) 'Adres açýlýyor...
Open App.Path & "\" & Program For Binary Access Write As #1 'Etkin dizine belirtilen isim ve uzantýda dosya oluþturuluyor...
Put #1, , Mx() 'Dosya kaydediliyor...
Close #1 '#1 Kapatýlýyor... Güncelleme iþlemimiz bitti....

'data download ediliyor
'Mx2() = Inet1.OpenURL(Site, 1) 'Adres açýlýyor...
'Open App.Path & "\" & "datakafe.mdb" For Binary Access Write As #1
'Put #1, , Mx2() 'Dosya kaydediliyor...
'Close #1 ' #1 Kapatýlýyor... Güncelleme iþlemimiz bitti....


Timer1.Interval = 0
lbldurum = "Tamamlandý..."
prgislem.Value = 100
cmdbaslat.Enabled = True

dbkafe.Close
Shell (App.Path & "\Özer Kafe Hesap.exe")
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Inet1.Cancel
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If prgislem.Value < 100 Then
        prgislem.Value = prgislem.Value + 1
    Else
        prgislem.Value = 0
    End If
End Sub


