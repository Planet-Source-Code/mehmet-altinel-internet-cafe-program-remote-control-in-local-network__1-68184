VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmversiyon 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Versiyon Kontrolü::."
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frmversiyon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraversiyon 
      BackColor       =   &H00FFE7E3&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtgbilgi 
         Appearance      =   0  'Flat
         Height          =   3135
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   3135
      End
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
         TabIndex        =   1
         Top             =   3360
         Width           =   1815
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   0
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versiyon bilgisine ulaþýlamadý internet baðlantýnýzý kontrol ediniz."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmversiyon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstucret As Recordset
Dim rstnot As Recordset
Dim rstuyar As Recordset
Dim rstkafe As Recordset
Dim dtkafe As Database

Private Sub GUNCELLE()
On Error Resume Next
Dim Version As String, News As String
Dim Site As String

Site = rstucret!versite
Version = Inet1.OpenURL(Site & "Versiyon.txt")
    
Dim Uzunluk
Uzunluk = Len(Version)
    
'if not olabilir
If Not Uzunluk = 0 And Not Uzunluk > 10 Then 'eðer versiyon bilgisine ulaþýlamýyorsa yada 404 hata sayfasý geliyorsa güncelleme iptal edilir.
    If Not Trim(Version) = "kilit" Then 'her ihtimale karþý kilitleme durumlarýnda
        If Trim(Version) > App.Major & "." & App.Minor & "." & App.Revision Then
            fraversiyon.Visible = True
            Me.Height = fraversiyon.Height
            txtgbilgi = Replace(Inet1.OpenURL(Site & "Yenilikler.txt"), Chr(10), vbCrLf)
        Else
            Unload Me
            MsgBox "Þuanda En Son Versiyonu Kullanýyorsunuz.", vbInformation
        End If
    Else
        MsgBox "Programýnýz Program Sahibi Tarafýndan Kilitlenmiþtir..."
        End
    End If
End If

End Sub

Private Sub cmdvyukselt_Click()
On Error Resume Next
cevap = MsgBox("Yeni versiyonu yüklemek istiyor musunuz?" + vbCrLf + "Not: Programýnýzýn veri tabanýnýn yedeðini alýnýz. (datakafe.mdb)", vbYesNo + vbInformation)
If cevap = vbYes Then
    Shell App.Path & "\ÖKH Güncelle.exe", vbNormalFocus 'güncellemeyi yapacak olan program çalýþtýrýlýyor
    End                                           'program kapatýlýyor ki güncellensin
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstucret = dtkafe.OpenRecordset("ucretler")

fraversiyon.Move 0, 0

RENK_VER
GUNCELLE

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
End Sub
