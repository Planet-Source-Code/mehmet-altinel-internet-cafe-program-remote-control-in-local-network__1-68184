VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdtrans 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Veri Transferi::."
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frmdtrans.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
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
      Height          =   4695
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmdtrans.frx":144A
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdtransfer 
      BackColor       =   &H006CFBD3&
      Caption         =   "VerileriTransfer Et"
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
      MouseIcon       =   "frmdtrans.frx":15DF
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4950
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4762
            MinWidth        =   4762
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
            TextSave        =   "11:02"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Mehmet ALTINEL & Türker ÖZER"
            TextSave        =   "Mehmet ALTINEL & Türker ÖZER"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4560
      X2              =   4560
      Y1              =   0
      Y2              =   4920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmdtrans.frx":18E9
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
      Height          =   1215
      Left            =   0
      TabIndex        =   7
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label lbldurum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kaynak Seçiniz (datakafe.mdb)"
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
      Height          =   420
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label lblkaynak 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D:\ökh en 1.5.0 (02.01.2005)\ÖKH 1.5.0\server\datakafeE.mdb"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
End
Attribute VB_Name = "frmdtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************HEDEF**************************
Dim dtkafe As Database

Dim rstclient As Recordset
Dim rstclientayar As Recordset
Dim rstkasa As Recordset
Dim rstmailler As Recordset
Dim rstmasalar As Recordset
Dim rstmrapor As Recordset
Dim rstmusteriler As Recordset
Dim rstnotlar As Recordset
Dim rstraporlar As Recordset
Dim rstucretler As Recordset
Dim rstuyailar As Recordset
Dim rstuyeler As Recordset
'**********************************************************

'***************************KAYNAK*************************
Dim dtkafe2 As Database

Dim rstclient2 As Recordset
Dim rstclientayar2 As Recordset
Dim rstkasa2 As Recordset
Dim rstmailler2 As Recordset
Dim rstmasalar2 As Recordset
Dim rstmrapor2 As Recordset
Dim rstmusteriler2 As Recordset
Dim rstnotlar2 As Recordset
Dim rstraporlar2 As Recordset
Dim rstucretler2 As Recordset
Dim rstuyailar2 As Recordset
Dim rstuyeler2 As Recordset
'***********************************************************

Private Sub cmdtransfer_Click()
On Error Resume Next

If Right(lblkaynak, 12) = "datakafe.mdb" Then

'***************************KAYNAK********************************
Set dtkafe2 = OpenDatabase(lblkaynak)

Set rstclient2 = dtkafe2.OpenRecordset("client")
Set rstclientayar2 = dtkafe2.OpenRecordset("clientayar")
Set rstkasa2 = dtkafe2.OpenRecordset("kasa")
Set rstmailler2 = dtkafe2.OpenRecordset("mailler")
Set rstmasalar2 = dtkafe2.OpenRecordset("masalar")
Set rstmrapor2 = dtkafe2.OpenRecordset("mrapor")
Set rstmusteriler2 = dtkafe2.OpenRecordset("musteriler")
Set rstnotlar2 = dtkafe2.OpenRecordset("notlar")
Set rstraporlar2 = dtkafe2.OpenRecordset("raporlar")
Set rstucretler2 = dtkafe2.OpenRecordset("ucretler")
Set rstuyarilar2 = dtkafe2.OpenRecordset("uyarilar")
Set rstuyeler2 = dtkafe.OpenRecordset("uyeler")
'******************************************************************

lbldurum = "Transfer ediliyor..."
Timer1.Interval = 1

'anaformun timerý durduruluyor
frmana.Timer1.Interval = 0
frmmasa.Timer1.Interval = 0

With rstclient2
    rstclient.MoveFirst
    .MoveFirst
    For i = 1 To .RecordCount
        rstclient.Edit
            rstclient!ip = !ip
            rstclient!dport = !dport
        .MoveNext
        rstclient.Update
        rstclient.MoveNext
    Next i
End With

With rstclientayar2
    rstclientayar.MoveFirst
    .MoveFirst
    For i = 1 To .RecordCount
        rstclientayar.Edit
            rstclientayar!sifre = !sifre
            rstclientayar!ekran = !ekran
            rstclientayar!ek = !ek
            rstclientayar!kafeadi = !kafeadi
            rstclientayar!chat = !chat
            rstclientayar!eyaz = !eyaz
            rstclientayar!hesap = !hesap
            rstclientayar!flash = !flash
            rstclientayar!kontor = !kontor
            rstclientayar!arayuz = !arayuz
        .MoveNext
        rstclientayar.Update
        rstclientayar.MoveNext
    Next i
End With

With rstkasa2
    rstkasa.MoveFirst
    .MoveFirst
    For i = 1 To .RecordCount
        rstkasa.Edit
            rstkasa!kira = !kira
            rstkasa!elektrik = !elektrik
            rstkasa!su = !su
            rstkasa!telefon = !telefon
            rstkasa!internet = !internet
            rstkasa!internet = !internet
            rstkasa!eleman = !eleman
            rstkasa!diger = !diger
            rstkasa!tkira = !tkira
            rstkasa!telektrik = !telektrik
            rstkasa!tsu = !tsu
            rstkasa!ttelefon = !ttelefon
            rstkasa!tinternet = !tinternet
            rstkasa!tinternet = !tinternet
            rstkasa!teleman = !teleman
            rstkasa!ekgelir = !ekgelir
            rstkasa!ay = !ay
            rstkasa!egaciklama = !egaciklama
            rstkasa!ekbolum = !ekbolum
        .MoveNext
        rstkasa.Update
        rstkasa.MoveNext
    Next i
End With

With rstmailler2
    rstmailler.MoveFirst
    For j = 1 To rstmailler.RecordCount
        rstmailler.Delete
        rstmailler.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstmailler.AddNew
            rstmailler!adres = !adres
            rstmailler!hmail = !hmail
            rstmailler!aID = !aID
            rstmailler!Password = !Password
        .MoveNext
        rstmailler.Update
    Next i
End With

With rstmasalar2
    rstmasalar.MoveFirst
    .MoveFirst
    For i = 1 To .RecordCount
        rstmasalar.Edit
            rstmasalar!acilis1 = !acilis1
            rstmasalar!ucret1 = !ucret1
            rstmasalar!sure1 = !sure1
            rstmasalar!sucret = !sucret
            rstmasalar!ssure = !ssure
            rstmasalar!eucret = !eucret
            rstmasalar!kod = !kod
            rstmasalar!notssure = !notsucret
            rstmasalar!masanot = !masanot
            rstmasalar!secucret2 = !secucret2
            rstmasalar!masaad = !masaad
            rstmasalar!aciklama = !aciklama
            rstmasalar!uye = !uye
        .MoveNext
        rstmasalar.Update
        rstmasalar.MoveNext
    Next i
End With

With rstmrapor2
    rstmrapor.MoveFirst
    For j = 1 To rstmrapor.RecordCount
        rstmrapor.Delete
        rstmrapor.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstmrapor.AddNew
            rstmrapor!ad = !ad
            rstmrapor!tarih = !tarih
            rstmrapor!islem = !islem
            rstmrapor!miktar = !miktar
            rstmrapor!ekbolum = !ekbolum
        .MoveNext
        rstmrapor.Update
    Next i
End With

With rstmusteriler2
    rstmusteriler.MoveFirst
    For j = 1 To rstmusteriler.RecordCount
        rstmusteriler.Delete
        rstmusteriler.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstmusteriler.AddNew
            rstmusteriler!ad = !ad
            rstmusteriler!tel = !tel
            rstmusteriler!adres = !adres
            rstmusteriler!aciklama = !aciklama
            rstmusteriler!borc = !borc
            rstmusteriler!sonodeme = !sonodeme
            rstmusteriler!sotarih = !sotarih
        .MoveNext
        rstmusteriler.Update
    Next i
End With

With rstnotlar2
    rstnotlar.MoveFirst
    For j = 1 To rstnotlar.RecordCount
        rstnotlar.Delete
        rstnotlar.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstnotlar.AddNew
            rstnotlar!Not = !Not
            rstnotlar!tarih = !tarih
            rstnotlar!saat = !saat
            rstnotlar!baslik = !baslik
        .MoveNext
        rstnotlar.Update
    Next i
End With

With rstraporlar2
    rstraporlar.MoveFirst
    For j = 1 To rstraporlar.RecordCount
        rstraporlar.Delete
        rstraporlar.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount - 1
        If rstraporlar2!acilis <> "" Then
            rstraporlar.AddNew
                rstraporlar!acilis = !acilis
                rstraporlar!sure = !sure
                rstraporlar!ucret = !ucret
                rstraporlar!tarih = !tarih
                rstraporlar!masa = !masa
                rstraporlar!bitis = !bitis
                rstraporlar!ekbolum = !ekbolum
                rstraporlar!aciklama = !aciklama
            .MoveNext
            rstraporlar.Update
        End If
    Next i
'eðer boþ rapor giriþi varsa silinecek
rstraporlar.MoveFirst
For i = 1 To rstraporlar.RecordCount
    If rstraporlar!acilis = "" Then
        rstraporlar.Delete
    End If
    rstraporlar.MoveNext
Next i

End With

With rstucretler2
    rstucretler.MoveFirst
    .MoveFirst
    For i = 1 To .RecordCount
        rstucretler.Edit
            rstucretler!ucret = !ucret
            rstucretler!basucret = !basucret
            rstucretler!birim = !birim
            rstucretler!eu1 = !eu1
            rstucretler!eu2 = !eu2
            rstucretler!eu3 = !eu3
            rstucretler!eu4 = !eu4
            rstucretler!eu5 = !eu5
            rstucretler!eu6 = !eu6
            rstucretler!eu7 = !eu7
            rstucretler!euisim1 = !euisim1
            rstucretler!euisim2 = !euisim2
            rstucretler!euisim3 = !euisim3
            rstucretler!euisim4 = !euisim4
            rstucretler!euisim5 = !euisim5
            rstucretler!euisim6 = !euisim6
            rstucretler!euisim7 = !euisim7
            rstucretler!msayisi = !msayisi
            rstucretler!konrapor = !konrapor
            rstucretler!konkasa = !konkasa
            rstucretler!konanarapor = !konanarapor
            rstucretler!songuncelleme = !songuncelleme
            rstucretler!webadresi = !webadresi
            rstucretler!anasayfa = !anasayfa
            rstucretler!hgonder = !hgonder
            rstucretler!hgonderdk = !hgonderdk
            rstucretler!parabirimi = !parabirimi
            rstucretler!sifre = !sifre
            rstucretler!konasaat = !konasaat
            rstucretler!ucret2 = !ucret2
            rstucretler!hesapac = !hesapac
            rstucretler!client = !client
            rstucretler!yedek = !yedek
            rstucretler!otodurum = !otodurum
            rstucretler!otobaglan = !otobaglan
            rstucretler!arkarenk = !arkarenk
            rstucretler!onrenk = !onrenk
            rstucretler!tusrenk = !tusrenk
            rstucretler!vrenk = !vrenk
            rstucretler!yyuvarla = !yyuvarla
            rstucretler!konuye = !konuye
            rstucretler!otokapat = !otokapat
        .MoveNext
        rstucretler.Update
        rstucretler.MoveNext
    Next i
End With

With rstuyarilar2
    rstruyarilar.MoveFirst
    For j = 1 To rstuyarilar.RecordCount
        rstuyarilar.Delete
        rstuyarilar.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstuyarilar.AddNew
            rstuyarilar!utarih = !utarih
            rstuyarilar!usaat = !usaat
            rstuyarilar!unot = !unot
            rstuyarilar!ubaslik = !ubaslik
        .MoveNext
        rstuyarilar.Update
    Next i
End With

With rstuyeler2
    rstruyeler.MoveFirst
    For j = 1 To rstuyeler.RecordCount
        rstuyeler.Delete
        rstuyeler.MoveNext
    Next j
    .MoveFirst
    For i = 1 To .RecordCount
        rstuyeler.AddNew
            rstuyeler!adsoyad = !adsoyad
            rstuyeler!tel = !tel
            rstuyeler!aciklama = !aciklama
            rstuyeler!ad = !ad
            rstuyeler!sifre = !sifre
            rstuyeler!kontor = !kontor
            rstuyeler!fiyat = !fiyat
            rstuyeler!tarih = !tarih
            rstuyeler!topfiyat = !topfiyat
        .MoveNext
        rstuyeler.Update
    Next i
End With

End

Else
    lbldurum = "Geçersiz dosya biçimi seçtiniz !"
End If
End Sub



Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dir1.ToolTipText = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
On Error Resume Next

lblkaynak = File1.Path & "\" & File1.FileName
If Mid(lblkaynak, InStr(1, lblkaynak, "\") + 1, 1) = "\" Then
    lblkaynak = File1.Path & File1.FileName
End If

lbldurum = "Verileri Transfer Et tuþuna basýn"
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
File1.ToolTipText = File1.FileName
End Sub

Private Sub Form_Load()
On Error Resume Next
'***************************HEDEF*********************************
Set dtkafe = OpenDatabase(App.Path & "\datakafe")

Set rstclient = dtkafe.OpenRecordset("client")
Set rstclientayar = dtkafe.OpenRecordset("clientayar")
Set rstkasa = dtkafe.OpenRecordset("kasa")
Set rstmailler = dtkafe.OpenRecordset("mailler")
Set rstmasalar = dtkafe.OpenRecordset("masalar")
Set rstmrapor = dtkafe.OpenRecordset("mrapor")
Set rstmusteriler = dtkafe.OpenRecordset("musteriler")
Set rstnotlar = dtkafe.OpenRecordset("notlar")
Set rstraporlar = dtkafe.OpenRecordset("raporlar")
Set rstucretler = dtkafe.OpenRecordset("ucretler")
Set rstuyarilar = dtkafe.OpenRecordset("uyarilar")
Set rstuyeler = dtkafe.OpenRecordset("uyeler")
'*****************************************************************
lblkaynak = ""

RENK_VER

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 1

If ProgressBar1.Value = ProgressBar1.Max Then
    lbldurum = "Transfer baþarýyla tamamlandý"
    ProgressBar1.Value = 0
    Timer1.Interval = 0
    End
End If

End Sub
Private Sub RENK_VER()
On Error Resume Next
'renk deðiþimi**************************************************************

With rstucretler
.MoveFirst
Me.BackColor = !arkarenk

Dim C
For Each C In Me.Controls
    If TypeOf C Is CommandButton Then C.BackColor = !tusrenk
    If TypeOf C Is CheckBox Then C.BackColor = !arkarenk
    If TypeOf C Is CheckBox Then C.ForeColor = !onrenk
    If TypeOf C Is OptionButton Then C.BackColor = !arkarenk
    If TypeOf C Is OptionButton Then C.ForeColor = !onrenk
    If TypeOf C Is Frame Then C.BackColor = !arkarenk
Next
End With
'****************************************************************************
End Sub
