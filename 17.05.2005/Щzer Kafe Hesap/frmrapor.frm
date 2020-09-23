VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrapor 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Rapor::."
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmrapor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frasifre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   5520
      TabIndex        =   15
      Top             =   6480
      Width           =   7395
      Begin VB.Frame Frame1 
         BackColor       =   &H00BFA3C9&
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   2400
         TabIndex        =   16
         Top             =   2280
         Width           =   2535
         Begin VB.CommandButton cmdgiris 
            BackColor       =   &H006CFBD3&
            Caption         =   "*GÝRÝÞ* )>)>)>"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtsifre 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   600
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   0
            Width           =   2535
         End
      End
   End
   Begin VB.Frame fraduzenle 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      BorderStyle     =   0  'None
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
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox txtaciklama 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   30
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtucret 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtsure 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtbitis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtacilis 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtmasa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdiptal 
         BackColor       =   &H006CFBD3&
         Caption         =   "Ýptal"
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
         Left            =   3600
         MouseIcon       =   "frmrapor.frx":144A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmddegistir 
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
         Left            =   2400
         MouseIcon       =   "frmrapor.frx":1754
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   7320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DÜZENLE"
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
         Left            =   3000
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   5
      ItemData        =   "frmrapor.frx":1A5E
      Left            =   4440
      List            =   "frmrapor.frx":1A60
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   4
      ItemData        =   "frmrapor.frx":1A62
      Left            =   3360
      List            =   "frmrapor.frx":1A64
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   3
      ItemData        =   "frmrapor.frx":1A66
      Left            =   2640
      List            =   "frmrapor.frx":1A68
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   2
      ItemData        =   "frmrapor.frx":1A6A
      Left            =   1920
      List            =   "frmrapor.frx":1A6C
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   1
      ItemData        =   "frmrapor.frx":1A6E
      Left            =   1200
      List            =   "frmrapor.frx":1A70
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox lsttarih 
      Height          =   255
      ItemData        =   "frmrapor.frx":1A72
      Left            =   6720
      List            =   "frmrapor.frx":1A74
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdlistele 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6435
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   9701
            MinWidth        =   9701
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18.04.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1552
            MinWidth        =   1552
            TextSave        =   "10:45"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "frmrapor.frx":1A76
      Left            =   2040
      List            =   "frmrapor.frx":1A9E
      TabIndex        =   1
      Text            =   "Ay Seçiniz"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdsil 
      BackColor       =   &H006CFBD3&
      Caption         =   "Raporu Sil"
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
      MouseIcon       =   "frmrapor.frx":1B1D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   1095
   End
   Begin VB.ListBox lstrapor 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00C00000&
      Height          =   4710
      Index           =   0
      ItemData        =   "frmrapor.frx":1E27
      Left            =   120
      List            =   "frmrapor.frx":1E29
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txttoplam 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   6000
      Width           =   3255
   End
   Begin VB.ComboBox cmbtarih 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     Masa       Açýlýþ     Bitiþ      Süre         Ücret                    Açýklama"
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
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GENEL TOPLAM"
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
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      BackStyle       =   0  'Transparent
      Caption         =   "Günlük Rapor               Aylýk Rapor"
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
      TabIndex        =   10
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu mnurapor 
      Caption         =   "Rapor"
      Visible         =   0   'False
      Begin VB.Menu mnusil 
         Caption         =   "Sil"
      End
      Begin VB.Menu mnuduzenle 
         Caption         =   "Düzenle"
      End
   End
End
Attribute VB_Name = "frmrapor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstrapor As Recordset
Dim rstucret As Recordset
Dim dtkafe As Database

Private Sub cmbtarih_Click()
On Error Resume Next

For i = 0 To 5
lstrapor(i).Clear
Next i

rstrapor.MoveFirst
For i = 1 To rstrapor.RecordCount
    If rstrapor!tarih = cmbtarih Then
        lstrapor(0).AddItem rstrapor![masa]
        lstrapor(1).AddItem rstrapor![acilis]
        lstrapor(2).AddItem rstrapor![bitis]
        lstrapor(3).AddItem rstrapor![sure]
        lstrapor(4).AddItem rstrapor![ucret]
        lstrapor(5).AddItem rstrapor![aciklama]
    End If
    rstrapor.MoveNext
Next

rstliste.MoveFirst
txttoplam = 0
If rstucret!parabirimi = 0 Then
    rstrapor.MoveFirst
    For i = 1 To rstrapor.RecordCount
        If rstrapor!tarih = cmbtarih Then
            txttoplam = CLng(txttoplam) + CLng(rstrapor![ucret]) & " TL"
            txttoplam = Format(txttoplam, "#,00")
        End If
        rstrapor.MoveNext
    Next i
Else
    rstrapor.MoveFirst
    For i = 1 To rstrapor.RecordCount
        If rstrapor!tarih = cmbtarih Then
            txttoplam = CDbl(txttoplam) + CDbl(rstrapor![ucret]) & " TL"
            txttoplam = Format(txttoplam, "#,00.00")
        End If
        rstrapor.MoveNext
    Next i
End If
'---------------------------------------------
cmbay.Text = "Ay Seçiniz"
End Sub
Private Sub cmbay_Click()
On Error Resume Next

For i = 0 To 5
lstrapor(i).Clear
Next i

rstrapor.MoveFirst
For i = 1 To rstrapor.RecordCount
    If Not rstrapor!tarih = "" Then
        If Val(Mid(cmbay.Text, 1, 2)) = Val(Mid(rstrapor![tarih], 4, 2)) Then
        lstrapor(0).AddItem rstrapor![masa]
        lstrapor(1).AddItem rstrapor![acilis]
        lstrapor(2).AddItem rstrapor![bitis]
        lstrapor(3).AddItem rstrapor![sure]
        lstrapor(4).AddItem rstrapor![ucret]
        lstrapor(5).AddItem rstrapor![aciklama]
        End If
    End If
    rstrapor.MoveNext
Next i


rstrapor.MoveFirst
txttoplam = 0
cmbtarih.Text = ""
If rstucret!parabirimi = 0 Then
    rstrapor.MoveFirst
    For i = 1 To rstrapor.RecordCount
        If Val(Mid(cmbay.Text, 1, 2)) = Val(Mid(rstrapor![tarih], 4, 2)) Then
            txttoplam = CLng(txttoplam) + CLng(rstrapor![ucret])
            txttoplam = Format(txttoplam, "#,00")
        End If
    rstrapor.MoveNext
    Next i
Else
    rstrapor.MoveFirst
    For i = 1 To rstrapor.RecordCount
        If Val(Mid(cmbay.Text, 1, 2)) = Val(Mid(rstrapor![tarih], 4, 2)) Then
            txttoplam = CDbl(txttoplam) + CDbl(rstrapor![ucret])
            txttoplam = Format(txttoplam, "#,00.00")
        End If
        rstrapor.MoveNext
    Next i
End If

If lstrapor(4).List(0) = "" Then
    txttoplam = 0
End If

End Sub

Private Sub cmddegistir_Click()
On Error Resume Next
Bindex = cmbtarih.ListIndex
With rstrapor
.Edit
    !masa = txtmasa
    !acilis = txtacilis
    !bitis = txtbitis
    !sure = txtsure
    !ucret = txtucret
    !aciklama = txtaciklama
.Update

fraduzenle.Visible = False


For a = 0 To 5
    lstrapor(a).Enabled = True
Next a

cmbay.Enabled = True
cmbtarih.Enabled = True

cmbtarih.ListIndex = 0
cmbtarih.ListIndex = Bindex

End With
End Sub

Private Sub cmdgiris_Click()
On Error Resume Next

If txtsifre = rstucret!sifre Then
    frasifre.Visible = False
    cmdlistele_Click
Else
    MsgBox "Yanlýþ Þifre Girdiniz!!!", vbCritical
End If

End Sub



Private Sub cmdiptal_Click()
On Error Resume Next
cmbay.Enabled = True
cmbtarih.Enabled = True
fraduzenle.Visible = False

For a = 0 To 5
    lstrapor(a).Enabled = True
Next a
End Sub

Private Sub cmdlistele_Click()
'On Error Resume Next
rstrapor.MoveFirst
For i = 1 To rstrapor.RecordCount - 6
    If Not rstrapor![tarih] = "" Then
        lsttarih.AddItem (rstrapor![tarih])
    End If
    rstrapor.MoveNext
Next i
    
'Dim i As Integer
'Dim j As Integer
'Dim var As Boolean

For i = 1 To lsttarih.ListCount
    Var = False
        For j = 1 To cmbtarih.ListCount + 1
            If lsttarih.List(i) = cmbtarih.List(j - 1) Then
                Var = True
                Exit For
            End If
        Next j
    If Var = False Then
        cmbtarih.AddItem lsttarih.List(i)
    End If
Next i

cmbtarih.Text = cmbtarih.List(cmbtarih.ListCount - 1)
cmbtarih_Click

End Sub

Private Sub cmdsil_Click()
On Error Resume Next
If cmbtarih.Text <> "" Then
    '---------günlük rapor sil----------------
    cevap = MsgBox(cmbtarih.Text & " Tarihli Raporu silmek istiyor musunuz?" & vbCrLf & "NOT: Eðer Raporu silerseniz kasadan hesap düþülecektir(Tavsiye edilmez)", vbCritical + vbYesNo)
    If cevap = vbYes Then
        rstrapor.MoveFirst
        For i = 1 To rstrapor.RecordCount
            If rstrapor!tarih = cmbtarih Then
                rstrapor.Delete
            End If
            rstrapor.MoveNext
        Next i
        
        For i = 0 To 5
        lstrapor(i).Clear
        Next i
        
        cmbtarih.RemoveItem (cmbtarih.ListIndex)
        txttoplam = 0
    End If
Else
    '-----------aylýk rapor sil------------
    cevap = MsgBox(cmbay.Text & " Ay'ýnýn Raporunu silmek istiyor musunuz?", vbCritical + vbYesNo)
    If cevap = vbYes Then
        For i = 0 To 5
        lstrapor(i).Clear
        Next i
        
        rstrapor.MoveFirst
        For i = 1 To rstrapor.RecordCount
            If Val(Mid(rstrapor!tarih, 4, 2)) = Val(Mid(cmbay, 1, 2)) Then
                rstrapor.Delete
            End If
        rstrapor.MoveNext
        Next i
        txttoplam = 0
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'---
Me.Width = 7395
frasifre.Move 0, 0
'---
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstrapor = dtkafe.OpenRecordset("raporlar")
Set rstucret = dtkafe.OpenRecordset("ucretler")
'********
rstucret.MoveFirst
'------
If rstucret!konanarapor = "1" Then
    frasifre = True
Else
    frasifre.Visible = False
    cmdlistele_Click
End If

RENK_VER

End Sub

Private Sub lstrapor_Click(Index As Integer)
On Error Resume Next
For i = 0 To 5
    lstrapor(i).ListIndex = lstrapor(Index).ListIndex
Next i
lstrapor(Index).ToolTipText = lstrapor(Index).Text
End Sub

Private Sub lstrapor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If cmbay.Text = "Ay Seçiniz" Then
    If Button = 2 Then
        PopupMenu mnurapor
    End If
End If
End Sub

Private Sub mnuduzenle_Click()
On Error Resume Next
fraduzenle.Visible = True
cmbtarih.Enabled = False
cmbay.Enabled = False

'temizlik
txtmasa = ""
txtacilis = ""
txtbitis = ""
txtsure = ""
txtucret = ""
txtaciklama = ""
            
With rstrapor
.MoveFirst
    For i = 1 To .RecordCount
        If !tarih = CDate(cmbtarih.Text) And !masa = lstrapor(0).Text And !acilis = lstrapor(1) And !bitis = lstrapor(2) Then
            txtmasa = !masa
            txtacilis = !acilis
            txtbitis = !bitis
            txtsure = !sure
            txtucret = !ucret
            txtaciklama = !aciklama
            
            For a = 0 To 5
                lstrapor(a).Enabled = False
            Next a
            
            Exit For
        Else
        
        End If
        .MoveNext
    Next i
End With

End Sub

Private Sub mnusil_Click()
On Error Resume Next
cevap = MsgBox("Bu kaydý silmek istiyor musunuz?", vbInformation + vbYesNo)
If cevap = vbYes Then
    Aindex = cmbtarih.ListIndex
    With rstrapor
    .MoveFirst
        For i = 1 To .RecordCount
            If !tarih = CDate(cmbtarih.Text) And !masa = lstrapor(0).Text And !acilis = lstrapor(1) And !bitis = lstrapor(2) Then
                .Delete
                
                cmbtarih.ListIndex = 0
                cmbtarih.ListIndex = Aindex
                
                Exit For
            Else
            
            End If
            .MoveNext
        Next i
    End With
End If
End Sub

Private Sub txtsifre_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdgiris_Click
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
Label5.ForeColor = vbWhite

End Sub


