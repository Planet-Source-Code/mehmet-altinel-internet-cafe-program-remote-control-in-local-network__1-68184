VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmnot 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Notlar - Uyarýlar::."
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmnot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdteksil 
      BackColor       =   &H006CFBD3&
      Caption         =   "Seçili Uyarýyý Sil"
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
      Left            =   4560
      MouseIcon       =   "frmnot.frx":144A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ListBox lstuyar 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "frmnot.frx":1754
      Left            =   4560
      List            =   "frmnot.frx":1756
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdlistele 
      BackColor       =   &H006CFBD3&
      Caption         =   "Uyarýlar >"
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
      Left            =   3480
      MouseIcon       =   "frmnot.frx":1758
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtbaslik 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox txtnot 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   840
      Width           =   4335
   End
   Begin VB.CommandButton cmduyar 
      BackColor       =   &H006CFBD3&
      Caption         =   "Uyarý Ekle"
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
      Left            =   2280
      MouseIcon       =   "frmnot.frx":1B9A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox cmbtarih 
      Appearance      =   0  'Flat
      BackColor       =   &H00D5F9FF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdson 
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
      Height          =   375
      Left            =   4080
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdon 
      BackColor       =   &H006CFBD3&
      Caption         =   "<"
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
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdsil 
      BackColor       =   &H006CFBD3&
      Caption         =   "Sil"
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
      Left            =   1200
      MouseIcon       =   "frmnot.frx":1EA4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdkaydet 
      BackColor       =   &H006CFBD3&
      Caption         =   "Yeni Not"
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
      MouseIcon       =   "frmnot.frx":21AE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4755
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "ÖZER KAFE HESAP"
            TextSave        =   "ÖZER KAFE HESAP"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "31.01.2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "22:15"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24510465
      CurrentDate     =   38240
   End
   Begin VB.TextBox txthh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "00"
      Top             =   4320
      Width           =   375
   End
   Begin MSComCtl2.UpDown updhh 
      Height          =   300
      Left            =   1920
      TabIndex        =   15
      Top             =   4320
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   24
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtmm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "00"
      Top             =   4320
      Width           =   375
   End
   Begin MSComCtl2.UpDown updmm 
      Height          =   300
      Left            =   2520
      TabIndex        =   16
      Top             =   4320
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Max             =   60
      Min             =   -1
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdtumsil 
      BackColor       =   &H006CFBD3&
      Caption         =   "Tüm Uyarýlarý Sil"
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
      Left            =   4560
      MouseIcon       =   "frmnot.frx":24B8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Uyarýlar"
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
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Tarih-Saat"
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
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BackStyle       =   0  'Transparent
      Caption         =   "      Tarih                Saat"
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
      TabIndex        =   19
      Top             =   4080
      Width           =   2655
   End
End
Attribute VB_Name = "frmnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstuyar As Recordset
Dim rstnot As Recordset
Dim rstucret As Recordset
Dim dtnot As Database


Private Sub cmbtarih_Click()
On Error Resume Next
rstnot.Index = "indextarih"
rstnot.Seek "=", cmbtarih.Text
txtbaslik = rstnot!baslik
txtnot = rstnot![not]
cmbtarih.Text = rstnot![tarih]
End Sub

Private Sub cmdkaydet_Click()
On Error Resume Next
'***
If cmdkaydet.Caption = "Yeni Not" Then
    '***
    cmdkaydet.Caption = "Kaydet"
    cmdsil.Caption = "Ýptal"
    txtnot.Locked = False
    txtbaslik.Locked = False
    txtnot = ""
    txtbaslik = ""
    txtbaslik.SetFocus
    cmbtarih.Text = Date & " - " & Time
    '***
    cmdon.Enabled = False
    cmdson.Enabled = False
    cmduyar.Enabled = False
    Me.Height = 4815
    '***
Else
    '***
    cmdkaydet.Caption = "Yeni Not"
    cmdsil.Caption = "Sil"
    txtnot.Locked = True
    txtnot.Locked = True
    cmdkaydet.SetFocus
    '***
    cmdon.Enabled = True
    cmdson.Enabled = True
    cmduyar.Enabled = True
    '***
    If cmbtarih.Text <> "" Then
        '---
        rstnot.AddNew
        rstnot![not] = txtnot
        rstnot![tarih] = cmbtarih
        rstnot!baslik = txtbaslik
        rstnot.Update
        '***
        cmbtarih.Clear
        rstnot.MoveFirst
        Do Until rstnot.EOF
            cmbtarih.AddItem (rstnot![tarih])
            rstnot.MoveNext
        Loop
        '***
        rstnot.Requery
        cmbtarih.AddItem (cmbtarih.Text)
        End If
    '***
End If
'***
End Sub

Private Sub cmdlistele_Click()
On Error Resume Next

If cmdlistele.Caption = "Uyarýlar >" Then
    cmdlistele.Caption = "Uyarýlar <"
    Me.Width = 6780
    cmbtarih.Locked = True
    
    cmdkaydet.Enabled = False
    cmdsil.Enabled = False
    cmdon.Enabled = False
    cmdson.Enabled = False
    '-------listele--------
    lstuyar.Clear
    rstuyar.MoveFirst
    For i = 1 To rstuyar.RecordCount
        lstuyar.AddItem (rstuyar!ubaslik)
        rstuyar.MoveNext
    Next i
    lstuyar.ListIndex = 0
Else
    cmdlistele.Caption = "Uyarýlar >"
    
    cmbtarih.Locked = False
    cmdkaydet.Enabled = True
    cmdsil.Enabled = True
    cmdon.Enabled = True
    cmdson.Enabled = True
    Me.Width = 4635
End If

End Sub

Private Sub cmdon_Click()
On Error Resume Next
'------------------------
If Not Me.Height = 5085 Then
    If Not rstnot.BOF Then
        rstnot.MovePrevious
        txtnot = rstnot![not]
        txtbaslik = rstnot!baslik
        cmbtarih.Text = rstnot![tarih]
    End If
Else
    If Not rstuyar.BOF Then
        rstuyar.MovePrevious
        txtnot = rstnot![not]
        txtbaslik = rstbaslik
        cmbtarih.Text = rstnot![tarih]
    End If
End If

End Sub

Private Sub cmdsil_Click()
On Error Resume Next
'****
If cmdsil.Caption = "Sil" Then
    If txtbaslik <> "" Then
        cevap = MsgBox("Bu Notu silmek istiyor musunuz?", vbCritical + vbYesNo)
        If cevap = vbYes Then
            rstnot.Delete
            txtnot = ""
            txtbaslik = ""
            cmdson_Click
        End If
    End If
Else
    '***
    rstnot.MoveFirst
    txtnot = rstnot![not]
    txtbaslik = rstnot!baslik
    cmbtarih.Text = rstnot![tarih]
    '***
    cmdkaydet.Caption = "Yeni Not"
    cmdsil.Caption = "Sil"
    txtnot.Locked = True
    cmdkaydet.SetFocus
    '***
    cmdon.Enabled = True
    cmdson.Enabled = True
    cmduyar.Enabled = True
    '***
End If
'****
cmbtarih.Clear
rstnot.MoveFirst
Do Until rstnot.EOF
    cmbtarih.AddItem (rstnot![tarih])
    rstnot.MoveNext
Loop
'****
End Sub

Private Sub cmdson_Click()
On Error Resume Next
If Not Me.Height = 5085 Then
    If Not rstnot.EOF Then
        rstnot.MoveNext
        txtnot = rstnot![not]
        txtbaslik = rstnot!baslik
        cmbtarih.Text = rstnot![tarih]
    End If
Else
    If Not rstuyar.EOF Then
            rstuyar.MovePrevious
            txtnot = rstnot![not]
            txtbaslik = rstbaslik
            cmbtarih.Text = rstnot![tarih]
    End If
End If
End Sub

Private Sub cmduyar_Click()
On Error Resume Next
'***
If cmduyar.Caption = "Uyarý Ekle" Then
    cmduyar.Caption = "Uyarý Kaydet"
    frmnot.Height = 5475
    '***
    txtnot.Locked = False
    txtbaslik.Locked = False
    txtnot = ""
    txtbaslik = ""
    txtbaslik.SetFocus
    cmbtarih.Text = Date & " - " & Time
    '***
    txtmm = "00"
    txthh = "00"
    DTPicker1.Value = Date
    '***
Else
    cmduyar.Caption = "Uyarý Ekle"
    frmnot.Height = 4815
    '***
    cmdkaydet.SetFocus
    '***
    If txtbaslik <> "" Then
        rstuyar.AddNew
        rstuyar![unot] = txtnot
        rstuyar!ubaslik = txtbaslik
        rstuyar![utarih] = DTPicker1.Value
        rstuyar![usaat] = txthh & ":" & txtmm & ":" & "00"
        rstuyar!ubaslik = txtbaslik
        rstuyar.Update
        
        cmdlistele_Click
        cmdlistele_Click
        
    End If
    
    txtnot.Locked = True
    txtbaslik.Locked = True
    
End If
'***
End Sub

Private Sub cmdteksil_Click()
On Error Resume Next
If lstuyar.Text <> "" Then
    cevap = MsgBox("Uyarý silinsin mi?", vbInformation + vbYesNo)
    If cevap = vbYes Then
        rstuyar.MoveFirst
        For i = 1 To rstuyar.RecordCount
            If lstuyar.Text = rstuyar!ubaslik Then
                txtbaslik = rstuyar!ubaslik
                rstuyar.Delete
                
                txtbaslik = ""
                txtnot = ""
                cmbtarih = ""
                
                cmdlistele_Click
                cmdlistele_Click
            End If
            rstuyar.MoveNext
        Next i
    End If
End If
End Sub

Private Sub cmdtumsil_Click()
On Error Resume Next
SS = lstuyar.ListCount
rstuyar.MoveFirst
cevap = MsgBox("Tüm uyarýlarý silmek istiyormusunuz?", vbYesNo + vbQuestion)
If cevap = vbYes Then
    For i = 0 To rstuyar.RecordCount
        rstuyar.Delete
        rstuyar.MoveNext
    Next i
    
    lstuyar.Clear
    txtbaslik = ""
    txtnot = ""
    cmbtarih = ""
    
    MsgBox "Bu güne kadar oluþturulan toplam " & SS & " uyarý silinmiþtir", vbInformation
End If
End Sub

Private Sub cmduyarýsil_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
On Error Resume Next
Set dtnot = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstnot = dtnot.OpenRecordset("notlar")
Set rstuyar = dtnot.OpenRecordset("uyarilar")
Set rstucret = dtnot.OpenRecordset("ucretler")
'-----
Me.Width = 4650
Me.Height = 4815
'-----
rstnot.MoveFirst
Do Until rstnot.EOF
    cmbtarih.AddItem (rstnot![tarih])
    rstnot.MoveNext
Loop
'--
rstnot.MoveFirst
txtnot = rstnot![not]
cmbtarih.Text = rstnot![tarih]
txtbaslik = rstnot!baslik
'***
cmdkaydet.SetFocus

RENK_VER

End Sub



Private Sub lstuyar_Click()
On Error Resume Next
rstuyar.MoveFirst
For i = 1 To rstuyar.RecordCount
    If lstuyar.Text = rstuyar!ubaslik Then
        txtbaslik = rstuyar!ubaslik
        txtnot = rstuyar!unot
        cmbtarih = rstuyar!utarih & "-" & rstuyar!usaat
    End If
    rstuyar.MoveNext
Next i

End Sub

Private Sub updhh_Change()
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
End Sub

Private Sub updmm_Change()
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
End Sub
