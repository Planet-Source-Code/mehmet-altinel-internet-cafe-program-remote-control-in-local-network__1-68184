VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmsohbet 
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".::Browser::."
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmsohbet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   8355
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H006CFBD3&
      Caption         =   "Git"
      Height          =   300
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdsil1 
      BackColor       =   &H006CFBD3&
      Caption         =   "SÝL"
      Height          =   465
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   825
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6615
      Left            =   0
      TabIndex        =   18
      Top             =   1560
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   11668
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
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.CommandButton cmdanasayfa 
      BackColor       =   &H006CFBD3&
      Caption         =   "Ana Sayfam yap"
      Height          =   300
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   9600
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9600
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdgir 
      BackColor       =   &H006CFBD3&
      Caption         =   "ID ve Password GÝR"
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdyenile 
      BackColor       =   &H006CFBD3&
      Caption         =   "YENÝLE"
      Height          =   615
      Left            =   2400
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "YENÝLE"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmddur 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      Caption         =   "DUR"
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "DUR"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdgit 
      BackColor       =   &H006CFBD3&
      Caption         =   ">>GÝT >>"
      Height          =   330
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "GÝT"
      Top             =   600
      Width           =   930
   End
   Begin VB.CommandButton cmdmkaydet 
      BackColor       =   &H006CFBD3&
      Caption         =   "KAYDET"
      Height          =   465
      Left            =   3360
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "@ MAÝL KAYDET"
      Top             =   120
      Width           =   825
   End
   Begin VB.CommandButton cmdtumsil 
      Appearance      =   0  'Flat
      BackColor       =   &H006CFBD3&
      Caption         =   "TÜM. SÝL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5040
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "GEÇMÝÞÝ SÝL"
      Top             =   120
      Width           =   930
   End
   Begin VB.CommandButton cmdson 
      BackColor       =   &H006CFBD3&
      Height          =   615
      Left            =   1680
      Picture         =   "frmsohbet.frx":144A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "ÝLERÝ"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdon 
      BackColor       =   &H006CFBD3&
      Height          =   615
      Left            =   120
      Picture         =   "frmsohbet.frx":169B
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "GERÝ"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbhmail 
      BackColor       =   &H00D5F9FF&
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cmbadres 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   8415
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   8280
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3969
            MinWidth        =   3969
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
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "09:19"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   3325
            MinWidth        =   3325
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsayfaadi 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADRES"
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
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   8640
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   8640
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmsohbet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstmail As Recordset
Dim rstucret As Recordset
Dim dtkafe As Database
Dim AA

Private Sub cmbadres_Change()
ProgressBar1.Visible = True
End Sub

Private Sub cmbadres_GotFocus()
On Error Resume Next
AA = 0
End Sub

Private Sub cmbadres_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim CtrlDown
    CtrlDown = (Shift And vbCtrlMask) > 0
    If CtrlDown = True Then
        If KeyCode = 13 Then
            cmbadres.Text = "www." & cmbadres.Text & ".com"
            cmdGO_Click
        End If
    End If
End Sub

Private Sub cmbadres_KeyPress(KeyAscii As Integer)
On Error Resume Next
'***
If KeyAscii = 13 Then
    cmdGO_Click
End If
'***
End Sub

Private Sub cmbadres_LostFocus()
On Error Resume Next
AA = 1
End Sub

Private Sub cmdanasayfa_Click()
On Error Resume Next

rstucret.MoveFirst
rstucret.Edit
rstucret!anasayfa = cmbadres.Text
rstucret.Update

End Sub

Private Sub cmddur_Click()
On Error Resume Next
WebBrowser1.Stop
End Sub

Private Sub cmdgir_Click()
    Print txtid
    SendKeys vbTab
    Print txtpassword
    SendKeys vbEnter
End Sub

Private Sub cmdgit_Click()
On Error Resume Next
'***
rstmail.Index = "indexhmail"
rstmail.Seek "=", cmbhmail.Text
cmbadres.Text = rstmail![adres]
'***
WebBrowser1.Navigate2 cmbadres.Text
cmbadres.AddItem cmbadres.Text
cmdon.Enabled = True
cmdson.Enabled = True
'***
End Sub

Private Sub cmdGO_Click()
On Error Resume Next
WebBrowser1.Navigate2 cmbadres.Text
cmbadres.AddItem cmbadres.Text
cmdon.Enabled = True
cmdson.Enabled = True
End Sub

Private Sub cmdmkaydet_Click()
On Error Resume Next
'***
If cmbhmail.Text <> "" And cmbadres.Text <> "" Then
    '***
    rstmail.Index = "indexhmail"
    rstmail.Seek "=", cmbhmail.Text
    If rstmail.NoMatch Then
        '***
        rstmail.AddNew
        rstmail![adres] = cmbadres.Text
        rstmail![hmail] = cmbhmail.Text
        rstmail![aID] = txtid
        rstmail![Password] = txtpassword
        rstmail.Update
        '***
        rstmail.MoveFirst
        cmbhmail.Clear
        cmbhmail.Text = rsthmail!hmail
        
        Do Until rstmail.EOF
            cmbhmail.AddItem (rstmail![hmail])
            rstmail.MoveNext
        Loop
        '***
    Else: MsgBox "Bu kayýt zaten var!!!", vbInformation
    End If
Else: MsgBox "Boþluklarý doldurun!!!", vbInformation
End If
'***
End Sub

Private Sub cmdon_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub cmdsil1_Click()
On Error Resume Next
'-------
If cmbhmail <> "" Then
    rstmail.Index = "indexhmail"
    rstmail.Seek "=", cmbhmail.Text
        If Not rstmail.NoMatch Then
            rstmail.Delete
            
            rstmail.MoveFirst
            cmbhmail.Clear
            cmbhmail.Text = rsthmail!hmail

            Do Until rstmail.EOF
                cmbhmail.AddItem (rstmail![hmail])
                rstmail.MoveNext
            Loop
        End If
Else
    MsgBox "Silinecek kayýt seçilmedi"
End If

End Sub

Private Sub cmdtumsil_Click()
On Error Resume Next
'***
cevap = MsgBox("Tüm kayýtlarý silmek istiyor musunuz?", vbCritical + vbYesNo)
If cevap = vbYes Then
    '---
    rstmail.MoveFirst
    Do Until rstmail.EOF
        rstmail.Delete
        rstmail.MoveNext
    Loop
    
    cmbhmail.Clear
    cmbadres.Clear
    txtid = ""
    txtpassword = ""
    '---
End If
'***
End Sub

Private Sub cmdson_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub cmdyenile_Click()
On Error Resume Next
WebBrowser1.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next
'***
AA = 1
Set dtkafe = OpenDatabase(App.Path & "\datakafe.mdb")
Set rstmail = dtkafe.OpenRecordset("mailler")
Set rstucret = dtkafe.OpenRecordset("ucretler")
'***
rstucret.MoveFirst
rstmail.MoveFirst
cmbhmail.Text = rsthmail![hmail]
'---
Do Until rstmail.EOF
    cmbhmail.AddItem (rstmail![hmail])
    rstmail.MoveNext
Loop
'***
'WebBrowser1.Navigate "http://members.lycos.co.uk/ozerinternetkafe/Sohbet"
WebBrowser1.Navigate rstucret!anasayfa
'***

RENK_VER

End Sub

Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Height = Me.Height - (9000 - 6615)
WebBrowser1.Width = Me.Width - (12000 - 11895)
ProgressBar1.Top = Me.Height - 655
ProgressBar1.Left = 3960
End Sub


Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
StatusBar1.Panels(5).Text = " Siteye baðlanýyor...   " & URL
MousePointer = vbHourglass
End Sub

Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
StatusBar1.Panels(5).Text = "  Yüklendi..."
MousePointer = vbDefault
lblsayfaadi = WebBrowser1.LocationName

ProgressBar1.Value = 0
ProgressBar1.Visible = False

End Sub

Private Sub WebBrowser1_GotFocus()
On Error Resume Next
cmbadres.Text = cmbadres.Text
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
ProgressBar1.Value = (Progress * 100) / ProgressMax
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
On Error Resume Next
If AA = 1 Then
    cmbadres.Text = WebBrowser1.LocationURL
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
End Sub

