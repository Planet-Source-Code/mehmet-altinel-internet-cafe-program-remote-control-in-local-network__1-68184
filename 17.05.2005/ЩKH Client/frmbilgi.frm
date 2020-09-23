VERSION 5.00
Begin VB.Form frmbilgi 
   BackColor       =   &H00AC7222&
   BorderStyle     =   0  'None
   ClientHeight    =   615
   ClientLeft      =   12690
   ClientTop       =   135
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblucret 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblmasa 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "MASA 50"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmbilgi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbclient As Database
Dim rstclient As Recordset
Dim i As Long
Private Sub Form_Load()
On Error Resume Next
App.TaskVisible = False

Me.Top = 0
Me.Left = Screen.Width - Me.Width

lblmasa = "MASA " & Mid(frmclient.txtbport, 4, 2)

Set dbclient = OpenDatabase(App.Path & "\dataclient.mdb")
Set rstclient = dbclient.OpenRecordset("client")
rstclient.MoveFirst

'Timer1.Interval = 1000
'Timer2.Interval = 15 * 1000
Timer3.Interval = 1000
i = 1
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Me.Top = 0
Me.Left = Screen.Width - Me.Width

If rstclient!sinir = 1 Then
    i = i + 1
    If i / 2 = i \ 2 Then
        Me.BackColor = vbRed
        
    Else
        Me.BackColor = vbBlack
    End If
Else
    Me.BackColor = vbBlack
End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
rstclient.MoveFirst
rstclient.Edit
    rstclient!sinir = "0"
rstclient.Update
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Me.Top = 0
Me.Left = Screen.Width - Me.Width
End Sub
