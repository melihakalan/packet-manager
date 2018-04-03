VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "O K"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   495
      TabIndex        =   2
      Top             =   630
      Width           =   1230
   End
   Begin VB.Shape shpClose 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   440
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.snoxd.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   440
      TabIndex        =   1
      Top             =   360
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KoJD"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   920
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   955
      Left            =   20
      Top             =   15
      Width           =   2130
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = vbWhite
shpClose.BorderColor = vbWhite
shpClose.FillColor = vbBlack
End Sub

Private Sub lblClose_Click()
frmAbout.Hide
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = vbBlack
shpClose.BorderColor = vbRed
shpClose.FillColor = vbWhite
End Sub
