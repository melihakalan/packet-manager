VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KoJD PKT Manager 1745"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   2790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frHack 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   30
      TabIndex        =   3
      Top             =   410
      Width           =   2775
      Begin VB.Timer tmPacket 
         Enabled         =   0   'False
         Left            =   1800
         Top             =   120
      End
      Begin VB.Timer TimerOfTheKojd 
         Interval        =   50
         Left            =   2160
         Top             =   120
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "1000"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CheckBox btnTimer 
         Caption         =   "Start Timer"
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton btnSendAll 
         Caption         =   "Send All Now"
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtPacket 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "Send This"
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.ListBox lstPacket 
         Height          =   1425
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press SpaceBar to put item in List."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   1950
         Width           =   2445
      End
      Begin VB.Shape shpHelp 
         BackColor       =   &H00000000&
         BorderColor     =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label LabelOfTheKOJD 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "KoJD (Snoxd.net)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   11
         Top             =   2640
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Int(ms)"
         Height          =   195
         Left            =   1245
         TabIndex        =   9
         Top             =   2310
         Width           =   525
      End
      Begin VB.Shape shpHack 
         BorderColor     =   &H00800000&
         Height          =   2925
         Left            =   0
         Top             =   0
         Width           =   2730
      End
   End
   Begin VB.Frame frAttach 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   30
      TabIndex        =   0
      Top             =   -80
      Width           =   2775
      Begin VB.CommandButton btnStart 
         Caption         =   "Start"
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   160
         Width           =   975
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Text            =   "Knight OnLine Client"
         Top             =   160
         Width           =   1575
      End
      Begin VB.Shape shpAttach 
         BorderColor     =   &H00800000&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2730
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "Option"
      Visible         =   0   'False
      Begin VB.Menu mnuSend 
         Caption         =   "Send This"
      End
      Begin VB.Menu mnuSendAll 
         Caption         =   "Send All"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "Drop This"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear List"
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuLog 
         Caption         =   "View Log"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTopPresets 
      Caption         =   "Presets"
      Begin VB.Menu mnuTopPreset 
         Caption         =   "Not Used"
         Index           =   0
      End
   End
   Begin VB.Menu mnuGoAbout 
      Caption         =   "KoJD"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hwnd As Long, ByVal ptx As Long, ByVal pty As Long, ByVal bAutoScroll As Long) As Long

Public Coloring
Public PresetCount As Integer



Private Sub btnSend_Click()
On Error Resume Next
Dim a
a = Replace(txtPacket.Text, " ", "")
txtPacket.Text = a
If txtPacket.Text <> "" Then
Dim pBytes() As Byte
Dim pStr As String
pStr = txtPacket.Text
frmLog.lstLog.AddItem Time & " | " & pStr
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End If
End Sub

Private Sub btnSendAll_Click()
Dim pBytes() As Byte
Dim pStr As String
Dim I As Integer
lstPacket.ListIndex = 0

For I = 0 To lstPacket.ListCount - 1

pStr = lstPacket.Text

If Not lstPacket.ListIndex = lstPacket.ListCount - 1 Then
lstPacket.ListIndex = lstPacket.ListIndex + 1
Else
lstPacket.ListIndex = 0
End If
frmLog.lstLog.AddItem Time & " | " & pStr
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
Next

End Sub

Private Sub btnStart_Click()
KO_TITLE = txtTitle.Text
LoadOffsets
If AttachKO = False Then
Exit Sub
End If
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
KO_ADR_DLG = ReadLong(KO_PTR_DLG)
frHack.Enabled = True
frAttach.Enabled = False
End Sub

Private Sub btnTimer_Click()
If btnTimer.Value = 1 Then
tmPacket.Interval = CInt(txtInterval.Text)
tmPacket.Enabled = True
Else
tmPacket.Enabled = False
End If
End Sub


Private Sub Form_Load()
frHack.Enabled = False
Load frmLog
frmLog.Hide
lblHelp.Visible = False
shpHelp.Visible = False
LoadIni
LoadPresets
LoadMenus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Killapp "PacketManager.exe"
End Sub

Private Sub LabelOfTheKOJD_Click()
Shell "explorer http://www.snoxd.net/"
End Sub

Private Sub lstPacket_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = vbRightButton Then
PopupMenu mnuList

    Dim aPt As POINTAPI
    Dim anInd As Long
    GetCursorPos aPt
    With lstPacket
        .ListIndex = LBItemFromPt(.hwnd, aPt.X, aPt.Y, False)
    End With

End If
End Sub

Private Sub mnuBDW_Click()
lstPacket.Clear
lstPacket.AddItem "5F080100"
lstPacket.AddItem "5F080200"
lstPacket.AddItem "5F080300"
lstPacket.AddItem "5F080400"
lstPacket.AddItem "5F080500"
lstPacket.AddItem "5F080600"
lstPacket.AddItem "5F080700"
lstPacket.AddItem "5F080800"
lstPacket.AddItem "5F080900"
lstPacket.AddItem "5F080A00"
End Sub

Private Sub mnuClear_Click()
lstPacket.Clear
End Sub

Private Sub mnuDrop_Click()
On Error Resume Next
lstPacket.RemoveItem (lstPacket.ListIndex)
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuGoAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuLog_Click()
frmLog.Show
End Sub

Private Sub mnuSend_Click()
On Error Resume Next
Dim pBytes() As Byte
Dim pStr As String
pStr = lstPacket.Text
frmLog.lstLog.AddItem Time & " | " & pStr
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
End Sub

Private Sub mnuSendAll_Click()
On Error Resume Next
If Not lstPacket.ListCount = 0 Then
Dim pBytes() As Byte
Dim pStr As String
Dim I As Integer
lstPacket.ListIndex = 0

For I = 0 To lstPacket.ListCount - 1

pStr = lstPacket.Text

If Not lstPacket.ListIndex = lstPacket.ListCount - 1 Then
lstPacket.ListIndex = lstPacket.ListIndex + 1
Else
lstPacket.ListIndex = 0
End If
frmLog.lstLog.AddItem Time & " | " & pStr
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
Next
End If
End Sub

Private Sub mnuTopPreset_Click(Index As Integer)

LoadPreset (PresetFile(Index))

End Sub

Private Sub TimerOfTheKojd_Timer()
Coloring = Coloring + 1

Select Case Coloring
Case 1
LabelOfTheKOJD.ForeColor = &HE0E0E0
LabelOfTheKOJD.BackColor = &H404040
Case 2
LabelOfTheKOJD.ForeColor = &HC0C0C0
LabelOfTheKOJD.BackColor = &H808080
Case 3
LabelOfTheKOJD.ForeColor = &H808080
LabelOfTheKOJD.BackColor = &HC0C0C0
Case 4
LabelOfTheKOJD.ForeColor = &H404040
LabelOfTheKOJD.BackColor = &HE0E0E0
Case 5
LabelOfTheKOJD.ForeColor = &H0
LabelOfTheKOJD.BackColor = &HFFFFFF
Case 6
LabelOfTheKOJD.ForeColor = &H404040
LabelOfTheKOJD.BackColor = &HE0E0E0
Case 7
LabelOfTheKOJD.ForeColor = &H808080
LabelOfTheKOJD.BackColor = &HC0C0C0
Case 8
LabelOfTheKOJD.ForeColor = &HC0C0C0
LabelOfTheKOJD.BackColor = &H808080
Case 9
LabelOfTheKOJD.ForeColor = &HE0E0E0
LabelOfTheKOJD.BackColor = &H404040
Case 10
LabelOfTheKOJD.ForeColor = &HFFFFFF
LabelOfTheKOJD.BackColor = &H0
Coloring = 0
End Select
End Sub

Private Sub tmPacket_Timer()
Dim pBytes() As Byte
Dim pStr As String
Dim I As Integer
lstPacket.ListIndex = 0

For I = 0 To lstPacket.ListCount - 1

pStr = lstPacket.Text

If Not lstPacket.ListIndex = lstPacket.ListCount - 1 Then
lstPacket.ListIndex = lstPacket.ListIndex + 1
Else
lstPacket.ListIndex = 0
End If
frmLog.lstLog.AddItem Time & " | " & pStr
ConvHEX2ByteArray pStr, pBytes
SendPackets pBytes()
Next
End Sub

Private Sub txtPacket_Change()
Call ShowHelp
End Sub

Private Sub txtPacket_Click()
btnSend.Default = True
End Sub

Private Sub txtPacket_KeyDown(KeyCode As Integer, Shift As Integer)
Call ShowHelp
If KeyCode = 32 Then
Dim a
a = Replace(txtPacket.Text, " ", "")
If Not txtPacket.Text = "" And Not txtPacket.Text = " " And Not Len(txtPacket.Text) < 4 Then
lstPacket.AddItem a
End If
txtPacket.Text = ""
Call ShowHelp
End If
End Sub

Private Sub ShowHelp()

If Not txtPacket.Text = "" And Not txtPacket.Text = " " Then
lblHelp.Visible = True
shpHelp.Visible = True
btnSendAll.Visible = False
Else
lblHelp.Visible = False
shpHelp.Visible = False
btnSendAll.Visible = True
End If
End Sub

Public Sub LoadIni()
On Error GoTo IniError

If Dir(App.Path & "/Presets.ini") <> "" Then
PresetCount = ReadIni(App.Path & "/Presets.ini", "PRESETS", "PRESETCOUNT")

Else
MsgBox "Could not find Presets.ini, Application will be closed.", vbExclamation, "Error"
End
End If
Exit Sub
IniError:
MsgBox "Loading settings from the .ini file has failed. Presets.ini might be corrupted.", vbExclamation, "Error"
End
End Sub

Public Sub LoadPresets()
On Error GoTo PresetError

Dim X As Integer

For X = 1 To PresetCount

If X < 10 Then
PresetFile(X) = ReadIni(App.Path & "/Presets.ini", "LOADPRESETS", "PRESETFILE_0" & X)
PresetName(X) = ReadIni(App.Path & "/Presets.ini", "LOADPRESETS", "PRESETNAME_0" & X)
Else
PresetFile(X) = ReadIni(App.Path & "/Presets.ini", "LOADPRESETS", "PRESETFILE_" & X)
PresetName(X) = ReadIni(App.Path & "/Presets.ini", "LOADPRESETS", "PRESETNAME_" & X)
End If
Next

Exit Sub
PresetError:
MsgBox "Loading settings from the .ini file has failed. Presets.ini might be corrupted.", vbExclamation, "Error"
End

End Sub

Public Sub LoadMenus()
On Error GoTo menuerr

Dim X As Integer

For X = 1 To PresetCount
Load mnuTopPreset(X)
mnuTopPreset(X).Caption = PresetName(X)
Next
mnuTopPreset(0).Visible = False
Exit Sub
menuerr:
MsgBox "Could not load the menu.", vbExclamation, "Error"
End
End Sub
