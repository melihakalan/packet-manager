Attribute VB_Name = "Module3"
Public PresetFile(1 To 99) As String
Public PresetName(1 To 99) As String

Public Function LoadPreset(FileName As String)
frmMain.lstPacket.Clear

If Dir(App.Path & "/Presets/" & FileName) <> "" Then

Dim sFileText As String
Dim iFileNo As Integer
iFileNo = FreeFile
'open the file for reading
Open App.Path & "/Presets/" & FileName For Input As #iFileNo
'change this filename to an existing file!  (or run the example below first)

'read the file until we reach the end
Do While Not EOF(iFileNo)
Input #iFileNo, sFileText

frmMain.lstPacket.AddItem sFileText

Loop

'close the file (if you dont do this, you wont be able to open it again!)
Close #iFileNo
Else
MsgBox "Could not find the preset file (" & FileName & ")", vbExclamation, "Error"
End If

End Function
