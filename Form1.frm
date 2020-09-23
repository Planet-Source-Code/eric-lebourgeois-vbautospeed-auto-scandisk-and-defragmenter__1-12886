VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbAutoSpeed"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   2925
   Begin VB.CommandButton Command10 
      Caption         =   "Combo[A]"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Combo Box Manipulation Alpha Test"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Listbox[A]"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Listbox Manipulation Alpha Test"
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   1200
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   2250
      ScaleWidth      =   1500
      TabIndex        =   8
      ToolTipText     =   "Version Info"
      Top             =   960
      Width           =   1560
   End
   Begin VB.CommandButton Command6 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Active Options"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Start Scan"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Scan Options"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Log Options"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select OS"
      Height          =   375
      Left            =   1125
      TabIndex        =   3
      ToolTipText     =   "Operating System Menu"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Log"
      Height          =   375
      Left            =   70
      TabIndex        =   2
      ToolTipText     =   "Show Log"
      Top             =   360
      Width           =   975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drive:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   35
      TabIndex        =   1
      Top             =   40
      Width           =   1095
   End
   Begin VB.Menu selos 
      Caption         =   "Select OS"
      Visible         =   0   'False
      Begin VB.Menu win98 
         Caption         =   "Windows 98"
         Checked         =   -1  'True
      End
      Begin VB.Menu winme 
         Caption         =   "Windows ME"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu logopt 
      Caption         =   "Log Options"
      Visible         =   0   'False
      Begin VB.Menu savelogonexit 
         Caption         =   "Save Log on Exit"
         Checked         =   -1  'True
      End
      Begin VB.Menu displogonstartup 
         Caption         =   "Display Log on Startup"
      End
   End
   Begin VB.Menu mainopt 
      Caption         =   "vbAutoSpeed Options"
      Visible         =   0   'False
      Begin VB.Menu scandisk 
         Caption         =   "Scandisk"
         Checked         =   -1  'True
      End
      Begin VB.Menu defrag 
         Caption         =   "Defragment"
         Checked         =   -1  'True
      End
      Begin VB.Menu SWD 
         Caption         =   "Save Options When Done"
         Checked         =   -1  'True
      End
      Begin VB.Menu XWD 
         Caption         =   "Exit When Done"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu miscop 
      Caption         =   "Miscellaneous Options"
      Visible         =   0   'False
      Begin VB.Menu scanopt 
         Caption         =   "Scandisk Options"
         Begin VB.Menu stoss 
            Caption         =   "Standard Scan"
            Checked         =   -1  'True
         End
         Begin VB.Menu stots 
            Caption         =   "Thorough Scan"
         End
         Begin VB.Menu spacer4 
            Caption         =   "-"
         End
         Begin VB.Menu sto1 
            Caption         =   "System and Data areas"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu sto2 
            Caption         =   "System areas only"
            Enabled         =   0   'False
         End
         Begin VB.Menu sto3 
            Caption         =   "Data area only"
            Enabled         =   0   'False
         End
         Begin VB.Menu spacer 
            Caption         =   "-"
         End
         Begin VB.Menu sto4 
            Caption         =   "Do not perform write testing"
            Enabled         =   0   'False
         End
         Begin VB.Menu sto5 
            Caption         =   "Do not repair bad sectors in hidden or system files"
            Enabled         =   0   'False
         End
         Begin VB.Menu autofixe 
            Caption         =   "Automatically fix errors"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu defropt 
         Caption         =   "Defragment Options"
         Enabled         =   0   'False
         Begin VB.Menu dso1 
            Caption         =   "Rearrange program files"
            Checked         =   -1  'True
         End
         Begin VB.Menu dso2 
            Caption         =   "Check the drive for errors"
            Checked         =   -1  'True
         End
         Begin VB.Menu spacer2 
            Caption         =   "-"
         End
         Begin VB.Menu dso3 
            Caption         =   "Use these settings this time only"
         End
         Begin VB.Menu dso4 
            Caption         =   "Use these settings every time"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Qwer
Interior As String
End Type
Private beans As Qwer

Private Sub autofixe_Click()
If autofixe.Checked = True Then autofixe.Checked = False: GoTo nextme24
If autofixe.Checked = False Then autofixe.Checked = True
nextme24:
End Sub

Private Sub Command1_Click()
Form2.Text1 = Form2.Text1.text & setid & "- Log Displayed" & vbCrLf
setid = setid + 1
Form2.Show
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
PopupMenu selos, , Command2.Left, Command2.Top + Command2.Height
End Sub

Private Sub Command3_Click()
PopupMenu logopt, , Command2.Left, Command2.Top + Command2.Height
End Sub

Private Sub Command4_Click()
PopupMenu mainopt, , Command2.Left, Command2.Top + Command2.Height
End Sub

Private Sub Command5_Click()
If win98.Checked = True Then
Command5.Enabled = False
If scandisk.Checked = True Then
If Right(App.Path, 1) = "\" Then Shell App.Path & "det98.exe": GoTo nextme14
Shell App.Path & "\det98.exe"
nextme14:
If Right(App.Path, 1) = "\" Then Shell App.Path & "sdrun98.exe": GoTo nextme2
Shell App.Path & "\sdrun98.exe"
nextme2:
If stoss.Checked = True Then
Form2.Text1 = Form2.Text1.text & setid & "- Starting ScanDisk[Standard]" & "(at " & Time & ")" & vbCrLf
setid = setid + 1
End If
If stots.Checked = True Then
Form2.Text1 = Form2.Text1.text & setid & "- Starting ScanDisk[Thorough]" & "(at " & Time & ")" & vbCrLf
setid = setid + 1
End If
DoEvents
Do
but1% = FindWindow("ScanDskWDlgClass", vbNullString)
If but1% <> 0 Then Exit Do
Loop
If stoss.Checked = True Then
main% = FindWindow("ScanDskWDlgClass", vbNullString)
opt1% = FindChildByTitle(main%, "Stan&dard")
AOLButton (opt1%)
End If
If stots.Checked = True Then
main% = FindWindow("ScanDskWDlgClass", vbNullString)
opt1% = FindChildByTitle(main%, "&Thorough")
AOLButton (opt1%)
End If
'
'
'
'
but1% = FindWindow("ScanDskWDlgClass", vbNullString)
but2% = FindChildByTitle(but1%, "&Start")
thetext% = SendMessageByString(but1%, WM_SETTEXT, 0, "ScanDisk - vbAutoSpeed started at: " & Time)
AOLButton (but2%)
Do
but3% = FindChildByTitle(but1%, "Close")
If but3% <> 0 Then
AOLButton (but3%)
AOLButton (but3%)
Exit Do
End If
Loop
bs% = FindWindow("ThunderRT5Form", "vbAutoSpeed")
closes% = SendMessage(bs%, WM_CLOSE, 0, 0)
Form2.Text1 = Form2.Text1.text & setid & "- ScanDisk Complete" & "(at " & Time & ")" & vbCrLf
setid = setid + 1
End If
If defrag.Checked = True Then
If Right(App.Path, 1) = "\" Then Shell App.Path & "dfrun98.exe": GoTo nextme12
Shell App.Path & "\dfrun98.exe"
nextme12:
Form2.Text1 = Form2.Text1.text & setid & "- Starting Defrag" & "(at " & Time & ")" & vbCrLf
setid = setid + 1
Do
df1% = FindWindow("#32770", "Select Drive")
If df1% <> 0 Then
Exit Do
End If
Loop
df2% = FindWindow("#32770", "Select Drive")
df3% = FindChildByTitle(df2%, "OK")
AOLButton (df3%)
If Right(App.Path, 1) = "\" Then Shell App.Path & "det98.exe": GoTo nextme15
Shell App.Path & "\det98.exe"
nextme15:
Do
aa% = FindWindow("#32770", "Disk Defragmenter")
If aa% <> 0 Then Exit Do
Loop
df% = FindWindow("ThunderRT5Form", "vbAutoSpeed")
closes2% = SendMessage(df%, WM_CLOSE, 0, 0)
DoEvents
Form2.Text1 = Form2.Text1.text & setid & "- Defragmenter Complete" & "(at " & Time & ")" & vbCrLf
setid = setid + 1
End If
If XWD.Checked = True Then
Form2.Text1 = Form2.Text1.text & setid & "- Exiting..." & "(at " & Time & ")" & vbCrLf
setid = setid + 1
Unload Form1
End If
Command5.Enabled = True
End If
End Sub

Private Sub Command6_Click()
PopupMenu miscop, , Command2.Left, Command2.Top + Command2.Height
End Sub











Private Sub defrag_Click()
If defrag.Checked = True Then defrag.Checked = False: GoTo nextme3
If defrag.Checked = False Then defrag.Checked = True
nextme3:
End Sub

Private Sub displogonstartup_Click()
If displogonstartup.Checked = True Then displogonstartup.Checked = False: GoTo nextme20
If displogonstartup.Checked = False Then displogonstartup.Checked = True
nextme20:
End Sub

Private Sub dso1_Click()
If dso1.Checked = True Then dso1.Checked = False: GoTo nextme9
If dso1.Checked = False Then dso1.Checked = True
nextme9:
End Sub

Private Sub dso2_Click()
If dso2.Checked = True Then dso2.Checked = False: GoTo nextme10
If dso2.Checked = False Then dso2.Checked = True
nextme10:
End Sub

Private Sub dso3_Click()
If dso3.Checked = False Then dso3.Checked = True: dso4.Checked = False
End Sub

Private Sub dso4_Click()
If dso4.Checked = False Then dso4.Checked = True: dso3.Checked = False
End Sub

Private Sub Form_Load()
DataX = ""
setid = 1
setos = 0
On Error GoTo nextme19
If Right(App.Path, 1) = "\" Then Open App.Path & "opt.log" For Binary As #1: GoTo nextme17
Open App.Path & "\opt.log" For Binary As #1
nextme17:
Get #1, 1, beans
Close #1
DoEvents
DataX = beans.Interior
DataY = DataX
DoEvents
If Right(App.Path, 1) = "\" Then Kill App.Path & "\opt.log": GoTo nextme18
Kill App.Path & "\opt.log"
nextme18:
DoEvents
winl = InStr(DataX, "a")
dlsl = InStr(DataX, "b")
slxl = InStr(DataX, "c")
sdxl = InStr(DataX, "d")
dfxl = InStr(DataX, "e")
swd2l = InStr(DataX, "f")
xwd2l = InStr(DataX, "g")
winv = Mid(DataX, winl + 1, 1)
dlsv = Mid(DataX, dlsl + 1, 1)
slxv = Mid(DataX, slxl + 1, 1)
sdxv = Mid(DataX, sdxl + 1, 1)
dfxv = Mid(DataX, dfxl + 1, 1)
swd2v = Mid(DataX, swd2l + 1, 1)
xwd2v = Mid(DataX, xwd2l + 1, 1)
If winv = 1 Then win98.Checked = True
If winv = 2 Then winme.Checked = True
If dlsv = 1 Then displogonstartup.Checked = True
If dlsv = 0 Then displogonstartup.Checked = False
If slxv = 1 Then savelogonexit.Checked = True
If slxv = 0 Then savelogonexit.Checked = False
If sdxv = 1 Then scandisk.Checked = True
If sdxv = 0 Then scandisk.Checked = False
If dfxv = 1 Then defrag.Checked = True
If dfxv = 0 Then defrag.Checked = False
If swd2v = 1 Then SWD.Checked = True
If swd2v = 0 Then SWD.Checked = False
If xwd2v = 1 Then XWD.Checked = True
If xwd2v = 0 Then XWD.Checked = False
If displogonstartup.Checked = True Then
Form2.Text1 = Form2.Text1.text & setid & "- Log Displayed[Auto]" & vbCrLf
setid = setid + 1
Form2.Show
Form1.Command1.Enabled = False
End If
nextme19:
If DataX = "" Then DataY = "a1b0c1d1e1f1g1": DataX = "a1b0c1d1e1f1g1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If savelogonexit.Checked = True Then
If Form2.Text1.text = "" Then GoTo keepgoing
On Error Resume Next
If Right(App.Path, 1) = "\" Then Kill App.Path & "log.txt": GoTo nextme23
Kill App.Path & "\log.txt"
nextme23:
DoEvents
If Right(App.Path, 1) = "\" Then Open App.Path & "log.txt" For Binary Access Write As #1: GoTo nextme22
Open App.Path & "\log.txt" For Binary Access Write As #1
nextme22:
beans.Interior = Form2.Text1.text
Put #1, 1, beans
DoEvents
Close #1
DoEvents
End If
keepgoing:
If SWD.Checked = True Then
If Right(App.Path, 1) = "\" Then Open App.Path & "opt.log" For Binary As #1: GoTo nextme16
Open App.Path & "\opt.log" For Binary As #1
nextme16:
win = 0
dls = 0
slx = 0
sdx = 0
dfx = 0
swd2 = 0
xwd2 = 0
If win98.Checked = True Then win = 1
If winme.Checked = True Then win = 2
If displogonstartup.Checked = True Then dls = 1
If savelogonexit.Checked = True Then slx = 1
If scandisk.Checked = True Then sdx = 1
If defrag.Checked = True Then dfx = 1
If SWD.Checked = True Then swd2 = 1
If XWD.Checked = True Then xwd2 = 1
DataX = "a" & win & "b" & dls & "c" & slx & "d" & sdx & "e" & dfx & "f" & swd2 & "g" & xwd2
beans.Interior = DataX
Put #1, 1, beans
Close #1
End
Unload Me
End If
If SWD.Checked = False Then
If Right(App.Path, 1) = "\" Then Open App.Path & "opt.log" For Binary As #1: GoTo nextme21
Open App.Path & "\opt.log" For Binary As #1
nextme21:
beans.Interior = DataY
Put #1, 1, beans
Close #1
End
Unload Me
End If

End Sub

Private Sub savelogonexit_Click()
If savelogonexit.Checked = True Then savelogonexit.Checked = False: GoTo nextme6
If savelogonexit.Checked = False Then savelogonexit.Checked = True
nextme6:
End Sub

Private Sub scandisk_Click()
If scandisk.Checked = True Then scandisk.Checked = False: GoTo nextme4
If scandisk.Checked = False Then scandisk.Checked = True
nextme4:
End Sub

Private Sub sto1_Click()
If sto1.Checked = False Then sto1.Checked = True: sto2.Checked = False: sto3.Checked = False
End Sub

Private Sub sto2_Click()
If sto2.Checked = False Then sto2.Checked = True: sto1.Checked = False: sto3.Checked = False
End Sub

Private Sub sto3_Click()
If sto3.Checked = False Then sto3.Checked = True: sto1.Checked = False: sto2.Checked = False
End Sub

Private Sub sto4_Click()
If sto4.Checked = True Then sto4.Checked = False: GoTo nextme7
If sto4.Checked = False Then sto4.Checked = True
nextme7:
End Sub

Private Sub sto5_Click()
If sto5.Checked = True Then sto5.Checked = False: GoTo nextme8
If sto5.Checked = False Then sto5.Checked = True
nextme8:
End Sub

Private Sub stoss_Click()
If stoss.Checked = False Then stoss.Checked = True: stots.Checked = False
End Sub

Private Sub stots_Click()
If stots.Checked = False Then stots.Checked = True: stoss.Checked = False
End Sub

Private Sub SWD_Click()
If SWD.Checked = True Then SWD.Checked = False: GoTo nextme11
If SWD.Checked = False Then SWD.Checked = True
nextme11:
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub win98_Click()
win98.Checked = True
winme.Checked = False
setos = 0
End Sub

Private Sub winme_Click()
winme.Checked = True
win98.Checked = False
setos = 1
End Sub

Private Sub XWD_Click()
If XWD.Checked = True Then XWD.Checked = False: GoTo nextme5
If XWD.Checked = False Then XWD.Checked = True
nextme5:
End Sub
