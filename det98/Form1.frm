VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbAutoSpeed"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2265
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detecting Close Screen..."
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
m2% = FindWindow("#32770", vbNullString)
m3% = FindChildByTitle(m2%, "Close")
If m3% <> 0 Then
AOLButton (m3%)
End If
End Sub

Private Sub Timer2_Timer()
dfx1% = FindWindow("#32770", "Disk Defragmenter")
dfx2% = FindChildByTitle(dfx1%, "&Yes")
If dfx2% <> 0 Then
AOLButton (dfx2%)
End If
End Sub
