VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "URL Error Log"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOutput 
      Caption         =   "Write to File"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtLog 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cr As clsRegEntries
Private downLoad As String

Public Function writeLine(sLine As String)
    txtLog.Text = txtLog.Text & sLine & vbNewLine
End Function

Public Sub clearLog()
    txtLog.Text = ""
End Sub

Private Sub cmdOutput_Click()
Dim fFile As Long

fFile = FreeFile
Open downLoad & "\Log " & Format(Now, "mmmm dd yyyy hh mm ampm") & ".txt" For Binary As #fFile
    Put #fFile, , txtLog.Text
Close #fFile
clearLog
End Sub

Private Sub Form_Load()
Set cr = New clsRegEntries
With cr
    .RootKey = psrHKEY_LOCAL_MACHINE
    .MainBranch = "SOFTWARE"
    .RegBase = "BlueBill"
    .Program = "PSC Downloader"
    .Section = "Settings"
End With
    
downLoad = cr.ReadEntry("DownloadDir", "")

Set cr = Nothing
End Sub

Private Sub Form_Resize()
txtLog.Top = Me.ScaleTop
txtLog.Left = Me.ScaleLeft
txtLog.Width = Me.ScaleWidth
txtLog.Height = Me.ScaleHeight - 500

cmdOutput.Top = Me.ScaleHeight - 400
cmdOutput.Left = Me.ScaleWidth \ 2 - 600
End Sub
