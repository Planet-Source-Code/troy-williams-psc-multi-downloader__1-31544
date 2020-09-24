VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6660
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDownloads 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDownload 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Maximum 4 Connections    (Please don't overload the PSC servers)"
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Simultaneous Downloads"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Download Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdBrowse_Click()
Dim folder As String
    folder = FolderBrowse(0, "Locate the folder to store the Source code in...")
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    txtDownload.Text = folder
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim cr As New clsRegEntries

With cr
    .RootKey = psrHKEY_LOCAL_MACHINE
    .MainBranch = "SOFTWARE"
    .RegBase = "BlueBill"
    .Program = "PSC Downloader"
    .Section = "Settings"
End With
    
If Not IsNumeric(txtDownloads.Text) Then
    txtDownloads.Text = 1
Else
    If Val(txtDownloads.Text) > 4 Then
        txtDownloads.Text = 4
    End If
End If
    
cr.WriteEntry "DownloadDir", txtDownload.Text
cr.WriteEntry "Downloads", txtDownloads.Text
MsgBox "You need to restart the Program for changes to take effect!", vbExclamation
Set cr = Nothing
End Sub

Private Sub Form_Load()
    txtDownload.Text = mdiMain.getDownloadDir
    txtDownloads.Text = mdiMain.getNumDownloads
End Sub
