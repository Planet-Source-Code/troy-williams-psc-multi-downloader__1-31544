VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "PSC Downloader"
   ClientHeight    =   3720
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   6495
   Icon            =   "mdiMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewDownloads 
         Caption         =   "&New Downloads"
      End
   End
   Begin VB.Menu mnuParse 
      Caption         =   "&Parse Email"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private cr As clsRegEntries
Private downloadDir As String
Private numDownloads As Long


Public Property Get getDownloadDir() As String
    getDownloadDir = downloadDir
End Property

Public Property Get getNumDownloads() As Long
    getNumDownloads = numDownloads
End Property


Private Sub MDIForm_Load()
Set cr = New clsRegEntries
With cr
    .RootKey = psrHKEY_LOCAL_MACHINE
    .MainBranch = "SOFTWARE"
    .RegBase = "BlueBill"
    .Program = "PSC Downloader"
    .Section = "Settings"
End With
    
downloadDir = cr.ReadEntry("DownloadDir", "")
If Right$(downloadDir, 1) <> "\" Then downloadDir = downloadDir & "\"
numDownloads = cr.ReadEntry("Downloads", "1")

frmInfo.Show
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set cr = Nothing
  
End Sub



Private Sub mnuNewDownloads_Click()
    frmURLlist.Show
    'loadDownloadForm
End Sub








Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuParse_Click()
    frmParseEmail.Show
End Sub
