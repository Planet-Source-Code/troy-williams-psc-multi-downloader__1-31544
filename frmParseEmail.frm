VERSION 5.00
Begin VB.Form frmParseEmail 
   Caption         =   "Parse Email from PSC"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "frmParseEmail.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstEmail 
      Height          =   1425
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move To URL list"
      Default         =   -1  'True
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete From List"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "Clipboard"
      Visible         =   0   'False
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUrlCount 
         Caption         =   "Url Count:"
      End
   End
End
Attribute VB_Name = "frmParseEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
Dim i As Long, upper As Long

upper = lstEmail.ListCount - 1

While i < lstEmail.ListCount
    If lstEmail.Selected(i) Then
        lstEmail.RemoveItem i
        i = lstEmail.ListCount
    End If
    i = i + 1
Wend

mnuUrlCount.Caption = "Url Count: " & lstEmail.ListCount
End Sub

Private Sub cmdMove_Click()
Dim sList As String 'this holds the urls separated by a vbnewline statement
Dim i As Long, upper As Long

upper = lstEmail.ListCount - 1
For i = 0 To upper
    sList = sList & lstEmail.List(i) & vbNewLine
Next 'i

frmURLlist.Show
frmURLlist.loadURLfromStringList sList

End Sub

Private Sub Form_Resize()

If Me.Height < 6045 Then
    Me.Height = 6045
End If

If Me.Width < 5000 Then
    Me.Width = 5000
End If

lstEmail.Top = Me.ScaleTop
lstEmail.Left = Me.ScaleLeft
lstEmail.Height = Me.ScaleHeight - 400
lstEmail.Width = Me.ScaleWidth

cmdDelete.Top = Me.ScaleHeight - 350
cmdMove.Top = Me.ScaleHeight - 350

cmdDelete.Left = Me.ScaleWidth / 2 - 1400
cmdMove.Left = Me.ScaleWidth / 2 + 100


End Sub


Private Sub lstEMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then 'the right mouse button was clicked
        Me.PopupMenu mnuClipboard
    End If
End Sub

Private Sub lstEMail_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sInput As String

If Data.GetFormat(vbCFText) Then
   sInput = Data.GetData(vbCFText)
   parseEmail sInput
End If

End Sub

Private Sub mnuPaste_Click()
Dim sClipboardInfo As String
    sClipboardInfo = Clipboard.GetText
    parseEmail sClipboardInfo
End Sub


Private Sub parseEmail(sEmail As String)
'this sub goes through the input line by line looking for urls.
Dim sLines() As String
Dim i As Long, upper As Long
sLines = Split(sEmail, vbNewLine) 'the email message should be delimited by newlines
upper = UBound(sLines)

For i = 0 To upper
    If isURL(sLines(i)) Then
        If isPSC(sLines(i)) Then
            lstEmail.AddItem sLines(i)
        End If
    End If
Next i

mnuUrlCount.Caption = "Url Count: " & lstEmail.ListCount
Erase sLines
End Sub



Private Function isURL(url As String) As Boolean
isURL = False
If Left$(url, 7) = "http://" Then isURL = True

End Function


Private Function isPSC(url As String) As Boolean
isPSC = False
    If InStr(1, url, "www.planet-source-code", vbTextCompare) > 0 Then
        isPSC = True
        Exit Function
    ElseIf InStr(1, url, "www.planetsourcecode", vbTextCompare) > 0 Then
        isPSC = True
    End If
End Function
