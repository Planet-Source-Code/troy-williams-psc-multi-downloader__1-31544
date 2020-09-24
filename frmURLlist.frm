VERSION 5.00
Begin VB.Form frmURLlist 
   Caption         =   "Url List"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5925
   Icon            =   "frmURLlist.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   5925
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      ToolTipText     =   "Clear the list"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Remove Selected Rows"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstUrl 
      Height          =   2400
      ItemData        =   "frmURLlist.frx":0442
      Left            =   0
      List            =   "frmURLlist.frx":0444
      MultiSelect     =   1  'Simple
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Menu mnuClipboard 
      Caption         =   "ClipBoard"
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
      Begin VB.Menu mnuSave 
         Caption         =   "Save to File"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUrls 
         Caption         =   "Urls: "
      End
   End
End
Attribute VB_Name = "frmURLlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private urlCol As Collection


Public Function getURL() As String
'this function returns a valid url from the url collection
Dim i As Long, upper As Long
Dim url As clsUrl

upper = urlCol.count

For i = 1 To upper
    Set url = urlCol.Item(i)
    If Not url.DownLoaded Then
        getURL = url.url
        url.DownLoaded = True
    Exit Function
    End If
Next 'i

If i >= upper Then
    getURL = "Done" 'the message for the threads to quit
End If

Set url = Nothing
End Function


Private Sub cmdClear_Click()
lstUrl.Clear

Set urlCol = Nothing
Set urlCol = New Collection

End Sub

Private Sub cmdDownload_Click()
'this routine simple fires up the appropriate number of download windows
'It divides the number of urls to download from by the number of open windows
'and passes each of them their allocation.
Dim count As Long
Dim i As Long

Dim f As frmDownload

count = mdiMain.getNumDownloads
'create the forms

frmLog.Show 'start the log

'only create one for now
For i = 1 To count
    Set f = New frmDownload
    f.Show
    f.startDownload
    Set f = Nothing
Next 'i

End Sub

Private Sub cmdRemove_Click()
Dim i As Long, upper As Long

upper = lstUrl.ListCount - 1

While i < lstUrl.ListCount
    If lstUrl.Selected(i) Then
        lstUrl.RemoveItem i
        urlCol.Remove i
        i = lstUrl.ListCount
    End If
    i = i + 1
Wend
mnuUrls.Caption = "Url Count: " & lstUrl.ListCount
End Sub

Private Sub Form_Load()
Set urlCol = New Collection
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set urlCol = Nothing
End Sub

Private Sub Form_Resize()

If Me.Height < 6045 Then
    Me.Height = 6045
    
End If

If Me.Width < 5000 Then
    Me.Width = 5000
End If

lstUrl.Top = Me.ScaleTop
lstUrl.Left = Me.ScaleLeft
lstUrl.Height = Me.ScaleHeight - 400
lstUrl.Width = Me.ScaleWidth

cmdRemove.Top = Me.ScaleHeight - 350
cmdClear.Top = Me.ScaleHeight - 350
cmdDownload.Top = Me.ScaleHeight - 350

cmdRemove.Left = Me.ScaleWidth / 2 - 1830
cmdClear.Left = Me.ScaleWidth / 2 - 600
cmdDownload.Left = Me.ScaleWidth / 2 + 610

End Sub




Private Sub Form_Unload(Cancel As Integer)
Set urlCol = Nothing
End Sub

Private Sub lstUrl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then 'the right mouse button was clicked
        Me.PopupMenu mnuClipboard
    End If
End Sub

Private Sub lstUrl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sInput As String

If Data.GetFormat(vbCFText) Then
   sInput = Data.GetData(vbCFText)
   loadURLfromStringList sInput
End If
End Sub

Private Sub mnuPaste_Click()
Dim sClipboardInfo As String
    sClipboardInfo = Clipboard.GetText
    loadURLfromStringList sClipboardInfo
End Sub

Public Sub loadURLfromStringList(urlList As String)
Dim sArray() As String
Dim i As Long, upper As Long
Dim url As clsUrl

If Len(urlList) > 1 Then
    sArray = Split(urlList, vbNewLine)
    
    sArray = killDups(sArray)
    
    upper = UBound(sArray)
    For i = 0 To upper
        If isURL(sArray(i)) Then
            Set url = New clsUrl
            url.url = sArray(i)
            url.DownLoaded = False
            urlCol.Add url
        End If
    Next 'i
End If



loadURLlist
Set url = Nothing
End Sub


Private Function killDups(sIn() As String) As String()
'this function only strips dups out of the clipboard data and not the collection..
Dim sUnique() As String 'the unique string
Dim i As Long, upper As Long, lower As Long, j As Long
Dim count As Long 'used to hold the actual number of unique items in the sUnique
upper = UBound(sIn)
lower = LBound(sIn)


'count = upper '\ 2
'If count < 1 Then count = 5

ReDim sUnique(lower To upper) 'this could be set up a little better, but since it probably won't deal with more then a hundred links a at a time.... it shouldn't be a problem
count = 0

sUnique(lower) = sIn(lower)
count = count + 1

For i = lower To upper
    For j = lower To count
        If sIn(i) = sUnique(j) Then
            Exit For
        ElseIf j = count Then
            count = count + 1
            sUnique(count) = sIn(i)
        End If
    Next 'j
    
Next 'i

ReDim Preserve sUnique(lower To count)
killDups = sUnique

End Function


Private Function isURL(url As String) As Boolean

    If Left$(url, 7) = "http://" Then isURL = True

End Function


Private Sub loadURLlist()
'this sub routine takes the contents of the url collection and loads
'it into the listbox
Dim url As clsUrl

lstUrl.Clear
For Each url In urlCol
    lstUrl.AddItem url.url
Next 'url
lstUrl.ToolTipText = "Url Count: " & lstUrl.ListCount
mnuUrls.Caption = "Url Count: " & lstUrl.ListCount
Set url = Nothing
End Sub

Private Sub mnuSave_Click()
'this menu option allows the url list to be saved to a file....
Dim url As clsUrl
Dim fFile As Long
Dim sTemp As String


For Each url In urlCol
    sTemp = sTemp & url.url & vbNewLine
Next 'url

fFile = FreeFile

fFile = FreeFile
Open mdiMain.getDownloadDir & "\PSC urls to dl " & Format(Now, "mmmm dd yyyy hh mm ampm") & ".txt" For Binary As #fFile
    Put #fFile, , sTemp
Close #fFile

MsgBox "Saved in the Source download  folder..."

Set url = Nothing
End Sub
