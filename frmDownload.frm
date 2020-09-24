VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   Caption         =   "Download"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5610
   Icon            =   "frmDownload.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   3105
   ScaleWidth      =   5610
   Begin VB.TextBox txtlog 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
   Begin MSWinsockLib.Winsock winsockHTTP 
      Left            =   4920
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblBytes 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label lblUrl 
      Caption         =   "Current Url:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu mnuInvisible 
      Caption         =   "Invisible"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Sub sapiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private msFullURL As String ' the full url from the frmUrlList
Private cancelDownload As Boolean 'used to cancel the download..


'note these properties are for the current url being downloaded
Private sServer As String 'used to hold the www.planetsourcecode.com part of the full url
Private urlDir As String 'used to hold the /vb/blahbalh part of the full url

Private sResponseData As String 'this is used to hold the data from the winsock dataarive event
Private bytesDownloaded As Long 'this is used to hold the number of bytes that are downloaded from the server
Private foundHeaderBreak As Boolean 'this is used to indicate the header break has been found


Private httpHeader As String 'this is used to hold the header information of the url
Private httpResponse As String 'this is used to hold the html of the url
Private headerSize As Long 'the size of the header in bytes
Private contentSize As Long 'the size of the payload in bytes
Private bFile As Boolean 'this indicates if the httpHeader indicates the presences of a file

Private downLoadFolder As String 'the directory to download the file

Private httpRequestHeader As String 'used to hold the header used to request info from the webserver



'===========================================
Private Sub Form_Resize()
If Me.Width < 5640 Then
    Me.Width = 5640
End If

If Me.Height < 3990 Then
    Me.Height = 3990
End If

lblUrl.Width = Me.ScaleWidth
lblStatus.Width = Me.ScaleWidth
txtLog.Width = Me.ScaleWidth
txtLog.Height = Me.ScaleHeight - 1560
txtLog.Top = 1560
txtLog.Left = 0

End Sub

'==================================
Public Sub startDownload()
'This sub is called from the frmUrlList

getAnURL

If Not cancelDownload Then
        downLoadHTML
Else
    'Unload Me '-uncomment later
End If

End Sub

Private Sub getAnURL()
'this sub retrieves an url from the frmURL to see if it has a download equivalent
Dim sTemp() As String
lblStatus.Caption = "Retrieving an URL to download"

msFullURL = frmURLlist.getURL

If StrComp(msFullURL, "Done", vbTextCompare) = 0 Then
'we are finished downloading
    lblStatus.Caption = "No More Urls to download"
    cancelDownload = True
    Exit Sub
End If

parseUrl msFullURL
lblUrl.Caption = msFullURL

End Sub

Private Sub parseUrl(url As String)
'this function returns the server from an url
Dim lStartPos

url = Replace(url, "http://", "")

lStartPos = InStr(1, url, "/", vbTextCompare)
If lStartPos < 1 Then 'there is no /vb/adfkja.zip
    urlDir = "/"
    sServer = url
    Exit Sub
End If
sServer = Left$(url, lStartPos - 1)
urlDir = Right$(url, Len(url) - Len(sServer))

End Sub

Private Sub downLoadHTML()
''this is the part that downloads the html page
'
lblUrl.Caption = sServer & urlDir
lblStatus.Caption = "Begining Download"
'strHttpResponse = ""
'foundHeaderBreak = False
'headerSize = 0
'contentSize = 0
''==========================

'initialize totalizer variables
sResponseData = "" 'clear out this value before we start receiving data from the http server
bytesDownloaded = 0
foundHeaderBreak = False



'start the ball rolling
httpRequestHeader = createRequestHeader(sServer, urlDir)

contactServer

End Sub


Private Sub contactServer()

lblStatus.Caption = "Attempting to Connect to " & sServer
If winsockHTTP.State = sckClosed Then
    winsockHTTP.Connect sServer, 80
Else
    sendRequest
End If

End Sub


Private Sub sendRequest()
lblStatus.Caption = "Sending Request for a connnection to " & sServer

winsockHTTP.SendData httpRequestHeader

End Sub


'================================
'winsock events
Private Sub winsockHTTP_Connect()
    sendRequest
End Sub

Private Sub winsockHTTP_DataArrival(ByVal bytesTotal As Long)
Dim sData As String


lblStatus.Caption = "Retrieving Data from " & sServer
If Not foundHeaderBreak Then
    bytesDownloaded = bytesDownloaded + bytesTotal
    lblBytes.Caption = "Bytes: " & bytesDownloaded
Else
    'now we can accurately report the bytes being downloaded
    bytesDownloaded = bytesDownloaded + bytesTotal
    lblBytes.Caption = "Bytes: " & bytesDownloaded - headerSize & " / " & contentSize
End If

winsockHTTP.GetData sData, vbString
sResponseData = sResponseData & sData

txtLog.Text = sResponseData 'remove later

    If InStr(1, sData, vbCrLf & vbCrLf) And Not foundHeaderBreak Then
        'the first one found will indicate the break between the header and the body
        foundHeaderBreak = True
        httpHeader = sResponseData
        headerSize = Len(sResponseData)
        contentSize = contentLength(sResponseData)
        'we should analyze the header for -  Content-Type: application/x-zip-compressed or Content-Type: text/html
        If StrComp(contentType(httpHeader), "File", vbTextCompare) = 0 Then
            bFile = True
        End If
    End If
    
    

    If bytesDownloaded - headerSize = contentSize Then
    'if it gets in here then it is done.....As that's what the file says
    lblStatus.Caption = "Finished retrieving data from " & sServer
        foundHeaderBreak = False
        headerSize = 0
        contentSize = 0
        
        
        'we have to process the webpage that we downloaded - see processHTML
        processWebData
        
    End If

End Sub
'============================

Private Sub processWebData()
'this sub takes the raw data received from the webserver and
'extracts the body
'it also determines what the error code is and the appropriate response to take

lblStatus.Caption = "Finished Retrieving Data from " & sServer & ". Beginning to Process..."
lblBytes.Caption = ""


httpResponse = extractHTML(sResponseData)

'================================
txtLog.Text = httpHeader & vbNewLine
txtLog.Text = txtLog.Text & httpResponse & vbNewLine
'================================

Select Case modHTTP.GetHttpResponseCode(httpHeader)
    Case 300, 301, 302, 303, 307
        'handle the redirection
        lblStatus.Caption = "Redirecting to another URL"
        redirection
        downLoadHTML
        Exit Sub
    Case 200 'ok
        'in here we are at a webpage that doesn't have a redirect....
        'search the page and see if it has the "Download Code" link...
        
        'check the content type and handle it appropriately...
        'see the whatContentType function
        
        If Not bFile Then
            processHTML
        Else
            'we have a file
            'at this point a directory should already be set up for the download of the file
            'we just need to name the file
            writeFile
            bFile = False
            startDownload
        End If

    Case 400
        txtLog.Text = txtLog.Text & "Bad Request..."
        frmLog.writeLine "Bad Request... " & sServer & urlDir
        Exit Sub
    Case Else
        txtLog.Text = txtLog.Text & "No Links Found on this page..."
        frmLog.writeLine "No Links Found on this Page... " & sServer & urlDir
End Select

'startDownload

End Sub

Private Sub writeFile()
'this sub will write the file to disk in the directory that has been established
Dim sFilename As String
Dim sTemp() As String
Dim sTempHTML As String
Dim fFile As Long

'first replace all the relative references.....
If Right$(downLoadFolder, 1) = "\" Then downLoadFolder = Left$(downLoadFolder, Len(downLoadFolder) - 1)
sTemp = Split(downLoadFolder, "\")

sFilename = downLoadFolder & "\" & sTemp(UBound(sTemp)) & ".zip"

fFile = FreeFile

Open sFilename For Binary As #fFile
    Put #fFile, , httpResponse
Close #fFile

fFile = FreeFile


Open downLoadFolder & "\fileHeader.txt" For Binary As #fFile
    Put #fFile, , httpHeader
Close #fFile

txtLog.Text = sFilename & "  Written to disk..."
End Sub

Private Sub redirection()
'this sub handles a redirection
'what it does is search the PSC page for a reference link.....
Dim redirectUrl As String
Dim pos As Long
Dim posEnd As Long
pos = InStr(1, httpResponse, "<a HREF=", vbTextCompare)
posEnd = InStr(pos, httpResponse, ">", vbTextCompare)
pos = pos + 8
redirectUrl = Mid$(httpResponse, pos, posEnd - pos)
redirectUrl = Replace$(redirectUrl, Chr$(34), "")

If Left$(redirectUrl, 7) <> "http://" Then
    urlDir = redirectUrl
Else
    parseUrl redirectUrl
End If
End Sub

Private Sub processHTML()
'This is called, because it reached a page that has no redirects...

'This sub will:
'-grab the title of the html page
'-create a folder with the name of title portion of the webpage
'-save a copy of the webpage in the folder
'-make sure there is a download link....'
Dim sDirName As String
Dim sDownloadLink As String

sDirName = modHTTP.getDocumentTitle(httpResponse)

sDirName = createDir(mdiMain.getDownloadDir & sDirName)

writeHTML sDirName

downLoadFolder = sDirName

'check to see if there is a download link
sDownloadLink = findDownloadLink(httpResponse)

If StrComp(sDownloadLink, "No Link Found", vbTextCompare) = 0 Then
    
    frmLog.writeLine "No Download Link... " & sServer & urlDir
    GoTo ReStart 'as there is no link on this page and no file to download
End If

'figure out if the sDownloadLink is a relative reference or not

If LCase$(Left$(sDownloadLink, 7)) = "http://" Then
    parseUrl sDownloadLink
Else
    urlDir = sDownloadLink
End If
'start downloading the file.....
'downLoadFile = True
downLoadHTML

Exit Sub
ReStart:
startDownload

End Sub

'besure to write the address as a shortcut file *.url
Private Sub writeHTML(dir As String)
'dir - the directory to deposit the html file
'this subroutine will write the strHttpResponse to a file called by the title of the document
Dim sFilename As String
Dim sTemp() As String
Dim sTempHTML As String
Dim fFile As Long

'first replace all the relative references.....
If Right$(dir, 1) = "\" Then dir = Left$(dir, Len(dir) - 1)
sTemp = Split(dir, "\")

sFilename = dir & "\" & sTemp(UBound(sTemp)) & ".html"

fFile = FreeFile
'output the header
Open sFilename & ".txt" For Binary As fFile
    Put #fFile, , httpHeader
Close #fFile

fFile = FreeFile
'output the unaltered html
Open sFilename For Binary As fFile
    Put #fFile, , httpResponse
Close #fFile

'to create an internet shortcut....
'place the following text into a file called *.url
'[InternetShortcut]
'URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.0&ar=IStart

fFile = FreeFile
'output the unaltered html
sFilename = dir & "\PSC Download page.url"
Open sFilename For Output As fFile
    Print #fFile, "[InternetShortcut]"
    Print #fFile, "URL=" & sServer & urlDir
Close #fFile

End Sub


'====================================================================
'Winsock states
'Constant               Value                   Description
'sckClosed              0                       Default. Closed
'sckOpen                1                       Open
'sckListening           2                       Listening
'sckConnectionPending   3                       Connection pending
'sckResolvingHost       4                       Resolving host
'sckHostResolved        5                       Host resolved
'sckConnecting          6                       Connecting
'sckConnected           7                       Connected
'sckClosing             8                       Peer is closing the connection
'sckError               9                       Error

