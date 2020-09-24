Attribute VB_Name = "modMain"
Option Explicit


Private Sub main()
    
    mdiMain.Show

End Sub

'below is the contents of the frmDownload
'Option Explicit
'
''Private Declare Sub sapiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
'
'
'
'Private myURL As String 'this is the url that will be downloaded
'Private myServer As String 'used to hold the name of the computer - www.planet-source-code.com
'Private urlDir As String 'used to hold  /vb/default.asp?lngCId=10613&lngWId=1
'Private fileURL As String 'holds the url to the file that is to be downloaded
'Private downLoadFolder As String 'the name of the folder that the file and url are to be
''downloaded to....
'Private strHttpResponse As String 'the raw message from the server
'Private httpResponse As String 'the html part of the raw message
'Private httpHeader As String 'the header part of the raw message
'Private httpRequestHeader As String
'Private documentTitle As String
'Private cancelDownload As Boolean 'used to stop downloading
'Private downLoadFile As Boolean 'a flag used to indicate whether a file is being downed
'
'Private bytesDownloaded As Long 'used to keep track of the number of bytes downloaded
'Private foundHeaderBreak As Boolean 'used to indicate if the break between a header has been found in the dataarival event
'Private headerSize As Long 'this is the size of the header in bytes
'Private contentSize As Long 'this is the size of the body of the html transmission
'
'Private Sub Form_Resize()
'If Me.Width < 5640 Then
'    Me.Width = 5640
'End If
'
'If Me.Height < 3990 Then
'    Me.Height = 3990
'End If
'
'lblUrl.Width = Me.ScaleWidth
'lblStatus.Width = Me.ScaleWidth
'txtlog.Width = Me.ScaleWidth
'txtlog.Height = Me.ScaleHeight - 1560
'txtlog.Top = 1560
'txtlog.Left = 0
'
'End Sub
'
'Public Sub startDownload()
''This is used to initially start the download
'getAnURL
'
'If Not cancelDownload Then
'    downLoadFile = False
'    downLoadHTML
'
'Else
'    'Unload Me '-uncomment later
'End If
'
'End Sub
'
'Private Sub parseUrl(url As String)
''this function returns the server from an url
'Dim lStartPos
'
'url = Replace(url, "http://", "")
'
'lStartPos = InStr(1, url, "/", vbTextCompare)
'If lStartPos < 1 Then
'    urlDir = "/"
'    myServer = url
'    Exit Sub
'End If
'myServer = Left$(url, lStartPos - 1)
'urlDir = Right$(url, Len(url) - Len(myServer))
'
'End Sub
'
'Private Sub getAnURL()
''this sub retrieves an url from the frmURL to see if it has a download equivalent
'Dim sTemp() As String
'lblStatus.Caption = "Retrieving an URL to download"
'
'myURL = frmURLlist.getURL
'
'If StrComp(myURL, "Done", vbTextCompare) = 0 Then
''we are finished downloading
'    lblStatus.Caption = "No More Urls to download"
'    cancelDownload = True
'    Exit Sub
'End If
'
'parseUrl myURL
'
'End Sub
'
'
'Private Function getDocumentTitle(html As String) As String
'Dim posStart As Long
'Dim posEnd As Long
'posStart = InStr(1, html, "<title>", vbTextCompare) 'this will give us the starting position
'posEnd = InStr(1, html, "</title>", vbTextCompare) 'this will give us the ending position
'
'If posStart = 0 Or posEnd = 0 Then 'a webpage with no title has been downloaded
'    getDocumentTitle = "Unknown Error"
'    Exit Function
'End If
'
'posStart = posStart + 7 '7-is the number of chars in <title>
'
'getDocumentTitle = Mid$(html, posStart, posEnd - posStart) 'now we need to strip out "- visual basic, vb, vbscript"
'getDocumentTitle = Left$(getDocumentTitle, InStr(1, getDocumentTitle, "- visual basic, vb, vbscript", vbTextCompare) - 1)
'getDocumentTitle = Trim$(getDocumentTitle)
'
'End Function
'
'Private Sub processHTML()
''This is called, because it reached a page that has no redirects...Also should check to see if
''there was a problem with the number of connections....
'
''This sub will:
''-grab the title of the html page
''-create a folder with the name of title portion of the webpage
''-save a copy of the webpage in the folder
''-make sure there is a download link....'
''-start the file download process
''-save the file to the folder
'Dim sDirName As String
'Dim sDownloadLink As String
'
'sDirName = getDocumentTitle(httpResponse)
'
'sDirName = createDir(mdiMain.getDownloadDir & sDirName)
'
'downLoadFolder = sDirName
'
'writeHTML sDirName
'
'sDownloadLink = findDownloadLink(httpResponse)
'
'If StrComp(sDownloadLink, "No Link Found", vbTextCompare) = 0 Then
'    GoTo ReStart 'as there is no link on this page and no file to download
'End If
'
''figure out if the sDownloadLink is a relative reference or not
'
'If LCase$(Left$(sDownloadLink, 7)) = "http://" Then
'    parseUrl sDownloadLink
'Else
'    urlDir = sDownloadLink
'End If
''start downloading the file.....
'downLoadFile = True
'downLoadHTML
'
'
'
'
'Exit Sub
'ReStart:
'startDownload
'
'End Sub
'
'
''besure to write the address as a shortcut file *.url
'Private Sub writeHTML(dir As String)
''dir - the directory to deposit the html file
''this subroutine will write the strHttpResponse to a file called by the title of the document
'Dim sFileName As String
'Dim sTemp() As String
'Dim sTempHTML As String
'Dim fFile As Long
'
''first replace all the relative references.....
'If Right$(dir, 1) = "\" Then dir = Left$(dir, Len(dir) - 1)
'sTemp = Split(dir, "\")
'
'sFileName = dir & "\" & sTemp(UBound(sTemp)) & ".html"
'
'fFile = FreeFile
''output the header
'Open sFileName & ".txt" For Binary As fFile
'    Put #fFile, , httpHeader
'Close #fFile
'
'fFile = FreeFile
''output the unaltered html
'Open sFileName For Binary As fFile
'    Put #fFile, , httpResponse
'Close #fFile
'
''to create an internet shortcut....
''place the following text into a file called *.url
''[InternetShortcut]
''URL=http://www.microsoft.com/isapi/redir.dll?prd=ie&pver=5.0&ar=IStart
'
'fFile = FreeFile
''output the unaltered html
'sFileName = dir & "\PSC Download page.url"
'Open sFileName For Output As fFile
'    Print #fFile, "[InternetShortcut]"
'    Print #fFile, "URL=" & myServer & urlDir
'Close #fFile
'
'End Sub
'
'Private Sub writeFile(f() As Byte, dir As String, fileName As String)
'
'End Sub
'
''==============================
'Private Sub downLoadHTML()
''this is the part that downloads the html page
'
'lblUrl.Caption = myServer & urlDir
'lblStatus.Caption = "Begining Download"
'strHttpResponse = ""
'foundHeaderBreak = False
'headerSize = 0
'contentSize = 0
''==========================
'
'
'getURL
'
'End Sub
'
'
'Private Sub getURL()
'    Dim strPureURL As String
'    Dim strServerAddress As String
'    Dim strServerHostIP As String
'    Dim strDocumentURI As String
'    Dim lngStartPos As Long
'    Dim lngServerPort As Long
'    Dim strRequestTemplate As String
'
'    ' Note: This section of code (header) is based on code posted
'    ' by Tair Abdurman on http://www.planetsourcecode.com
'
'    httpHeader = ""
'    strRequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
'    "Accept: text/html, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
'    "Accept-Language: en-ca" & Chr(13) & Chr(10) & _
'    "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
'    "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows Donut; PSC Rulez!!)" & _
'    "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
'    "Connection: Keep-Alive" & Chr(13) & Chr(10) & _
'    "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
'    "Host: @$@@$@" & Chr(13) & Chr(10)
'
''    If strServerAddress = "" Or strDocumentURI = "" Then
''        MsgBox "Unable To detect target page!", vbCritical + vbOK
''        Exit Sub
''    End If
'
''    If mblnIsProxyUsed Then
''        strServerHostIP = txtProxy.Text
''        mstrRequestHeader = strRequestTemplate
''        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", mstrURL)
''        lngServerPort = 80
''    Else
''        strServerHostIP = strServerAddress
''        lngServerPort = 80
''        mstrRequestHeader = strRequestTemplate
''        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", strDocumentURI)
''    End If
'
'    'we are not worrying about a proxy for now
'
'    lblStatus.Caption = "Generating Request Header"
'
'    httpRequestHeader = strRequestTemplate
'    httpRequestHeader = Replace(httpRequestHeader, "_$-$_$-", urlDir) 'the relative path to the file to be downloaded -if this was proxied then it would be the full url
'    httpRequestHeader = Replace(httpRequestHeader, "@$@@$@", myServer)
'    httpRequestHeader = httpRequestHeader & vbCrLf
'
'    strHttpResponse = "" 'reset the raw data collector
'    lblStatus.Caption = "Attempting to Connect to " & myServer
'
'    If winsockHTTP.State = sckClosed Then
'        winsockHTTP.Connect myServer, 80
'    Else
'        sendRequest
'    End If
'End Sub
'
'
'
'Private Sub redirection()
''this sub handles a redirection if any
'Dim redirectUrl As String
'Dim pos As Long
'Dim posEnd As Long
'pos = InStr(1, httpResponse, "<a HREF=", vbTextCompare)
'posEnd = InStr(pos, httpResponse, ">", vbTextCompare)
'pos = pos + 8
'redirectUrl = Mid$(httpResponse, pos, posEnd - pos)
'redirectUrl = Replace$(redirectUrl, Chr$(34), "")
'
'If Left$(redirectUrl, 7) <> "http://" Then
'    urlDir = redirectUrl
'Else
'    parseUrl redirectUrl
'End If
'End Sub
'
''==============================
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
'
'winsockHTTP.Close
'
'End Sub
'
''===============================
'Private Sub sendRequest()
'lblStatus.Caption = "Connected to " & myServer & ", sending request"
'bytesDownloaded = 0
''Sleep 1000
'winsockHTTP.SendData httpRequestHeader
'lblStatus.Caption = "Request sent to " & myServer
'End Sub
'
'Private Sub winsockHTTP_Connect()
'sendRequest
''winsockHTTP.SendData httpRequestHeader
'
'    'txtStatus.Text = txtStatus.Text & "Connected, try To obtain page ..." & vbCrLf
'    'txtStatus.Refresh
'    'frmMainWin.txtResponse.Text = ""
'    'frmMainWin.txtResponse.Refresh
'End Sub
'
'Private Sub winsockHTTP_DataArrival(ByVal bytesTotal As Long)
'Dim sData As String
'Dim b() As Byte
'Dim fName As String
'
'lblStatus.Caption = "Retrieving Data from " & myServer
'bytesDownloaded = bytesDownloaded + bytesTotal
'lblBytes.Caption = bytesDownloaded
'
'winsockHTTP.GetData sData, vbString
'    strHttpResponse = strHttpResponse & sData
''
''If Not foundHeaderBreak Then
''    winsockHTTP.GetData sData, vbString
''    strHttpResponse = strHttpResponse & sData
''
''    If InStr(1, sData, vbCrLf & vbCrLf) And Not foundHeaderBreak Then
''        'the first one found will indicate the break between the header and the body
''        foundHeaderBreak = True
''        headerSize = Len(strHttpResponse)
''        contentSize = contentLength(strHttpResponse)
''    End If
''
''
''
''Else
''    If downLoadFile Then
''    'winsockHTTP.GetData b(), vbByte
''    Else
''        winsockHTTP.GetData sData, vbString
''        strHttpResponse = strHttpResponse & sData
''    End If
''End If
''
'txtlog.Text = strHttpResponse
'
'    If InStr(1, sData, vbCrLf & vbCrLf) And Not foundHeaderBreak Then
'        'the first one found will indicate the break between the header and the body
'        foundHeaderBreak = True
'        headerSize = Len(strHttpResponse)
'        contentSize = contentLength(strHttpResponse)
'    End If
'
'
'    If bytesDownloaded - headerSize = contentSize Then
'        foundHeaderBreak = False
'        headerSize = 0
'        contentSize = 0
'        If downLoadFile Then
'            'write the thing to the hardisk
'           ' writeFile b(), downLoadFolder, fName
'        Else
'            processWebData
'        End If
'    End If
'
'End Sub
'
'Private Sub processWebData()
''this sub takes the raw data received from the webserver and
''extracts the header and the body
''it also determines what the error code is and the appropriate response to take
'
'Dim sTemp(0 To 1) As String
'Dim posStart As Long
'Dim posEnd As Long
'
'lblStatus.Caption = "Finished Retrieving Data from " & myServer & ". Beginning to Process..."
'lblBytes.Caption = ""
'
'posStart = InStr(strHttpResponse, vbCrLf & vbCrLf)
'sTemp(0) = Left$(strHttpResponse, posStart)
'
'posStart = posStart + Len(vbCrLf & vbCrLf)
'posEnd = Len(strHttpResponse) + 1
'sTemp(1) = Mid$(strHttpResponse, posStart, posEnd - posStart)
'
'httpHeader = sTemp(0)
'httpResponse = sTemp(1)
'
'txtlog.Text = httpHeader & vbNewLine
'txtlog.Text = txtlog.Text & httpResponse & vbNewLine
'
'Select Case modHTTP.GetHttpResponseCode(httpHeader)
'    Case 300, 301, 302, 303, 307
'        'handle the redirection
'        lblStatus.Caption = "Redirecting to another URL"
'        redirection
'        downLoadHTML
'        Exit Sub
'    Case 200 'ok
'        If Not downLoadFile Then
'            processHTML
'        Else
'            'write the file to the disk
'            'downLoadFile = False
'
'        End If
'    Case 400
'        txtlog.Text = txtlog.Text & "Bad Request..."
'        Exit Sub
'End Select
'
''startDownload
'
'End Sub
'
'
'Private Sub winsockHTTP_Close()
''
'End Sub
'
'Private Sub winsockHTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
''txtStatus.Text = txtStatus.Text & "Errors occured ..." & vbCrLf
''txtStatus.Text = txtStatus.Text & "Number: " & Number & "  Description: " & Description & vbCrLf
''txtStatus.Refresh
'End Sub
'
'
'Private Sub Sleep(lngMilliSec As Long)
'    If lngMilliSec > 0 Then
'        Call sapiSleep(lngMilliSec)
'    End If
'End Sub
'
'
'
''Winsock states
''Constant               Value                   Description
''sckClosed              0                       Default. Closed
''sckOpen                1                       Open
''sckListening           2                       Listening
''sckConnectionPending   3                       Connection pending
''sckResolvingHost       4                       Resolving host
''sckHostResolved        5                       Host resolved
''sckConnecting          6                       Connecting
''sckConnected           7                       Connected
''sckClosing             8                       Peer is closing the connection
''sckError               9                       Error
'
'

