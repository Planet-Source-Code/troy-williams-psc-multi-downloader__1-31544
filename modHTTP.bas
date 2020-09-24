Attribute VB_Name = "modHTTP"
Option Explicit

'this module will hold generic http functions...

Public Function GetHttpResponseCode(strHeader As String) As Long
'this routine parses a string header and returns an error code
Dim varCode As String

varCode = Mid(strHeader, InStr(1, strHeader, " ") + 1, 3)
If IsNumeric(varCode) Then
    GetHttpResponseCode = CInt(varCode)
End If

End Function


Public Function createDir(dirName As String) As String
'this routine will create a folder with the name of the html file
'and return the name and path of the newly created directory
Dim i As Long
Dim sTemp As String

If Len(dir$(dirName, vbDirectory)) = 0 Then
    MkDir dirName
Else
    sTemp = dirName
    For i = 1 To 1000 'the number of folders that can have the same name
        If Len(dir$(sTemp & "." & i, vbDirectory)) = 0 Then
            MkDir sTemp & "." & i
            dirName = sTemp & "." & i
            Exit For
        End If
    Next 'i
End If

createDir = dirName
End Function


Public Function findDownloadLink(html As String) As String
Dim posStart As Long, posEnd As Long, pos As Long
Dim sLink As String

'<a href="/vb/scripts/ShowZip.asp?lngWId=1&lngCodeId=30651&strZipAccessCode=PDAT306514112"><img border="0" src="/vb/images//winzipicon_medium.gif" alt="winzip icon" width="42" height="41">Download code</a>

pos = InStr(1, html, "Download code", vbTextCompare) 'this will give us a position after the download link...

If pos = 0 Then
    findDownloadLink = "No Link Found"
    Exit Function
End If

posStart = InStrRev&(html, "<a href", pos, vbTextCompare)
posStart = posStart + 7
posEnd = InStr(posStart, html, ">", vbTextCompare) 'this will give us the ending position

sLink = Mid$(html, posStart, posEnd - posStart)
sLink = Replace$(sLink, Chr$(34), "")

pos = InStr(1, sLink, "=", vbTextCompare)
sLink = Mid$(sLink, pos + 1, Len(sLink))
sLink = Trim$(sLink)

findDownloadLink = sLink
End Function

Public Function contentLength(sHeader As String) As Long
'this function finds the "Content-Length:" in the header and returns the
'number of bytes
Dim startPos As Long
Dim endPos As Long
Dim i As Long
Dim upper As Long

startPos = InStr(1, sHeader, "Content-Length:", vbTextCompare)
startPos = startPos + Len("Content-Length: ")

endPos = startPos
upper = Len(sHeader)
For i = 1 To upper
    If IsNumeric(Mid$(sHeader, endPos, 1)) Then
        endPos = endPos + 1
    Else
        Exit For
    End If
Next i

contentLength = CLng(Trim$(Mid$(sHeader, startPos, endPos - startPos)))

End Function

Public Function contentType(sHeader As String) As String
'this function is used to determine what type of content is being downloaded
'if the content type is: Content-Type: application/x-zip-compressed then return "File"
'if it is: Content-Type: text/html then return "HTML"
Dim startPos As Long
Dim endPos As Long
Dim i As Long
Dim upper As Long

startPos = InStr(1, sHeader, "Content-Type:", vbTextCompare)
startPos = startPos + Len("Content-Type: ")

endPos = InStr(startPos, sHeader, "/", vbTextCompare)

contentType = Trim$(Mid$(sHeader, startPos, endPos - startPos))

If StrComp(contentType, "text", vbTextCompare) = 0 Then
    contentType = "HTML"
    Exit Function
End If

If StrComp(contentType, "application", vbTextCompare) = 0 Then
    'we know it is an application, but is it a zip
    If Mid$(sHeader, endPos + 1, Len("x-zip-compressed")) = "x-zip-compressed" Then
    contentType = "File"
    Exit Function
    End If
End If

contentType = "Unknown"

End Function

'Private Function getHTMLDocumentTitle(html As String) As String
''this function will return the document title ie the words between <title> </title>
''html - the html document
'Dim posStart As Long
'Dim posEnd As Long
'posStart = InStr(1, html, "<title>", vbTextCompare) 'this will give us the starting position
'posEnd = InStr(1, html, "</title>", vbTextCompare) 'this will give us the ending position
'
'If posStart = 0 Or posEnd = 0 Then 'a webpage with no title has been downloaded
'    getDocumentTitle = "Unknown"
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

Public Function createRequestHeader(theServer As String, htmlURL As String) As String
Dim strRequestTemplate As String


    

    strRequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
    "Accept: text/html, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
    "Accept-Language: en-ca" & Chr(13) & Chr(10) & _
    "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
    "User-Agent: Mozilla/4.0 (compatible; MSIE 1.0; Windows 3.11; PSC Rulez!!)" & _
    "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
    "Connection: Keep-Alive" & Chr(13) & Chr(10) & _
    "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
    "Host: @$@@$@" & Chr(13) & Chr(10)

'    If strServerAddress = "" Or strDocumentURI = "" Then
'        MsgBox "Unable To detect target page!", vbCritical + vbOK
'        Exit Sub
'    End If

'    If mblnIsProxyUsed Then
'        strServerHostIP = txtProxy.Text
'        mstrRequestHeader = strRequestTemplate
'        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", mstrURL)
'        lngServerPort = 80
'    Else
'        strServerHostIP = strServerAddress
'        lngServerPort = 80
'        mstrRequestHeader = strRequestTemplate
'        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", strDocumentURI)
'    End If

    'we are not worrying about a proxy for now



    createRequestHeader = strRequestTemplate
    createRequestHeader = Replace(createRequestHeader, "_$-$_$-", htmlURL) 'the relative path to the file to be downloaded -if this was proxied then it would be the full url
    createRequestHeader = Replace(createRequestHeader, "@$@@$@", theServer)
    createRequestHeader = createRequestHeader & vbCrLf

End Function

Public Function extractHeader(sRaw As String) As String
'this function will extract the header from a partial download....
'this function will generally be called when a complete header is detected.
'sRAW - is the data that is pulled from the server
Dim posStart As Long, posEnd As Long

'find the vbcrlf & vbcrlf
posEnd = InStr(1, sRaw, vbCrLf & vbCrLf)

posStart = 1
'if posend = posstart then 'there is a problem

extractHeader = Trim$(Left$(sRaw, posEnd - 1))

End Function

Public Function extractHTML(sRaw As String) As String
'this function is used to extract the html portion from
'the raw data sent by the server
Dim posStart As Long
Dim posEnd As Long

posStart = InStr(sRaw, vbCrLf & vbCrLf)

If posStart = 0 Then 'the html was not found
    extractHTML = sRaw 'return the input string as the seperator was not found
    'the error handling could be a little better
    Exit Function
End If

posStart = posStart + Len(vbCrLf & vbCrLf)
posEnd = Len(sRaw) + 1


extractHTML = Mid$(sRaw, posStart, posEnd - posStart)

End Function

Public Function getDocumentTitle(html As String) As String
Dim posStart As Long
Dim posEnd As Long
posStart = InStr(1, html, "<title>", vbTextCompare) 'this will give us the starting position
posEnd = InStr(1, html, "</title>", vbTextCompare) 'this will give us the ending position

If posStart = 0 Or posEnd = 0 Then 'a webpage with no title has been downloaded
    getDocumentTitle = "Unknown Webpage"
    Exit Function
End If

posStart = posStart + Len("<title>")

getDocumentTitle = Mid$(html, posStart, posEnd - posStart) 'now we need to strip out "- visual basic, vb, vbscript"
getDocumentTitle = Left$(getDocumentTitle, InStr(1, getDocumentTitle, "- visual basic, vb, vbscript", vbTextCompare) - 1)
getDocumentTitle = Trim$(stripIllegalChars(getDocumentTitle))
'we need to strip out the special chars in the title if it contains any...


End Function

Private Function stripIllegalChars(sTitle As String) As String
'        ? [ ] / \ = + < > : ; " ,    are illegal characters for file names
'chr$(63) = ?
'chr$(91) = [
'chr$(93) = ]
'chr$(47) = /
'chr$(92) = \
'chr$(61) = =
'chr$(43) = +
'chr$(60) = <
'chr$(62) = >
'chr$(58) = :
'chr$(59) = ;
'chr$(34) = "
'chr$(44) = ,
'chr$(46) = .
'chr$(42) = *
'this funciton takes a string and strips out the illegal chars for a file name
'and replaces them with a space " "
Dim i As Long, upper As Long

upper = Len(sTitle)

For i = 1 To upper

Select Case Mid$(sTitle, i, 1)
    Case Chr$(63), Chr$(91), Chr$(93), Chr$(47), Chr$(92), Chr$(61), Chr$(43), Chr$(60), Chr$(62), Chr$(58), Chr$(59), Chr$(34), Chr$(44), Chr$(46), Chr$(42)
        Mid$(sTitle, i, 1) = " " 'replace the illegal char with a space
End Select
Next 'i
stripIllegalChars = sTitle
End Function


