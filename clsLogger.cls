VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private logName As String
Private logSize As Long 'in kb
Private logPath As String

Private Sub Class_Initialize()
logName = ""
logSize = 0
logPath = ""


End Sub

Public Property Let setLogName(name As String)
    logName = name
End Property

Public Property Let setLogSize(size As Long)
    logSize = size
End Property

Public Property Let setLogPath(Path As String)
    logPath = Path
End Property

Public Property Get getLogName() As String
    getLogName = logName
End Property

Public Property Get getLogSize() As Long
    getLogSize = logSize
End Property

Public Property Get getLogPath() As String
    getLogPath = logPath
End Property


Public Sub logger(message As String)
'The purpose of this routine is to append the message to the log file specified in logName

'check to make sure that the logname, logpath and logsize
If logName = "" Or logSize = 0 Then
    MsgBox "Please set the name and/or size of the log file!", vbExclamation, "Error"
    Exit Sub
End If
    Dim logNamePath As String
    If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    logNamePath = logPath & logName


    Dim fso As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    
    On Error Resume Next
        
    'check to see if the log file exists
    If fso.FileExists(logNamePath) Then 'then set the file object to it
        Set fil = fso.GetFile(logNamePath)
    Else 'else create the file
        Set ts = fso.CreateTextFile(logNamePath)
        Set fil = fso.GetFile(logNamePath)
    End If
     
        
    'MsgBox fil.Name & " " & fil.Size
    'The if statement below should be configurable in the settings program...
    If fil.size < (CInt(logSize) * 1000) Then 'is bigger then errorlogsize in bytes then make a backup
    'set the textstream to logName. if logname doesn't exist, then create it.
    Set ts = fso.OpenTextFile(logNamePath, ForAppending, True)
    ts.WriteLine (message)
    ts.Close
    Else 'rename the file and write to a new one
    'this section works.....
        Dim X As Integer
        For X = 1 To 30000
            If Not fso.FileExists(logNamePath & "." & X) Then Exit For
        Next X
        
        fil.Copy (logNamePath & "." & X)
        fil.Delete
        logger message
'        Set ts = fso.OpenTextFile(logName, ForAppending, True)
'        ts.WriteLine (message)
'        ts.Close
        
    End If

End Sub
