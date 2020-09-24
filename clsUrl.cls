VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This class is designed to hold an url and a boolean

Private myURL As String
Private isDownloaded As Boolean 'this is used to indicate whether this url
'has been downloaded or not.



Public Property Get url() As String
url = myURL
End Property

Public Property Let url(ByVal s As String)
myURL = s
End Property

Public Property Get DownLoaded() As Boolean
DownLoaded = isDownloaded
End Property

Public Property Let DownLoaded(ByVal b As Boolean)
isDownloaded = b
End Property
