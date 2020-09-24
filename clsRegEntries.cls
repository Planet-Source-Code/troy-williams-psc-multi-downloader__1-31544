VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'I can't remember where I got this from, but I believe it is from the
'Total VB SourceBook 5 software. I also believe that I had modified the code
'from the original (slightly). I just don't remember.



' Class       : CProgramSettings
' Description : Simplified registry access routines used
'               for saving program settings
' Source      :
'

' Declarations for Windows API calls
Private Declare Function RegCloseKey _
  Lib "advapi32.dll" _
  (ByVal hKey As Long) _
  As Long

Private Declare Function RegCreateKeyEx _
  Lib "advapi32.dll" _
  Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, _
   ByVal lpSubKey As String, _
   ByVal Reserved As Long, _
   ByVal lpClass As String, _
   ByVal dwOptions As Long, _
   ByVal samDesired As Long, _
   ByVal lpSecurityAttributes As Long, _
   phkResult As Long, _
   lpdwDisposition As Long) _
  As Long

Private Declare Function RegOpenKeyEx _
  Lib "advapi32.dll" _
  Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   phkResult As Long) _
  As Long
  
Private Declare Function RegQueryValueExString _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   ByVal lpData As String, _
   lpcbData As Long) _
  As Long

Private Declare Function RegQueryValueExLong _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   lpData As Long, _
   lpcbData As Long) _
  As Long

Private Declare Function RegQueryValueExNULL _
  Lib "advapi32.dll" _
  Alias "RegQueryValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   lpType As Long, _
   ByVal lpData As Long, _
   lpcbData As Long) _
  As Long

Private Declare Function RegSetValueExString _
  Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal Reserved As Long, _
   ByVal dwType As Long, _
   ByVal lpValue As String, _
   ByVal cbData As Long) _
  As Long

Private Declare Function RegSetValueExLong _
  Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
  (ByVal hKey As Long, _
   ByVal lpValueName As String, _
   ByVal Reserved As Long, _
   ByVal dwType As Long, _
   lpValue As Long, _
   ByVal cbData As Long) _
  As Long

' Constants for Windows API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
 
Private Const KEY_ALL_ACCESS = &H3F
Private Const KEY_QUERY_VALUE = &H1
 
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4

' Public property enumerated constants
Public Enum EnumProgramSettingsRoot
  psrHKEY_CLASSES_ROOT = &H80000000
  psrHKEY_CURRENT_USER = &H80000001
  psrHKEY_LOCAL_MACHINE = &H80000002
  psrHKEY_USERS = &H80000003
End Enum

' Local variables to hold Public Property values
Private m_lngRootKey As EnumProgramSettingsRoot
Private m_strMainBranch As String
Private m_strRegBase As String
Private m_strProgram As String
Private m_strSection As String

Private Sub Class_Initialize()
  ' Set initial values to defaults which may be overridden
  ' with property settings
  ' Source: Total VB SourceBook 5

  m_lngRootKey = psrHKEY_LOCAL_MACHINE
  m_strMainBranch = "SOFTWARE"
  m_strSection = "Settings"

End Sub

Public Property Get MainBranch() As String
  ' Returns: the current value of MainBranch
  ' Source: Total VB SourceBook 5

  MainBranch = m_strMainBranch

End Property

Public Property Let MainBranch(strValue As String)
  ' strMainBranch: Set the value for the MainBranch property
  ' Source: Total VB SourceBook 5
  
  m_strMainBranch = strValue

End Property

Public Property Get Program() As String
  ' Returns: the current value of the Program property
  ' Source: Total VB SourceBook 5
  
  Program = m_strProgram

End Property

Public Property Let Program(strValue As String)
  ' strValue: Set the value for the Program property
  ' Source: Total VB SourceBook 5
  
  m_strProgram = strValue

End Property

Public Property Get RegBase() As String
  ' Returns: the current value of RegBase
  ' Source: Total VB SourceBook 5
  
  RegBase = m_strRegBase

End Property

Public Property Let RegBase(strValue As String)
  ' strBase: Set the value for the RegBase property
  ' Source: Total VB SourceBook 5
  
  m_strRegBase = strValue

End Property

Public Property Get RootKey() As EnumProgramSettingsRoot
  ' Returns: the current value of RootKey
  ' Source: Total VB SourceBook 5
  
  RootKey = m_lngRootKey

End Property

Public Property Let RootKey(eValue As EnumProgramSettingsRoot)
  ' eValue: Set the value for the RootKey property
  ' Source: Total VB SourceBook 5
  
  m_lngRootKey = eValue

End Property

Public Property Get Section() As String
  ' Returns: the current value of the Section property
  ' Source: Total VB SourceBook 5
  
  Section = m_strSection

End Property

Public Property Let Section(strValue As String)
  ' strValue: Set the value for the Section property
  ' Source: Total VB SourceBook 5
  
  m_strSection = strValue

End Property

Public Function ReadEntry( _
  strEntry As String, _
  strDefault As String) _
  As String
  ' Comments  : Reads a string value from the location in the
  '             registry specified by the class properties
  ' Parameters: strEntry - The value to retrieve
  '             strDefault - The value to return if the entry
  '             is not found
  ' Returns   : Either the registry value, or the default value
  ' Source    : Total VB SourceBook 5
  '
  Dim strValue As String
  Dim strSearch As String
  
  On Error GoTo PROC_ERR
  
  strSearch = m_strMainBranch & "\" & _
    m_strRegBase & "\" & _
    m_strProgram & "\" & _
    m_strSection
  
  strValue = GetKeyValue(m_lngRootKey, strSearch, strEntry)
  
  If strValue = "" Then
    ReadEntry = strDefault
  Else
    ReadEntry = strValue
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReadEntry"
  Resume PROC_EXIT

End Function

Public Sub WriteEntry( _
  strEntry As String, _
  strValue As String)
  ' Comments  : Writes a string entry into the registry
  '             at the location specified by the class properties
  ' Parameters: strEntry - the string entry key value
  '             strValue - the value of the entry
  '
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 5
  '
  Dim strSearch As String
  
  On Error GoTo PROC_ERR
  
  strSearch = m_strMainBranch & "\" & _
    m_strRegBase & "\" & _
    m_strProgram & "\" & _
    m_strSection
    
  SetKeyValue m_lngRootKey, strSearch, strEntry, strValue, REG_SZ

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "WriteEntry"
  Resume PROC_EXIT

End Sub

Private Function GetKeyValue( _
  eValue As EnumProgramSettingsRoot, _
  strKeyName As String, _
  strValueName As String) _
  As Variant
  ' Comments  : Retrieves the specified key value
  ' Parameters: eValue - value indicating the root key value
  '             strKeyName - The name of the key to open
  '             strValueName - the value to open; vbNullString
  '             opens the default value
  ' Returns   : The value, if found, otherwise null
  ' Source    : Total VB SourceBook 5
  '
  Dim lngReturnValue As Long      'The result of the API functions
  Dim lngHKey As Long         'The handle of opened key
  Dim varValue As Variant    'The setting of queried value
  Dim strValueData As String 'Used if the value is a string data type
  Dim lngValueData As Long   'Used if the value is a long data type
  Dim lngValueType As Long   'The data type of the value
  Dim lngDataSize As Long    'The size of the data
 
  On Error GoTo PROC_ERR
 
  'Initialize return variable
  varValue = Empty
  
  lngReturnValue = RegOpenKeyEx( _
    eValue, _
    strKeyName, _
    0&, _
    KEY_ALL_ACCESS, _
    lngHKey)
    
  If ERROR_NONE = lngReturnValue Then
    
    lngReturnValue = RegQueryValueExNULL( _
      lngHKey, _
      strValueName, _
      0&, _
      lngValueType, _
      0&, _
      lngDataSize)
    If ERROR_NONE = lngReturnValue Then
      Select Case lngValueType
        ' Strings type
        Case REG_SZ:

          strValueData = String(lngDataSize, 0)
          
          lngReturnValue = RegQueryValueExString( _
            lngHKey, _
            strValueName, _
            0&, _
            lngValueType, _
            strValueData, _
            lngDataSize)
          If ERROR_NONE = lngReturnValue Then
                          
            If Len(strValueData) Then
              If Mid$(strValueData, lngDataSize, 1) = vbNullChar Then
                lngDataSize = lngDataSize - 1
              End If
              varValue = Left$(strValueData, lngDataSize)
            Else
              varValue = ""
            End If
            
          Else
            varValue = Empty
          End If
         
        ' Long type
        Case REG_DWORD:
          lngReturnValue = RegQueryValueExLong( _
            lngHKey, _
            strValueName, _
            0&, _
            lngValueType, _
            lngValueData, _
            lngDataSize)
          If ERROR_NONE = lngReturnValue Then
            varValue = lngValueData
          End If
          
        Case Else
          'No other data types supported
          lngReturnValue = True
      End Select
        
    End If
        
    RegCloseKey (lngHKey)
    
  End If
    
  'Return varValue
  GetKeyValue = varValue

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetKeyValue"
  Resume PROC_EXIT

End Function

Private Sub SetKeyValue( _
  eValue As EnumProgramSettingsRoot, _
  strKeyName As String, _
  strValueName As String, _
  varValue As Variant, _
  lngValueType As Long)
  ' Comments  : Sets a registry value Value. Will create key
  '             if it doesn't exist
  ' Parameters: eValue - The root key value
  '             strKeyName - The name of the key to open
  '             strValueName - The value to open, vbNullString will open the default value
  '             varValue - The data to assign to the value
  '             lngValueType - The data type of the value
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 5
  '
  Dim lngReturnValue As Long
  Dim lngHKey As Long
       
  On Error GoTo PROC_ERR
  
  ' Open the specified key; if it does not exist then create it
  lngReturnValue = RegCreateKeyEx( _
    eValue, _
    strKeyName, _
    0&, _
    vbNullString, _
    REG_OPTION_NON_VOLATILE, _
    KEY_ALL_ACCESS, _
    0&, _
    lngHKey, _
    0&)
  
  'Determine the data type of the key
  Select Case lngValueType
    Case REG_SZ
      varValue = varValue & vbNullChar
      
      lngReturnValue = RegSetValueExString( _
        lngHKey, _
        strValueName, _
        0&, _
        lngValueType, _
        varValue, _
        Len(varValue))
    
    Case REG_DWORD
      lngReturnValue = RegSetValueExLong( _
        lngHKey, _
        strValueName, _
        0&, _
        lngValueType, _
        CLng(varValue), _
        REG_DWORD)
  End Select
  
  RegCloseKey (lngHKey)
  

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SetKeyValue"
  Resume PROC_EXIT

End Sub

'example usage

'Private reg As clsRegEntries

'Set reg = New clsRegEntries
'With reg
'    .RootKey = psrHKEY_LOCAL_MACHINE
'    .MainBranch = "SOFTWARE"
'    .RegBase = "BlueBill"
'    .Program = "pscDownloader"
'    .Section = "Info"
'End With
'the code above points it to HKLM/SOFTWARE/BlueBill/pscDownloader/Info

'dbpath = reg.ReadEntry("DBpath", "")
'reg.WriteEntry "DBpath", txtDBpath.Text
'DBpath is a variable that holds a value
