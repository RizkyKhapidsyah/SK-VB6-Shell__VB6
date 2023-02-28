Attribute VB_Name = "Shell"
'File: vbshell.bas
' Copyright 1998 Andrew S. Dean

Option Explicit


' For adding files to the Recent Documents menu.
Declare Sub SHAddToRecentDocs Lib "Shell32" (ByVal uFlags As Long, ByVal lpBuffer As String)

'Global Const SHARD_PIDL = 1
Global Const SHARD_PATHA = 2
Global Const SHARD_PATHW = 3


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' For adding files to the Recycle Bin, etc.

Type SHFILEOPSTRUCT
        hwnd   As Long
        wFunc  As Long
        pFrom  As String
        pTo    As String
        fFlags As Integer
        fAnyOperationsAborted As Boolean
        hNameMappings         As Long
        lpszProgressTitle     As String '  only used if FOF_SIMPLEPROGRESS
End Type

Declare Function SHFileOperation Lib "Shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

' SHFileOperation wFunc settings
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4

' SHFileOperation fFlag settings
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_WANTMAPPINGHANDLE = &H20


''''''''''''''''''''''''''''''''''''''''''''''''''
' Registry functions
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001

Public Const REG_SZ = 1   ' String data type


Public Const SYNCHRONIZE = &H100000
' Reg Key Security Options
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = (KEY_READ)
Public Const STANDARD_RIGHTS_ALL = &H1F0000

Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Declare Function RegCreateKey Lib "advapi32.dll" _
      Alias "RegCreateKeyA" _
      (ByVal hKey As Long, ByVal lpctstr As String, _
      phkey As Long) As Long
      
Declare Function RegCloseKey Lib "advapi32.dll" _
      (ByVal hKey As Long) As Long
      
Declare Function RegSetValueEx Lib "advapi32.dll" _
      Alias "RegSetValueExA" _
      (ByVal hKey As Long, ByVal lpValueName As String, _
      ByVal Reserved As Long, ByVal dwType As Long, _
      lpData As Any, ByVal cbData As Long) As Long
      
Declare Function RegDeleteKey Lib "advapi32.dll" _
      Alias "RegDeleteKeyA" _
      (ByVal hKey As Long, ByVal lpszSubkey As String) _
      As Long
            
Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubkey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
      Alias "RegQueryValueExA" _
      (ByVal hKey As Long, ByVal lpszValueName As String, _
      ByVal lpdwReserved As Long, lpdwType As Long, _
      lpData As Any, lpcbData As Long) As Long
' Definition of lpdwReserved modified by adding BYVAL


''''''''''''''''''''''''''''''''''''''''''''''''''
' From VB5 Setup Kit
Declare Function OSfCreateShellLink Lib "VB5STKIT.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function OSfCreateShellGroup Lib "VB5STKIT.DLL" Alias "fCreateShellFolder" (ByVal lpstrDirName As String) As Long
Declare Function OSfRemoveShellLink Lib "VB5STKIT.DLL" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long



Sub Main()
  
  If Command() <> "" Then
          
    Dim CurChar As String
    Dim CmdLine As String
    Dim CmdLineLen As Long
    Dim NumArgs As Integer
    Dim InArg As Integer
    Dim PosInStr As Integer
    Dim strMsg As String
    Dim ArgArray() As String
    
    'Get command line arguments.
    CmdLine = Command()
    CmdLineLen = Len(CmdLine)
    
    ' Initialize counters and flags
    NumArgs = 0
    InArg = False
    
    ReDim ArgArray(NumArgs)
    
    ' Go thru command line one character at a time.
    ' We assume that a Space or Tab character is used as the delimiter.
    For PosInStr = 1 To CmdLineLen
        CurChar = Mid(CmdLine, PosInStr, 1)

        'Test for space or tab.
        If (CurChar <> " " And CurChar <> vbTab) Then
            'Neither space nor tab. Test if already building argument.
            If Not InArg Then
                'Begin new argument.
                NumArgs = NumArgs + 1
                InArg = True
            End If
            'Add character to end of current argument.
            ArgArray(NumArgs - 1) = ArgArray(NumArgs - 1) & CurChar
        Else
            'Found a space or tab. Set InArg flag to False.
            InArg = False
            ReDim Preserve ArgArray(NumArgs)
        End If
    Next PosInStr
      
    Dim I As Integer
    
    For I = 0 To NumArgs - 1
      strMsg = strMsg & "> " & ArgArray(I) & vbCrLf
    Next I
    
    MsgBox "Command line arguments: " & vbCrLf & strMsg
  
  End If
  
  frmVBShell.Show

End Sub

' Move a file to the recycle bin.
Function RemoveFile(strFile As String) As Long
  
  Dim SHFileOp As SHFILEOPSTRUCT
  
  With SHFileOp
     .wFunc = FO_DELETE
     .pFrom = strFile
     .fFlags = FOF_ALLOWUNDO
  End With
  
  RemoveFile = SHFileOperation(SHFileOp)
  
End Function


' Copy a file, displaying a progress window.
Function CopyFile(strFileOld As String, strFileNew As String) As Long
  
  Dim SHFileOp As SHFILEOPSTRUCT
  
  With SHFileOp
     .wFunc = FO_COPY
     .pFrom = strFileOld
     .pTo = strFileNew
  End With
  
  CopyFile = SHFileOperation(SHFileOp)

End Function


Sub AddToRecentDocs(strFile As String)

   On Error Resume Next
   
   ' Win 95 does not use UNICODE.  NT uses UNICODE by default.
   ' VB always uses UNICODE internally, but converts as necessary.
   
   ' This is great because it doesn't add duplicates and it
   ' does not add items that are not valid file names.
   
   ' It appears that this doesn't work if File Type has
   ' not been defined for the file extension (ie, the
   ' default value of the AppID key has not be set.
   
   SHAddToRecentDocs SHARD_PATHA, ByVal strFile

End Sub


Sub DeleteKey(szKey As String)
   
   Dim lResult As String
   Dim hKey    As Long
   
   If szKey <> "" Then
      lResult = RegDeleteKey(HKEY_CLASSES_ROOT, szKey)
   End If
      
End Sub


''''''''''''''''''''''''
' This routine sets up a file association so that a file
' can be opened by an application by double clicking on the
' file in Explorer.
' A file to use as an empty file will also be registered,
' so that a new file of this type can be created by clicking
' on the new menu in Explorer, on the desktop, etc.
'''''''''''''''''''''''''
Sub SetFileAssociation(strExt As String, strAppID As String, strCommand As String, strEmptyFile As String, strFileLabel As String, strIcon As String)

  ' strExt is the file extension
  ' strAppID is the Application Identifier.
  ' strCommand is the Open Command
  ' strEmptyFile is the file to use to create new files.
  ' strFileLabel is the string displayed in the various New menus.
  
  ' We want to create
  '  .ext -> AppID
  '      ShellNew
  '            FileName -> strNewValue
  '  AppID -> FileLabel
  ' and then some...

  Dim lResult As Long
  Dim hKey    As Long
  Dim strValueName As String
  
  '' IT APPEARS FROM LOOKING AT OTHER ENTRIES THAT LONG FILE NAMES
  '' MIGHT NOT BE VALID IN THE COMMAND?  TEST THIS!!!
  
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strExt, hKey)
  Debug.Assert lResult = 0
  
  'strValueName = ""
  ' lResult = RegSetValueEx(hKey, strValueName, 0, REG_SZ, ByVal strAppID, Len(strAppID))
  lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strAppID, Len(strAppID))
  Debug.Assert lResult = 0
  

  If strEmptyFile <> "" Then
     Dim strKey As String
     strKey = strExt & "\" & "ShellNew"
     lResult = RegCreateKey(HKEY_CLASSES_ROOT, strKey, hKey)
     Debug.Assert lResult = 0
  
     ' This could be either FileName, Command, or Data
     ' It should be an argument of the function.
     strValueName = "FileName"
     lResult = RegSetValueEx(hKey, strValueName, 0, REG_SZ, ByVal strEmptyFile, Len(strEmptyFile))
     Debug.Assert lResult = 0
  
     lResult = RegCloseKey(hKey)
     Debug.Assert lResult = 0
  End If
      
  
  '' HOW IMPORTANT IS IT TO CLOSE THE KEY????
  Dim strTemp As String
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strAppID, hKey)
  
  If strFileLabel <> "" Then
     'strValueName = ""
     lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strFileLabel, Len(strFileLabel))
  End If
  lResult = RegCloseKey(hKey)
  
  strTemp = strAppID & "\shell\open\command"
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strTemp, hKey)
  'strValueName = ""
  'lResult = RegSetValueEx(hKey, strValueName, 0, REG_SZ, ByVal strCommand, Len(strCommand))
  lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strCommand, Len(strCommand))
  lResult = RegCloseKey(hKey)
  Debug.Assert lResult = 0
  
  
  ' Register the default Icon
  If strIcon <> "" Then
    strTemp = strAppID & "\DefaultIcon"
    lResult = RegCreateKey(HKEY_CLASSES_ROOT, strTemp, hKey)
    strValueName = ""
    ' If the icon was passed in as a number, assume the
    ' DefaultIcon is supposed to be "this.exe,1"
    ' Otherwise, assume the entire file and icon number was used.
    If IsNumeric(strIcon) Then
      strTemp = App.Path & "\" & App.EXEName & ".exe," & strIcon
    Else
      strTemp = strIcon
    End If
    lResult = RegSetValueEx(hKey, strValueName, 0, REG_SZ, ByVal strTemp, Len(strTemp))
    Debug.Assert lResult = 0
    lResult = RegCloseKey(hKey)
    Debug.Assert lResult = 0
  End If
      

End Sub


Function GetShellFolder(szFolder As String) As String
  
  Dim lResult   As Long
  Dim strKey    As String
  Dim hKey      As Long
  Dim strBuffer As String
  Dim lLen      As Long
  
  ' A better approach than this (language independent, for example), would be to use
  ' the SHGetSpecialFolderLocation() function, and pass the appropriate CSIDL constant.
  ' CSIDL constants are defined in shlobj.h, a Windows Header File.
  '
  
  strKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  
'  Public Const KEY_QUERY_VALUE = &H1
  
  lResult = RegOpenKeyEx(HKEY_CURRENT_USER, strKey, 0, KEY_QUERY_VALUE, hKey)
  If lResult <> 0 Then
     GetShellFolder = ""
     Exit Function
  End If
  
  strBuffer = Space$(1024)
  lLen = Len(strBuffer)
  lResult = RegQueryValueEx(hKey, szFolder, 0, REG_SZ, ByVal strBuffer, lLen)
  If lResult <> 0 Then
     GetShellFolder = ""
     Exit Function
  End If
  
  ' Might still want to verify that lLen > 0
  GetShellFolder = Left$(strBuffer, lLen - 1)
    
End Function


Function GetFileAssociation(strExt As String) As String

  Dim hKey      As Long
  Dim strBuffer As String
  Dim strTemp   As String
  Dim lResult   As Long
  
  lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strExt, 0, KEY_READ, hKey)
  If lResult <> 0 Then
     GetFileAssociation = "Unregistered file extension " & strExt
     Exit Function
  End If
  
  Dim lLen As Long
  
  Dim strValueName As String
  strValueName = ""
  Dim lType As Long
  lType = REG_SZ
  strBuffer = Space$(128)
  lLen = Len(strBuffer)
  
  lResult = RegQueryValueEx(hKey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
  
  Debug.Print lResult
  Debug.Print lType
  Debug.Print lLen
  
  'lResult = RegQueryValue(hKey, strValueName, ByVal strBuffer, lLen)
  'lResult = RegQueryValue(hKey, ByVal strValueName, ByVal 0, lLen)
  If lResult <> 0 Then
     MsgBox lResult
     Exit Function
  End If
  strTemp = Mid$(strBuffer, 1, lLen - 1)
  Debug.Print strTemp
  
  strTemp = strTemp & "\shell\open\command"
  
  lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, strTemp, 0, KEY_READ, hKey)
  If lResult <> 0 Then
     GetFileAssociation = "File type is not associated with a program."
     Exit Function
  End If
  
  lLen = Len(strBuffer)
  strValueName = ""
  lResult = RegQueryValueEx(hKey, strValueName, 0, REG_SZ, ByVal strBuffer, lLen)
  strTemp = Mid$(strBuffer, 1, lLen - 1)
  
  GetFileAssociation = strTemp
  
End Function


''''''''''''''''''
' From VB5 Setup Kit:

'-----------------------------------------------------------
' SUB: CreateShellLink
'
' Creates (or replaces) a link in either Start>Programs or
' any of its immediate subfolders in the Windows 95 shell.
'
' IN: [strLinkPath] - full path to the target of the link
'                     Ex: 'c:\Program Files\My Application\MyApp.exe"
'     [strLinkArguments] - command-line arguments for the link
'                     Ex: '-f -c "c:\Program Files\My Application\MyApp.dat" -q'
'     [strLinkName] - text caption for the link
'     [fLog] - Whether or not to write to the logfile (default
'                is true if missing)
'
' OUT:
'   The link will be created in the folder strGroupName

' You can edit these manually with Explorer in the
' Windows\StartMenu\ directory.

'-----------------------------------------------------------

Sub CreateShellLink(ByVal strGroupName As String, ByVal strLinkName As String, ByVal strLinkPath As String, ByVal strLinkArguments As String)
    
    strLinkName = strUnQuoteString(strLinkName)
    strLinkPath = strUnQuoteString(strLinkPath)
    
    Dim fSuccess As Boolean
    
    fSuccess = OSfCreateShellLink(strGroupName, strLinkName, strLinkPath, strLinkArguments) 'the path should never be enclosed in double quotes
    
    If Not fSuccess Then
       MsgBox "Couldn't create link"
    End If

End Sub



'-----------------------------------------------------------
' SUB: fCreateShellGroup
'
' Creates a new program group off of Start>Programs in the
' Windows 95 shell if the specified folder doesn't already exist.
'
'-----------------------------------------------------------
Function fCreateShellGroup(ByVal strFolderName As String) As Boolean
    
    ReplaceDoubleQuotes strFolderName
    
    If strFolderName = "" Then
        Exit Function
    End If
    
    Dim fSuccess As Boolean
    
    fSuccess = OSfCreateShellGroup(strFolderName)
    fCreateShellGroup = fSuccess
    
End Function

'-----------------------------------------------------------
' SUB: RemoveShellLink
'
' Removes a link in either Start>Programs or any of its

' immediate subfolders in the Windows 95 shell.
'
' IN: [strFolderName] - text name of the immediate folder
'                       in which the link to be removed
'                       currently exists, or else the
'                       empty string ("") to indicate that
'                       the link can be found directly in
'                       the Start>Programs menu.
'     [strLinkName] - text caption for the link
'
' This action is never logged in the app removal logfile.
'
' PRECONDITION: strFolderName has already been created and is
'               an immediate subfolder of Start>Programs, if it
'               is not equal to ""
'-----------------------------------------------------------
'
Sub RemoveShellLink(ByVal strFolderName As String, ByVal strLinkName As String)
    
    Dim fSuccess As Boolean
    
    ReplaceDoubleQuotes strFolderName
    ReplaceDoubleQuotes strLinkName
    
    
    fSuccess = OSfRemoveShellLink(strFolderName, strLinkName)
End Sub

' Replace all double quotes with single quotes
Public Sub ReplaceDoubleQuotes(str As String)
    
    Dim I As Integer
    
    For I = 1 To Len(str)
        If Mid$(str, I, 1) = """" Then
            Mid$(str, I, 1) = "'"
        End If
    Next I
    
End Sub


Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim(strQuotedString)
    Dim strQUOTE As String
    
    strQUOTE = """"

    If Mid$(strQuotedString, 1, 1) = strQUOTE And Right$(strQuotedString, 1) = strQUOTE Then
        '
        ' It's quoted.  Get rid of the quotes.
        '
        strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
    End If
    
    strUnQuoteString = strQuotedString
    
End Function

