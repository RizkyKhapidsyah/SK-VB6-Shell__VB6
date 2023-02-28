VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileAssociation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'File: CFileAss.cls
'      CFileAssociation class for VB
' Copyright 1998 Andrew S. Dean

Option Explicit

' Public properties for the FileAssociation Class.
Public strExt               As String  ' .txt
Public strAppID             As String  ' txtfile
Public strOpenCommand       As String  ' C:\windows\notepad.exe %1
Public strFileType          As String  ' Text Document
Public strIcon              As String  ' c:\windows\notepad.exe,0
Public strNewFileType       As String  ' NullFile, FileName, Command (or Data)
Public strNewFileCommand    As String
Public strExePath           As String  ' c:\windows\notepad.exe




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Creates
'  HKEY_CLASSES_ROOT
'  .ext -> AppID
'      ShellNew
'            FileName -> strNewValue
'  AppID -> FileLabel
'      shell
'            open
'                  command -> strOpenCommand
' and then some...
'
' Of course robust error handling should be added!
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateAssociation() As Long
  
  If strExt = "" Or strAppID = "" Then
     CreateAssociation = 1
     Exit Function
  End If
  

  Dim lResult      As Long
  Dim hKey         As Long
  Dim strValueName As String
  Dim strKey       As String
  
  ' Create
  ' HKEY_CLASSES_ROOT
  '   .ext -> AppID
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strExt, hKey)
  lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strAppID, Len(strAppID))
  
  
  ' Create
  ' HKEY_CLASSES_ROOT
  '    .ext
  '         ShellNew
  '                FileName  -> strNewFile
  
  ' Don't handle the Date file type yet...
  If strNewFileType = "NullFile" Or strNewFileType = "Command" Or strNewFileType = "FileName" Then
     
     strKey = strExt & "\" & "ShellNew"
     lResult = RegCreateKey(HKEY_CLASSES_ROOT, strKey, hKey)
     
     ' This could be either NullFile, FileName, Command, or Data
     strValueName = strNewFileType
     If strValueName = "NullFile" Then
        strNewFileCommand = ""
     End If
     lResult = RegSetValueEx(hKey, strValueName, 0, REG_SZ, ByVal strNewFileCommand, Len(strNewFileCommand))
     lResult = RegCloseKey(hKey)
  End If
  
  
  ' Create
  ' HKEY_CLASSES_ROOT
  '    AppID  -> FileLabel
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strAppID, hKey)
  If strFileType <> "" Then
      lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strFileType, Len(strFileType))
  End If
  lResult = RegCloseKey(hKey)
  
  
  ' this block should be a separate routine, to make it easy
  ' to add multiple commands.
  ' Create
  ' HKEY_CLASSES_ROOT
  '     AppID
  '         shell
  '            open
  '               command -> strOpenCommand
'  strKey = strAppID & "\shell\open\command"
'  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strKey, hKey)
'  lResult = RegSetValueEx(hKey, "", 0, REG_SZ, ByVal strOpenCommand, Len(strOpenCommand))
'  lResult = RegCloseKey(hKey)
  
  Me.CreateContextMenuItem "open", strOpenCommand
  
  
  ' Register the default Icon
  ' Should we make the icon exe,0 by default?
  If strIcon <> "" Or strExePath <> "" Then
    strKey = strAppID & "\DefaultIcon"
    lResult = RegCreateKey(HKEY_CLASSES_ROOT, strKey, hKey)
    
    ' If the icon was passed in as a number, assume the
    ' DefaultIcon is supposed to be "this.exe,1"
    ' Otherwise, assume the entire file and icon number was used.
    
    Dim strTemp As String
    If strIcon = "" Then
       strTemp = strExePath    ' If no icon was specified, use the default icon from the exe.
    ElseIf IsNumeric(strIcon) Then
      strTemp = App.Path & "\" & App.EXEName & ".exe," & strIcon
    Else
      strTemp = strIcon
    End If
    
    lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strTemp, Len(strTemp))
    lResult = RegCloseKey(hKey)
  End If
  
  CreateAssociation = 0

End Function


'''''''''''''''''''''''''''''''''''''''''''
' Create a context menu for an existing file association.
' Typical call would be
'    .CreateContextMenuItem( "open", "c:\windows\notepad.exe %1"
'    .CreateContextMenuItem( "print", "c:\windows\notepad.exe /p %1"
'''''''''''''''''''''''''''''''''''''''''''
Public Sub CreateContextMenuItem(strText As String, strCommand As String)

  Dim strKey  As String
  Dim hKey    As Long
  Dim lResult As Long
  
  Debug.Assert strAppID <> ""
  
  strKey = strAppID & "\shell\" & strText & "\command"
  lResult = RegCreateKey(HKEY_CLASSES_ROOT, strKey, hKey)
  lResult = RegSetValueEx(hKey, vbNullString, 0, REG_SZ, ByVal strCommand, Len(strCommand))
  lResult = RegCloseKey(hKey)
  
End Sub


''''''''''''''''''''''''''''''''''''''''''''
' Delete an existing file association
' with minimal error checking
' Note that on NT, this will fail because
' NT does not cascade deletes.  You will need to
' query subkeys, and delete them before deleting
' a key.
''''''''''''''''''''''''''''''''''''''''''''
Public Function DeleteAssociation() As Long

  If strExt = "" Or strAppID = "" Then
     DeleteAssociation = 1
     Exit Function
  End If
  Dim lResult As Long
  
  On Error GoTo DeleteAssociation_EH
  
  lResult = RegDeleteKey(HKEY_CLASSES_ROOT, strExt)
  lResult = RegDeleteKey(HKEY_CLASSES_ROOT, strAppID)
  
  DeleteAssociation = 0
  
  Exit Function
  
  
DeleteAssociation_EH:
  MsgBox "Deleting the keys for " & strExt & " and " & strAppID & " failed.  If you are running NT, you will need to delete all child keys before deleting a parent key."
  
End Function


