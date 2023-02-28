VERSION 5.00
Begin VB.Form frmVBShell 
   Caption         =   "VBShell Integration Demo"
   ClientHeight    =   8190
   ClientLeft      =   3075
   ClientTop       =   795
   ClientWidth     =   6645
   Icon            =   "vbshell.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   6645
   Begin VB.Frame Frame3 
      Caption         =   "Notepad"
      Height          =   2055
      Left            =   3480
      TabIndex        =   26
      Top             =   6000
      Width           =   3015
      Begin VB.CommandButton cmdAddNotepadShell 
         Caption         =   "&Add To Start\Programs\VB Shell menu"
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddNotepadStartup 
         Caption         =   "Add To Start&Up"
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.ComboBox cboShellFolderName 
      Height          =   315
      Left            =   480
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   3015
      Begin VB.TextBox txtFilename 
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Text            =   "c:\junk.txt"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdJunkIt 
         Caption         =   "&Copy to c:\junk.old"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdRecycleFile 
         Caption         =   "&Move file to recycle bin"
         Height          =   375
         Left            =   600
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdGetShellFolder 
      Caption         =   "Get Shell &Folder"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtRecentDoc 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton cmdAddRecentDocs 
      Caption         =   "Add To &Recent Docs"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Add a document to Document in the Start Menu"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Associations"
      Height          =   4215
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   360
         TabIndex        =   8
         Text            =   ".txt"
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdGetAssoc 
         Caption         =   "&Get Association"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeleteAssoc 
         Caption         =   "&Delete Association"
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtEmptyFile 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtFileLabel 
         Height          =   315
         Left            =   3240
         TabIndex        =   17
         Text            =   "ZZZ demo file"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtOpen 
         Height          =   315
         Left            =   360
         TabIndex        =   15
         Text            =   "c:\windows\notepad.exe  %1"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtID 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Text            =   "zzzfile"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CommandButton cmdSetAssoc 
         Caption         =   "&Set Association"
         Height          =   375
         Left            =   1200
         TabIndex        =   20
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtExt 
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Text            =   ".zzz"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "File extension (including period)"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Empty File"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "File Label"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Open"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "ID"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "File Extension (including period)"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Shell Folder"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Document Name"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmVBShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' File: vbshell.frm
' Copyright 1998 Andrew S. Dean

Option Explicit


Private Sub cmdAddNotepadShell_Click()

   fCreateShellGroup "VBShell Demo"
   CreateShellLink "VBShell Demo", "Run Notepad", "C:\windows\notepad.exe", ""

End Sub

Private Sub cmdAddNotepadStartup_Click()

   ' Add the file to start when Windows starts.
   CreateShellLink "StartUp", "Run Notepad", "C:\windows\notepad.exe", ""

End Sub

Private Sub cmdAddRecentDocs_Click()

   AddToRecentDocs txtRecentDoc.Text

End Sub

Private Sub cmdDeleteAssoc_Click()

  Dim clsFileAssoc As New CFileAssociation
  
  With clsFileAssoc
     .strExt = txtExt
     .strAppID = txtID
  End With
  
  clsFileAssoc.DeleteAssociation
  
End Sub

Private Sub cmdGetAssoc_Click()

  Dim strExt As String
  
  strExt = Text1.Text
  
  MsgBox GetFileAssociation(strExt)

End Sub

Private Sub cmdGetShellFolder_Click()

  Dim strTemp As String
  
  strTemp = GetShellFolder(cboShellFolderName.Text)
  
  MsgBox strTemp

End Sub

Private Sub cmdJunkIt_Click()

   CopyFile txtFilename, "c:\junk.old"

End Sub

Private Sub cmdRecycleFile_Click()

  RemoveFile txtFilename.Text

End Sub

Private Sub cmdSetAssoc_Click()

  Dim clsFileAssoc As New CFileAssociation
  
  With clsFileAssoc
     .strExt = txtExt
     .strAppID = txtID
     .strOpenCommand = txtOpen
     .strExePath = App.Path & "\" & App.EXEName & ".exe"
     .strFileType = txtFileLabel
     ' .strIcon = "0"
     .strNewFileType = "NullFile"
  End With
  
  clsFileAssoc.CreateAssociation
  clsFileAssoc.CreateContextMenuItem "foobar", "C:\windows\notepad.exe %1"
  
End Sub











Private Sub Form_Load()

   cboShellFolderName.AddItem "Personal"
   cboShellFolderName.AddItem "Desktop"
   cboShellFolderName.AddItem "NetHood"
   cboShellFolderName.AddItem "Programs"
   cboShellFolderName.AddItem "Start Menu"
   cboShellFolderName.AddItem "StartUp"
   cboShellFolderName.AddItem "Favorites"
   cboShellFolderName.AddItem "Fonts"
   cboShellFolderName.AddItem "Recent"
   cboShellFolderName.AddItem "Sendto"
   cboShellFolderName.AddItem "Templates"
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Select Case UnloadMode
        
    Case vbAppWindows
         ' If the app is being closed because Windows is being closed,
         ' we write the registry settings that will cause
         ' the program to start up again when Windows starts up.
         Dim lResult   As Long
         Dim hKey As Long
         Dim strRunCmd As String
         
         strRunCmd = App.Path & "\" & App.EXEName & ".exe"
         
         lResult = RegCreateKey(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\RunOnce", hKey)
         lResult = RegSetValueEx(hKey, App.EXEName, 0&, REG_SZ, ByVal strRunCmd, Len(strRunCmd))
         lResult = RegCloseKey(hKey)
         
    Case Else
       ' Just fall through
       
  End Select
  
End Sub

