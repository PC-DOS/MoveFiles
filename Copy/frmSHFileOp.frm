VERSION 5.00
Begin VB.Form frmSHFileOp 
   AutoRedraw      =   -1  'True
   Caption         =   "SHFileOperation Demo"
   ClientHeight    =   4695
   ClientLeft      =   1620
   ClientTop       =   1530
   ClientWidth     =   4785
   Icon            =   "frmSHFileOp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   Begin VB.TextBox txTo 
      Height          =   300
      Left            =   165
      TabIndex        =   6
      Text            =   "C:\Test Folder"
      Top             =   480
      Width           =   4470
   End
   Begin VB.TextBox txFrom 
      Height          =   300
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1410
      Width           =   4470
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   165
      TabIndex        =   3
      Top             =   4200
      Width           =   3000
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   165
      TabIndex        =   2
      Top             =   1905
      Width           =   2970
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   3405
      TabIndex        =   1
      Top             =   2535
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   3405
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   11
      X2              =   312
      Y1              =   66
      Y2              =   66
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   9
      X2              =   310
      Y1              =   65
      Y2              =   65
   End
   Begin VB.Label Label1 
      Caption         =   "Copy Files To :"
      Height          =   240
      Index           =   1
      Left            =   165
      TabIndex        =   7
      Top             =   195
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Copy Files From :"
      Height          =   240
      Index           =   0
      Left            =   165
      TabIndex        =   5
      Top             =   1140
      Width           =   1380
   End
End
Attribute VB_Name = "frmSHFileOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
  ' demo project showing how to use the SHFileOperation API function
  ' by Bryan Stafford of New Vision Software® - newvision@imt.net
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.

  ' ###########################################################
  ' see the comments section at the end of this module for
  ' more info about this API function.
  ' ###########################################################

  ' constants used in the wFunc member of the UDT
  ' to determine which file operation to use.  see
  ' the comments section for further explaination
  Private Const FO_MOVE As Long = &H1&
  Private Const FO_COPY As Long = &H2&
  Private Const FO_DELETE As Long = &H3&
  Private Const FO_RENAME As Long = &H4&

  ' flag constants used in the fFlags member of the UDT
  ' to set the behavior of the dialog when it is displayed.
  ' see the comments section for further explaination
  Private Const FOF_CONFIRMMOUSE As Integer = &H2
  Private Const FOF_SILENT As Integer = &H4
  Private Const FOF_RENAMEONCOLLISION As Integer = &H8
  Private Const FOF_NOCONFIRMATION As Integer = &H10
  Private Const FOF_WANTMAPPINGHANDLE As Integer = &H20
  Private Const FOF_CREATEPROGRESSDLG As Integer = &H0
  Private Const FOF_ALLOWUNDO As Integer = &H40
  Private Const FOF_FILESONLY As Integer = &H80
  Private Const FOF_SIMPLEPROGRESS As Integer = &H100
  Private Const FOF_NOCONFIRMMKDIR As Integer = &H200
  Private Const FOF_NOERRORUI As Integer = &H400
  Private Const FOF_NOCOPYSECURITYATTRIBS As Integer = &H800

  ' UDT to hold the file information
  Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
  End Type ' SHFILEOPSTRUCT

  ' variable to hold the last dir copied to for comparison in the recycle test
  Private sLastCopyDir As String
  
  ' declaration for the API functions
  Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, _
                                                  hpvSource As Any, ByVal cbCopy As Long)
                        
  Private Declare Function SHFileOperation& Lib "Shell32.dll" Alias "SHFileOperationA" _
                                                      (lpFileOp As Any)

Private Sub Command1_Click()
  ' Copies the selected files to the destination location.
  ' NOTE 1: For the copy dialog to be displayed the files
  '         must be of sufficient size for the operation
  '         to take more than a couple of seconds.
  '
  ' NOTE 2: There appears to be a bug in the return of the
  '         fAnyOperationsAborted member of the structure.
  '         I have not found any way to get it to return
  '         any value except for zero which should indicate
  '         that the operation was not aborted.  Even if the
  '         user presses the cancel button on the dialog the
  '         value remains zero.  However, the function call
  '         does return an error value if the cancel button
  '         is selected so that we can know that the operation
  '         was not completed.

  ' dimension the variables
  Dim fileop As SHFILEOPSTRUCT
  Dim aFileOp() As Byte, nLenStruct&
  
  ' fill the UDT
  With fileop
    ' the dialog is modal so we need to supply the handle of the
    ' form that we want to be disabled while the dialog is active
    .hWnd = Me.hWnd

    ' supply the operation type
    .wFunc = FO_COPY

    ' The files to copy are separated by a single null (Chr$(0) or
    ' vbNullChar) and terminated by 2 nulls.
    ' single file example:
    '    "c:\File1.txt" & vbNullChar & vbNullChar
    '
    ' multipule file example:
    '    "c:\File1.txt" & vbNullChar & "c:\File2.txt" & vbNullChar & vbNullChar
    ' OR
    '    "c:\*.*" & vbNullChar & vbNullChar
    '
    .pFrom = txFrom & vbNullChar & vbNullChar

    ' save the dir name that we are copying to so that we can compare it to the
    ' selected dir in the recycle test function and issue a warning if the user
    ' is about to send a different dir to the recycle bin.
    sLastCopyDir = txTo
    
    ' the directory or filename(s) to copy to (terminated in 2 nulls).
    .pTo = txTo & vbNullChar & vbNullChar

    ' flags to determine the behavior of the dialog
    .fFlags = FOF_CREATEPROGRESSDLG Or FOF_FILESONLY

    ' if the FOF_SIMPLEPROGRESS flag is specified this member can be filled
    ' with a string to be displayed in the dialog instead of the names of
    ' the files that are being copied.
    '.lpszProgressTitle = "Copying Many Files...." & vbNullChar & vbNullChar
  End With
  
  ' because VB does not handle the alignment in the structure correctly
  ' we need to *fix* it by copying the structure into a byte array where
  ' the bytes may be manipulated to gain the proper alignment
  
  nLenStruct = LenB(fileop)    ' get the length of the struct
  ReDim aFileOp(1 To nLenStruct) ' dimention the byte array
  Call CopyMemory(aFileOp(1), fileop, nLenStruct) ' copy the struct to the array

  ' move the last 12 bytes in the array up 2 to correctly align the data
  Call CopyMemory(aFileOp(19), aFileOp(21), 12)
  
  ' call the function
  If SHFileOperation(aFileOp(1)) Then
    ' if the call returns anything but zero the operation failed
    ' or the user pressed the cancel button.   According to the
    ' documentation Err.LastDllError will hold the error if an
    ' error occurred.
    
    MsgBox "The cancel button was pressed or an error occourred."
    
  Else
    If fileop.fAnyOperationsAborted <> 0 Then
      ' the operation was aborted.
      ' Note: this does not appear to include the user
      '       pressing the 'Cancel' button
    End If
  End If
    
End Sub

Private Sub Command2_Click()
  ' send the files in the target directory to the Recycle Bin

  ' dim the struct
  Dim DelFileOp As SHFILEOPSTRUCT
  
  With DelFileOp
    .hWnd = Me.hWnd
    .wFunc = FO_DELETE
    
    ' Delete the files you just moved.
    ' let's make sure that we are pointing at the same dir
    If (txTo <> sLastCopyDir) And (sLastCopyDir <> "") Then _
      If MsgBox("You are about to send files othere than the ones that " & _
              "you just copied to the Recycle Bin." & vbCrLf & vbCrLf & _
                              "Continue Operation?", vbYesNoCancel) <> 6 Then Exit Sub
    
    sLastCopyDir = txTo
    If (Right$(sLastCopyDir, 1) <> "\") And _
                    (Right$(sLastCopyDir, 4) <> "\*.*") Then sLastCopyDir = sLastCopyDir & "\"
    If (Right$(sLastCopyDir, 3) <> "*.*") Then sLastCopyDir = sLastCopyDir & "*.*"

    ' assign the dir to the struct member
    .pFrom = sLastCopyDir & vbNullChar & vbNullChar
    
    ' Allow undo--in other words, place the files into the
    ' Recycle Bin
    .fFlags = FOF_ALLOWUNDO
    
  End With
  
  ' call the function
  If SHFileOperation(DelFileOp) Then
  
    'if we get here the operation failed
    'show the error that occurred in the call
    MsgBox Err.LastDllError
  
  Else
    If DelFileOp.fAnyOperationsAborted Then MsgBox "Operation Aborted"
  End If

End Sub

Private Sub Dir1_Change()

  txFrom = Dir1.Path

End Sub

Private Sub Drive1_Change()

  Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()

  Command1.Caption = "Copy Test"
  Command2.Caption = "Recycle Test"

  txFrom = Dir1.Path
 
End Sub

Private Sub txFrom_Change()
  
  Static bInSub As Boolean
  
  If bInSub Then Exit Sub
  bInSub = True
  
  If (Right$(txFrom, 1) <> "\") And (Right$(txFrom, 4) <> "\*.*") Then txFrom = txFrom & "\"
  If (Right$(txFrom, 3) <> "*.*") Then txFrom = txFrom & "*.*"
  
  bInSub = False
  
End Sub

'##  COMMENTS  ###############################################################
'
' SHFILEOPSTRUCT Members
'
'hWnd
'      Handle of the dialog box to use to display information about the status
'      of the operation. If fFlags includes the FOF_CREATEPROGRESSDLG value, this
'      parameter is the handle of the parent window for the progress dialog box
'      created by the system.
'
'wFunc
'      Operation to perform. This member can be one of the following values:
'
'      FO_COPY     Copies the files specified by pFrom to the location specified by pTo.
'      FO_DELETE   Deletes the files specified by pFrom (pTo is ignored).
'      FO_MOVE     Moves the files specified by pFrom to the location specified by pTo.
'      FO_RENAME   Renames the files specified by pFrom.
'
'pFrom
'      Pointer to a buffer that specifies one or more source file names. Multiple
'      names must be null-separated. The list of names must be double null-terminated.
'
'pTo
'      Pointer to a buffer that contains the name of the destination file or directory.
'      The buffer can contain mutiple destination file names if the fFlags member
'      specifies FOF_MULTIDESTFILES. Multiple names must be null-separated. The list of
'      names must be double null-terminated.
'
'fFlags
'      Flags that control the file operation. This member can be a combination of the
'      following values:
'
'      FOF_ALLOWUNDO          Preserves undo information, if possible.
'      FOF_CONFIRMMOUSE       Not implemented.
'      FOF_FILESONLY          Performs the operation only on files if a wildcard filename
'                             (*.*) is specified.
'      FOF_MULTIDESTFILES     Indicates that the pTo member specifies multiple destination
'                             files (one for each source file) rather than one directory
'                             where all source files are to be deposited.
'      FOF_NOCONFIRMATION     Responds with "yes to all" for any dialog box that is displayed.
'      FOF_NOCONFIRMMKDIR     Does not confirm the creation of a new directory if the operation
'                             requires one to be created.
'      FOF_RENAMEONCOLLISION  Gives the file being operated on a new name (such as "Copy #1 of...")
'                             in a move, copy, or rename operation if a file of the target name
'                             already exists.
'      FOF_SILENT             Does not display a progress dialog box.
'      FOF_SIMPLEPROGRESS     Displays a progress dialog box, but does not show the filenames.
'      FOF_WANTMAPPINGHANDLE  Fills in the hNameMappings member. The handle must be freed by
'                             using the SHFreeNameMappings function.
'
'fAnyOperationsAborted
'      Value that receives TRUE if the user aborted any file operations before they were completed
'      or FALSE otherwise.
'
'hNameMappings
'      Handle of a filename mapping object that contains an array of SHNAMEMAPPING structures.
'      Each structure contains the old and new path names for each file that was moved, copied,
'      or renamed. This member is used only if fFlags includes FOF_WANTMAPPINGHANDLE.
'
'lpszProgressTitle
'      Pointer to a string to use as the title for a progress dialog box. This member is used
'      only if fFlags includes FOF_SIMPLEPROGRESS.

