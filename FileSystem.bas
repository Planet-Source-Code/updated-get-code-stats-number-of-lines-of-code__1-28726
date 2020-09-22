Attribute VB_Name = "FileSystem"
Option Explicit

'API for the createpath function
Declare Function MakeSureDirectoryPathExists Lib "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

'API for the findfile function
Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, _
    ByVal lpInputName As String, ByVal lpOutputName As String) As Long
    
Public Const MAX_PATH = 260

'Global iError As Integer
Public fs As Object, a
Private iError As Integer

Public Sub CreateFile(fPath As String, Optional WriteLine As String)
'
On Error GoTo CreateError
  '
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set a = fs.CreateTextFile(fPath, True)
  a.WriteLine (WriteLine)
  a.Close
  iError = 0
  Exit Sub
  '
CreateError:
  iError = 1
  '
End Sub

Public Function FileExists(Path As String) As Boolean
'
On Error GoTo CreateError
  '
  Set fs = CreateObject("Scripting.FileSystemObject")
  FileExists = fs.FileExists(Path)
  iError = 0
  '
  Exit Function
  '
CreateError:
  '
  iError = 1
  '
End Function

Public Function FolderExists(Path As String) As Boolean
'
On Error GoTo CreateError
  '
  Set fs = CreateObject("Scripting.FileSystemObject")
  FolderExists = fs.FolderExists(Path)
  iError = 0
  Exit Function
  '
CreateError:
  iError = 1
  '
End Function

Public Sub CreateFolder(Path As String)
'
On Error GoTo CreateError
  '
  Set fs = CreateObject("Scripting.FileSystemObject")
  fs.CreateFolder (Path)
  iError = 0
  Exit Sub
  '
CreateError:
  iError = 1
  '
End Sub

Public Sub FileCopy(CopyFrom As String, CopyTo As String, Optional OverWrite As Boolean = False)

  Set fs = CreateObject("Scripting.FileSystemObject")
  fs.CopyFile CopyFrom, CopyTo, OverWrite

End Sub

Public Function CreatePath(NewPath) As Boolean

  'Add a trailing slash if none
  If Right(NewPath, 1) <> "\" Then
    NewPath = NewPath & "\"
  End If
  
  'Call API
  If MakeSureDirectoryPathExists(NewPath) <> 0 Then
    'No errors, return True
    CreatePath = True
  Else
    CreatePath = False
    End If

End Function

Public Function FindFile(RootPath As String, FileName As String) As String
    
Dim lNullPos As Long
Dim lResult As Long
Dim sBuffer As String
  
  On Error GoTo FileFind_Error
  
  'Allocate buffer
  sBuffer = Space(MAX_PATH * 2)
  
  'Find the file
  lResult = SearchTreeForFile(RootPath, FileName, sBuffer)
  
  'Trim null, if exists
  If lResult Then
    lNullPos = InStr(sBuffer, vbNullChar)
    If Not lNullPos Then
      sBuffer = Left(sBuffer, lNullPos - 1)
    End If
    'Return filename
    FindFile = sBuffer
  Else
    'Nothing found
    FindFile = vbNullString
  End If
  
  Exit Function
  
FileFind_Error:

  FindFile = vbNullString
    
End Function

