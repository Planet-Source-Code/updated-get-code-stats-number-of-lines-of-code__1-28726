VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Originally taken from Rick Meyer at http://pages.about.com/vbmakai/getfiles.htm
'Code Greatly Modified by Mike Rossi

'==========================================
' These are the API declarations needed for
'   the file searching operations
'==========================================
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * 260
  cAlternate        As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'==========================================
' These are the API declarations needed for
'   adding the listbox horizontal scrollbar
'==========================================
Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim maxWdth As Long

Public Sub arrayGetAllFiles(ByVal strDir As String, ByVal Extension As String, ByRef Results() As String, Optional NumFound As Long)
    
Dim fPath$, fName$, fPathName$
Dim hfind&, nameLen%, matchLen%
Dim WFD As WIN32_FIND_DATA
Dim found As Boolean
  
  fPath = strDir
  If Right(fPath, 1) <> "\" Then
    fPath = fPath & "\"
  End If
    
  matchLen = Len(Extension)
  Extension = LCase$(Extension)
  
  'The first API call is FindFirstFile
  hfind = FindFirstFile(fPath & "*", WFD)
  found = (hfind > 0)
  
  Do While found
  
    fName = TrimNull(WFD.cFileName)
    nameLen = Len(fName)
    fPathName = fPath & fName
    
    If fName = "." Or fName = ".." Then

    ElseIf WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then arrayGetAllFiles fPathName, Extension, Results(), NumFound
        
    ElseIf matchLen > nameLen Then
    
    ElseIf LCase$(Right$(fName, matchLen)) = Extension Then
      ReDim Preserve Results(NumFound)
      Results(NumFound) = fPathName
      NumFound = NumFound + 1
    End If
    
    'Subsequent API calls to FindNextFile
    found = FindNextFile(hfind, WFD)
    
  Loop
  
  'Then close the findfile operation
  FindClose hfind
  
End Sub

Public Sub lstBoxGetCurrentDir(ByVal strDir As String, ByVal Extension As String, FormName As Form, ListBoxName As ListBox, Optional NumFound As Long, Optional ResetMaxWidth As Boolean = True)

Dim fName As String
Dim fFile As String

  If ResetMaxWidth = True Then maxWdth = 0

  fName = Dir(strDir & "*." & Extension)
  Do Until fName = ""
    fFile = strDir & fName
    ListBoxName.AddItem fFile
    If FormName.TextWidth(fFile) > maxWdth Then
      maxWdth = FormName.TextWidth(fFile)
    End If
    NumFound = NumFound + 1
    fName = Dir
  Loop

End Sub
Public Sub lstBoxGetAllFiles(ByVal strDir As String, ByVal Extension As String, FormName As Form, ListBoxName As ListBox, Optional NumFound As Long, Optional ResetMaxWidth As Boolean = True)
    
Dim fPath$, fName$, fPathName$
Dim hfind&, nameLen%, matchLen%
Dim WFD As WIN32_FIND_DATA
Dim found As Boolean

  If ResetMaxWidth = True Then maxWdth = 0
  
  fPath = strDir
  If Right(fPath, 1) <> "\" Then
    fPath = fPath & "\"
  End If
    
  matchLen = Len(Extension)
  Extension = LCase$(Extension)
  
  'The first API call is FindFirstFile
  hfind = FindFirstFile(fPath & "*", WFD)
  found = (hfind > 0)
  
  Do While found
  
    fName = TrimNull(WFD.cFileName)
    nameLen = Len(fName)
    fPathName = fPath & fName
    
    If fName = "." Or fName = ".." Then

    ElseIf WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then lstBoxGetAllFiles fPathName, Extension, FormName, ListBoxName, NumFound, False
        
    ElseIf matchLen > nameLen Then
    
    ElseIf LCase$(Right$(fName, matchLen)) = Extension Then
      ListBoxName.AddItem fPathName
      NumFound = NumFound + 1
      If FormName.TextWidth(fPathName) > maxWdth Then
        maxWdth = FormName.TextWidth(fPathName)
      End If
    End If
    
    'Subsequent API calls to FindNextFile
    found = FindNextFile(hfind, WFD)
    
  Loop
  
  'Then close the findfile operation
  FindClose hfind
  
End Sub

Private Function TrimNull(ByVal Item As String) As String
    
Dim pos As Integer
    
  pos = InStr(Item, Chr$(0))
  If pos Then Item = Left$(Item, pos - 1)
    
  TrimNull = Item
  
End Function

Private Sub Class_Initialize()

  maxWdth = 0

End Sub

Public Sub SetHorizontalBar(FormName As Form, ListBoxName As ListBox)

  maxWdth = maxWdth + FormName.TextWidth("  ")
  maxWdth = maxWdth / Screen.TwipsPerPixelX
  'The API call to add the horizontal scrollbar
  SendMessageByNum ListBoxName.hwnd, LB_SETHORIZONTALEXTENT, maxWdth, 0
  
End Sub


