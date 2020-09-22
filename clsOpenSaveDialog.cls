VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
  "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
  
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
  Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Const OFN_ALLOWMULTISELECT = &H200
'Allow the user to select multiple files (Open File dialog box only).
Const OFN_CREATEPROMPT = &H2000
'Prompt if a non-existing file is chosen.
Const OFN_ENABLEHOOK = &H20
'Use the function specified by lpfnHook to process the dialog box's messages.
Const OFN_ENABLESIZING = &H800000
'Windows 98, 2000: Allow the dialog box to be resized. This is selected by default unless a hook function or custom template is specified.
Const OFN_ENABLETEMPLATE = &H40
'Use the dialog box template specifed by hInstance and lpTemplateName.
Const OFN_ENABLETEMPLATEHANDLE = &H80
'Use the preloaded dialog box template specified by hInstance.
Const OFN_EXPLORER = &H80000
'Use Windows Explorer-like additions to the file dialog box. This is selected by default unless a hook function or custom template is specified.
Const OFN_EXTENSIONDIFFERENT = &H400
'The function sets this flag if the user selects a file with an extension different than the one specified by lpstrDefExt.
Const OFN_FILEMUSTEXIST = &H1000
'Only allow the selection of existing files.
Const OFN_HIDEREADONLY = &H4
'Hide the Open As Read Only check box (Open File dialog box only).
Const OFN_LONGNAMES = &H200000
'Have the file dialog use long file names. This is automatically specified unless the Explorer-type extensions are not used.
Const OFN_NOCHANGEDIR = &H8
'Don't change Windows's current directory to match the one chosen in the dialog box.
Const OFN_NODEREFERENCELINKS = &H100000
'If a shortcut file (.lnk or .pif) is chosen, return the shortcut file itself instead of the file or directory it points to.
Const OFN_NOLONGNAMES = &H40000
'Have the file dialog use short (8.3) file names. This is ignored unless a file dialog without Explorer-type extensions are not used.
Const OFN_NONETWORKBUTTON = &H20000
'Hide and disable the Network button in the dialog box.
Const OFN_NOREADONLYRETURN = &H8000
'The function sets this flag if the selected file is not read-only (Open File dialog box only).
Const OFN_NOTESTFILECREATE = &H10000
'Do not create a test file before the box closes. Normally, this check is done to verify that the disk exists, that there is sufficient disk space, etc. However, this check should not be used on a create-nonmodify network share. Specifying this flag prevents this test from being done.
Const OFN_NOVALIDATE = &H100
'Don't check the filename for invalid characters.
Const OFN_OVERWRITEPROMPT = &H2
'Prompt the user if the chosen file already exists (Save File dialog box only).
Const OFN_PATHMUSTEXIST = &H800
'Only allow the selection of existing paths.
Const OFN_READONLY = &H1
'Check the Open As Read Only box. This flag is set after the function call if the box is checked after the user clicks OK.
Const OFN_SHAREAWARE = &H4000
'Ignore any file sharing violations.
Const OFN_SHOWHELP = &H10
'Show the Help button in the dialog box. The button sends the WM_HELP message to the hook function specified by the structure. If no hook function is used, the Help button will do nothing.

Public Enum FilterType
  fText = 0
  fDatabase = 1
  fExcel = 2
  fSystem = 3
  fCustom = 4
End Enum

Public Function OpenDialogBox(FormName As Form, Filter As FilterType, Optional InitialDir As String, Optional DialogTitle As String, Optional AllowMultiSelect As Boolean = False, Optional CustomFilter As String) As String
'
Dim ofn As OpenFileName, a
Dim strInitDir As String, dlgTitle As String
Dim strFilter
  '
  If Not IsEmpty(InitialDir) Then
    strInitDir = InitialDir
  Else
    strInitDir = CurDir
  End If
  '
  If Not IsEmpty(DialogTitle) Then
    dlgTitle = DialogTitle
  Else
    dlgTitle = "Open"
  End If
  '
  If IsEmpty(CustomFilter) And Filter = fCustom Then
    Filter = fText
  Else
    strFilter = CustomFilter
  End If
  '
  If Filter = fText Then
    strFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "List Files (*.lst)" + Chr$(0) + "*.lst" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fDatabase Then
    strFilter = "MSAccess Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fExcel Then
    strFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fSystem Then
    strFilter = "System Files (*.dll)" + Chr$(0) + "*.dll" + Chr$(0) + "ActiveX Files (*.ocx)" + Chr$(0) + "*.ocx" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  End If
  '
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = FormName.hwnd
  ofn.hInstance = App.hInstance
  'ofn.lpstrFilter = "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ofn.lpstrFilter = strFilter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = strInitDir
  ofn.lpstrTitle = dlgTitle
  If AllowMultiSelect = True Then
    ofn.flags = OFN_ALLOWMULTISELECT Or OFN_FILEMUSTEXIST
  Else
    ofn.flags = OFN_FILEMUSTEXIST
  End If
  '
  a = GetOpenFileName(ofn)
  '
  If (a) Then
    'Do the file open stuff here
    OpenDialogBox = StripNull(ofn.lpstrFile)
    '
  Else
    'Cancel was pressed
    OpenDialogBox = ""
    '
  End If
  '
  '
End Function

Public Function SaveDialogBox(FormName As Form, Filter As FilterType, Optional InitialDir As String, Optional DialogTitle As String, Optional OverwritePrompt As Boolean = False, Optional CustomFilter As String, Optional DefaultFileName As String) As String

Dim ofn As OpenFileName, a
Dim strInitDir As String, dlgTitle As String
Dim strFilter As String
  
  If Not IsEmpty(InitialDir) Then
    strInitDir = InitialDir
  Else
    strInitDir = CurDir
  End If
  
  If Not IsEmpty(DialogTitle) Then
    dlgTitle = DialogTitle
  Else
    dlgTitle = "Save As"
  End If
  
  If IsEmpty(CustomFilter) And Filter = fCustom Then
    Filter = fText
  Else
    strFilter = CustomFilter
  End If
  
  If IsEmpty(DefaultFileName) Then
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
  Else
    ofn.lpstrFile = DefaultFileName & Space$(1024) & vbNullChar & vbNullChar
    ofn.nMaxFile = Len(ofn.lpstrFile)
  End If
  
  If Filter = fText Then
    strFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "List Files (*.lst)" + Chr$(0) + "*.lst" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fDatabase Then
    strFilter = "MSAccess Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fExcel Then
    strFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  End If
  '
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = FormName.hwnd
  ofn.hInstance = App.hInstance
  'ofn.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Rich Text Files (*.rtf)" + Chr$(0) + "*.rtf" + Chr$(0)
  ofn.lpstrFilter = strFilter
  'ofn.lpstrFile = Space$(254)
  'ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = strInitDir
  ofn.lpstrTitle = dlgTitle
  If OverwritePrompt = True Then
    ofn.flags = OFN_OVERWRITEPROMPT
  Else
    ofn.flags = 0
  End If
  '
  a = GetSaveFileName(ofn)
  '
  If (a) Then
    'FileSave Stuff Here
    SaveDialogBox = StripNull(ofn.lpstrFile)
  Else
    'Cancel was pressed
    SaveDialogBox = False
  End If
  '
End Function

Public Function StripNull(ByVal strString As String)

Dim NullCharPos As Integer

    NullCharPos = InStr(strString, Chr(0))
    If NullCharPos > 0 Then
      StripNull = Left(strString, NullCharPos - 1)
    Else
      StripNull = strString
    End If

End Function
