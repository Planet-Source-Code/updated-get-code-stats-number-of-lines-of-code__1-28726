VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
 ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) _
 As Long

Public Function File(hwnd As Long, strFile As String, Optional Parameters As String)
  
  If Parameters <> "" Then
    File = ShellExecute(hwnd, "Open", strFile, Parameters, "", 1)
  Else
    File = ShellExecute(hwnd, "Open", strFile, "", "", 1)
  End If
  
End Function

Public Function Web(hwnd As Long, strWebAddress As String)
  
  Web = ShellExecute(hwnd, "Open", strWebAddress, "", "", 1)
  
End Function

Public Function Email(hwnd As Long, strEmail As String)
  
  Email = ShellExecute(hwnd, "Open", ("mailto:" & strEmail), "", "", 1)
  
End Function



