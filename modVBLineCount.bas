Attribute VB_Name = "modVBLineCount"
Option Explicit

Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Global cCode As Long, cComments As Long
Global cBlank As Long, cTotal As Long

Global cForms As Long, cModules As Long
Global cClasses As Long

Global cReferences As Long, cObjects As Long
Global cResources As Long

Global cProcedures As Long, cFunctions As Long

Global maxWdth As Long

Public Function GetLineCount(ByVal File As String, ByRef CodeCount As Long, CommentCount As Long, BlankCount As Long, ProcedureCount As Long, FunctionCount As Long)

Dim fName As String, fNum As Integer
Dim strData As String, aFound As Boolean

  CodeCount = 0
  CommentCount = 0
  BlankCount = 0
  ProcedureCount = 0
  FunctionCount = 0
  
  fName = File
  fNum = FreeFile
  
  If fName = "" Then
    MsgBox "Invalid File Name!", vbCritical, "Error"
    Exit Function
  End If
  
  If Left(fName, 10) = "zResFile32" Then
  ElseIf Left(fName, 7) = "zObject" Then
  ElseIf Left(fName, 10) = "zReference" Then
  Else
  
    Open fName For Input As fNum
    
    aFound = False
    
    Do Until EOF(fNum)
    
      Line Input #fNum, strData
      
      If Left(strData, 9) = "Attribute" And aFound = False Then
        aFound = True
      ElseIf Left(strData, 9) <> "Attribute" And aFound = True Then
              
        Call StripBeginingSpaces(strData)
        If strData = "" Then
          BlankCount = BlankCount + 1
        ElseIf Left(strData, 1) = "'" Then
          CommentCount = CommentCount + 1
        Else
          CodeCount = CodeCount + 1
        End If
        
        If Left(strData, 11) = "Private Sub" Then
          ProcedureCount = ProcedureCount + 1
        ElseIf Left(strData, 10) = "Public Sub" Then
          ProcedureCount = ProcedureCount + 1
        ElseIf Left(strData, 10) = "Friend Sub" Then
          ProcedureCount = ProcedureCount + 1
        ElseIf Left(strData, 10) = "Static Sub" Then
          ProcedureCount = ProcedureCount + 1
        ElseIf Left(strData, 3) = "Sub" Then
          ProcedureCount = ProcedureCount + 1
        ElseIf Left(strData, 16) = "Private Function" Then
          FunctionCount = FunctionCount + 1
        ElseIf Left(strData, 15) = "Public Function" Then
          FunctionCount = FunctionCount + 1
        ElseIf Left(strData, 15) = "Friend Function" Then
          FunctionCount = FunctionCount + 1
        ElseIf Left(strData, 15) = "Static Function" Then
          FunctionCount = FunctionCount + 1
        ElseIf Left(strData, 8) = "Function" Then
          FunctionCount = FunctionCount + 1
        End If
        
      End If
      
    Loop
    
    Close #fNum
  
  End If

End Function

Public Function StripBeginingSpaces(ByRef strData As String)

  Do Until Left(strData, 1) <> " "
    strData = Right(strData, Len(strData) - 1)
  Loop

End Function

Public Function GetFilePath(ByVal Data As String, ByVal FilePath As String)

Dim fPath As String, strData As String
Dim x As Integer, fName As String
Dim fDir As String, oas As New OpenSaveDialog
Dim tmpFile As String, tmpPath As String

  strData = Data
  fPath = FilePath
  
  If InStr(1, strData, "\") = 0 And InStr(1, strData, ";") = 0 Then
    
    x = InStr(1, strData, "=")
    GetFilePath = fPath & Right(strData, Len(strData) - x)
    
  ElseIf InStr(1, strData, ":\") <> 0 Then 'other drive
    
    x = InStr(1, strData, "=")
    GetFilePath = Right(strData, Len(strData) - x)
  
  ElseIf InStr(1, strData, "\\") <> 0 Then 'network share
  
    x = InStr(1, strData, "\\")
    GetFilePath = Right(strData, Len(strData) - (x - 1))
  
  ElseIf InStr(1, strData, "..\") <> 0 Then 'Remove the ..\ and reformat the path
    
    x = InStr(1, strData, ";")
    tmpFile = Right(strData, Len(strData) - (x + 1))
    tmpPath = fPath
    x = InStr(1, tmpFile, "..\")
    Do Until x = 0
      tmpFile = Right(tmpFile, Len(tmpFile) - 3)
      x = InStrRev(tmpPath, "\")
      If x = Len(tmpPath) Then
        tmpPath = Left(tmpPath, x - 1)
        x = InStrRev(tmpPath, "\")
      End If
      tmpPath = Left(tmpPath, x)
      
      x = InStr(1, tmpFile, "..\")
    Loop
    
    GetFilePath = tmpPath & tmpFile
  
  Else
  
    x = InStr(1, strData, ";")
    GetFilePath = fPath & Right(strData, Len(strData) - (x + 1))
            
  End If
  
  If frmMain.TextWidth(GetFilePath) > maxWdth Then
    maxWdth = frmMain.TextWidth(GetFilePath)
  End If
  
End Function

Public Sub SetHorizontalBar(FormName As Form, ListBoxName As ListBox)

  maxWdth = maxWdth + FormName.TextWidth("  ")
  maxWdth = maxWdth / Screen.TwipsPerPixelX
  'The API call to add the horizontal scrollbar
  SendMessageByNum ListBoxName.hwnd, LB_SETHORIZONTALEXTENT, maxWdth, 0
  
End Sub
