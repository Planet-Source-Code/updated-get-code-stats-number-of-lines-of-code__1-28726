VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Line Counter - Beta v0.3"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   630
   ClientWidth     =   12075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Double-Click on File for Individual Count"
      Height          =   3135
      Left            =   6480
      TabIndex        =   33
      Top             =   120
      Width           =   5295
      Begin VB.ListBox lstFiles 
         Height          =   2595
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Include Object Counts:"
      Height          =   1095
      Left            =   3360
      TabIndex        =   29
      ToolTipText     =   "For Project Searches Only!"
      Top             =   2160
      Width           =   2895
      Begin VB.CheckBox chkResources 
         Caption         =   "Resource Files"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkObjects 
         Caption         =   "Objects"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkReferences 
         Caption         =   "References"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Get Counts For:"
      Height          =   1095
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "For Directory and Project Searches Only!"
      Top             =   2160
      Width           =   2895
      Begin VB.CheckBox chkClassModules 
         Caption         =   "Class Modules"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkModules 
         Caption         =   "Modules"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkForms 
         Caption         =   "Forms"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Counts"
      Height          =   495
      Left            =   5430
      TabIndex        =   5
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6870
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Get Line Count"
      Default         =   -1  'True
      Height          =   495
      Left            =   3990
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results:"
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   11535
      Begin VB.Label lblProcedureTotals 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   50
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "Totals:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblFunctions 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   48
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "Functions:"
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblProcedures 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   6600
         TabIndex        =   46
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label15 
         Caption         =   "Procedures:"
         Height          =   255
         Left            =   5400
         TabIndex        =   45
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblObjectTotals 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   44
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblResources 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   43
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblObjects 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   42
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblReferences 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   41
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "Totals:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   40
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Resources:"
         Height          =   255
         Left            =   2640
         TabIndex        =   39
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Objects:"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "References:"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblFileTotals 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label10 
         Caption         =   "Totals:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblClasses 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblModules 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblForms 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label9 
         Caption         =   "Forms:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Modules:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Classes:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   18
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblBlank 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   9840
         TabIndex        =   17
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblComments 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   9840
         TabIndex        =   16
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   9840
         TabIndex        =   15
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Total Lines:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8160
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Blank Lines:"
         Height          =   255
         Left            =   8160
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Lines of Comments:"
         Height          =   255
         Left            =   8160
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Lines of Code:"
         Height          =   255
         Left            =   8160
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox chkSub 
         Caption         =   "Search SubDirectories"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   5280
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtFile 
         DragIcon        =   "frmMain.frx":0442
         Height          =   285
         Left            =   240
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
      End
      Begin VB.ComboBox cmbMethod 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File or Directory Location:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Select the type of file(s) you want to get a line count for:"
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reg As New clsRegistry
Dim SubDirs As clsSubDirs

Private Sub cmbMethod_Click()

  If cmbMethod.ListIndex = 0 Then
    chkSub.Enabled = True
  Else
    chkSub.Enabled = False
  End If

End Sub

Private Sub cmdBrowse_Click()

Dim oas As New OpenSaveDialog

  Select Case cmbMethod.ListIndex
    Case 0
      txtFile.Text = BrowseForFolder(Me.hwnd, "Select the foder you want to total all the lines in:")
    Case 1
      txtFile.Text = oas.OpenDialogBox(frmMain, fCustom, , , , "VB Forms (*.frm)" + Chr$(0) + "*.frm" + Chr$(0) + "VB Modules (*.bas)" + Chr$(0) + "*.bas" + Chr$(0) + "VB Class Modules (*.cls)" + Chr$(0) + "*.cls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    Case 2
      txtFile.Text = oas.OpenDialogBox(frmMain, fCustom, , , , "VB Projects (*.vbp)" + Chr$(0) + "*.vbp" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    Case 3
      txtFile.Text = oas.OpenDialogBox(frmMain, fText)
    Case Else
      MsgBox "You must first select a method to browse by", vbCritical, "Duh"
  End Select
  
End Sub

Private Sub cmdExit_Click()
  
  Unload Me
  
End Sub

Private Sub cmdOK_Click()

Dim fName As String, fPath As String
Dim fFile As String, lCode As Long
Dim lComments As Long, lBlank As Long
Dim fNum As Integer, strData As String
Dim pName As String, pMajor As String
Dim pMinor As String, pRev As String
Dim x As Integer
Static FolderList As Collection
Dim lngRef As Long, lngRes As Long, lngObj As Long

  'Make sure that the top half is filled in
  If cmbMethod.Text = "" Or txtFile.Text = "" Then
    MsgBox "You must select a method to search by and a file to search!", vbCritical, "Dumb Ass"
    Exit Sub
  End If
  
  Call ResetCounts(True)
  cmdOK.Enabled = False
  VB.Screen.MousePointer = 11
  
  'Reset counters
  cCode = 0
  cComments = 0
  cBlank = 0
  cTotal = 0
  cForms = 0
  cModules = 0
  cClasses = 0
  
  frameResults.Caption = "Results:"
  lstFiles.AddItem "All Files"
  
  If cmbMethod.ListIndex = 0 Then 'Directory
  
    If InStr(1, txtFile.Text, ".") Then
      'Not a folder
      txtFile.Text = ""
      lstFiles.Clear
      MsgBox "The file you have selected is not a Folder.  Please re-select and try again.", vbCritical, "Jack Ass"
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
    
    If Left(txtFile.Text, 1) <> "\" Then
      fPath = txtFile.Text & "\"
    Else
      fPath = txtFile.Text
    End If
    
    If chkSub.Value = 1 Then
      If chkForms.Value = 1 Then Call SubDirs.lstBoxGetAllFiles(fPath, "frm", frmMain, lstFiles)
      If chkModules.Value = 1 Then Call SubDirs.lstBoxGetAllFiles(fPath, "bas", frmMain, lstFiles, , False)
      If chkClassModules.Value = 1 Then Call SubDirs.lstBoxGetAllFiles(fPath, "cls", frmMain, lstFiles, , False)
      Call SubDirs.SetHorizontalBar(frmMain, lstFiles)
    Else
      If chkForms.Value = 1 Then Call SubDirs.lstBoxGetCurrentDir(fPath, "frm", frmMain, lstFiles)
      If chkModules.Value = 1 Then Call SubDirs.lstBoxGetCurrentDir(fPath, "bas", frmMain, lstFiles, , False)
      If chkClassModules.Value = 1 Then Call SubDirs.lstBoxGetCurrentDir(fPath, "cls", frmMain, lstFiles, , False)
      Call SubDirs.SetHorizontalBar(frmMain, lstFiles)
    End If
    
    If lstFiles.ListCount = 1 Then
      MsgBox "There were no VB Forms, Modules or Class Modules found in this directory.  Please re-select and try again.", vbInformation, "No files found"
      Call ResetCounts(True)
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
  
  ElseIf cmbMethod.ListIndex = 1 Then 'Single File
  
    If Right(txtFile.Text, 3) = "frm" Then
    ElseIf Right(txtFile.Text, 3) = "bas" Then
    ElseIf Right(txtFile.Text, 3) = "cls" Then
    Else
      'Not a valid file
      txtFile.Text = ""
      lstFiles.Clear
      MsgBox "The file you have selected is not a Visual Basic Form, Module, or Class Module.  Please re-select and try again.", vbCritical, "Bozo"
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
    
    'Add file to listbox
    lstFiles.AddItem txtFile.Text
      
  ElseIf cmbMethod.ListIndex = 2 Then 'VB Project File
    
    maxWdth = 0
    
    fName = txtFile.Text
    If Right(fName, 3) <> "vbp" Then
      'Not a valid file
      txtFile.Text = ""
      lstFiles.Clear
      MsgBox "The file you have selected is not a Visual Basic Project file.  Please re-select and try again.", vbCritical, "Wake Up"
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
    
    x = InStrRev(fName, "\")
    fPath = Left(fName, x)
    
    fNum = FreeFile
    
    Open fName For Input As fNum
    
      Do Until EOF(fNum)
        
        Line Input #fNum, strData
        
        If Left(strData, 5) = "Name=" Then
          pName = Mid(strData, 7)
          pName = Left(pName, Len(pName) - 1)
        ElseIf Left(strData, 9) = "MajorVer=" Then
          pMajor = Mid(strData, 10)
        ElseIf Left(strData, 9) = "MinorVer=" Then
          pMinor = Mid(strData, 10)
        ElseIf Left(strData, 12) = "RevisionVer=" Then
          pRev = Mid(strData, 13)
        ElseIf Left(strData, 5) = "Form=" And chkForms.Value = 1 Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
            
        ElseIf Left(strData, 7) = "Module=" And chkModules.Value = 1 Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
          
        ElseIf Left(strData, 6) = "Class=" And chkClassModules.Value = 1 Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
          
        ElseIf Left(strData, 10) = "Reference=" And chkReferences.Value = 1 Then
        
          x = InStrRev(strData, "#")
          lstFiles.AddItem "zReference = " & Right(strData, Len(strData) - x)
          cReferences = cReferences + 1
          
        ElseIf Left(strData, 7) = "Object=" And chkObjects.Value = 1 Then
        
          x = InStrRev(strData, ";")
          lstFiles.AddItem "zObject = " & Right(strData, Len(strData) - x)
          cObjects = cObjects + 1
        
        ElseIf Left(strData, 10) = "ResFile32=" And chkResources.Value = 1 Then
        
          lstFiles.AddItem "z" & strData
          cResources = cResources + 1
        
        End If
        
      Loop
      
    Close #fNum
    
    SetHorizontalBar frmMain, lstFiles
    lngRef = cReferences
    lngRes = cResources
    lngObj = cObjects
    
  ElseIf cmbMethod.ListIndex = 3 Then 'Other
  
    fNum = FreeFile
    Open txtFile.Text For Input As fNum
  
    Do Until EOF(fNum)
    
      Line Input #fNum, strData
      
      Call StripBeginingSpaces(strData)
      If strData = "" Then
        cBlank = cBlank + 1
      Else
        cCode = cCode + 1
      End If
    
    Loop
    
    Close #fNum
    
    lstFiles.Clear
    lblForms.Caption = cForms
    lblModules.Caption = cModules
    lblClasses.Caption = cClasses
    lblCode.Caption = cCode
    lblComments.Caption = cComments
    lblBlank.Caption = cBlank
    lblTotal.Caption = cCode + cComments + cBlank
    
    GoTo Other
    
  End If
    
  lstFiles.Refresh
  lstFiles.Selected(0) = True
  Call AnalyzeFile(lstFiles.Text)
  
  lblResources.Caption = lngRes
  lblReferences.Caption = lngRef
  lblObjects.Caption = lngObj
  lblObjectTotals.Caption = (lngRes + lngRef + lngObj)
    
  If cmbMethod.ListIndex = 2 Then frameResults.Caption = "Results: " & pName & " " & pMajor & "." & pMinor & "." & pRev
  
Other:
  
  VB.Screen.MousePointer = 0
  cmdOK.Enabled = True
  
  reg.SaveSettingString Local_Machine, "Software\Rossi\VBLineCounter", "Method", cmbMethod.Text
  reg.SaveSettingString Local_Machine, "Software\Rossi\VBLineCounter", "File", txtFile.Text
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "CheckSubs", chkSub.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkForms", chkForms.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkModules", chkModules.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkClassModules", chkClassModules.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkReferences", chkReferences.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkObjects", chkObjects.Value
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "chkResources", chkResources.Value
  
End Sub

Private Sub cmdReset_Click()

  Call ResetCounts(True)

End Sub

Private Sub Form_Load()

  With cmbMethod
    .AddItem "Directory", 0
    .AddItem "Single VB Item (Form, Class, Module)", 1
    .AddItem "Visual Basic Project (.vbp)", 2
    .AddItem "Other", 3
  End With
  
  cmbMethod.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\VBLineCounter", "Method", "Visual Basic Project (.vbp)")
  txtFile.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\VBLineCounter", "File")
  chkSub.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "CheckSubs", 1)
  chkForms.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkForms", 1)
  chkModules.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkModules", 1)
  chkClassModules.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkClassModules", 1)
  chkReferences.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkReferences", 1)
  chkObjects.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkObjects", 1)
  chkResources.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "chkResources", 1)
  
  Me.Caption = "VB Line Counter - v" & App.Major & "." & App.Minor & "." & App.Revision
  
  Set SubDirs = New clsSubDirs
  
End Sub

Public Sub ResetCounts(Optional ClearListBox As Boolean = False)
  
  If ClearListBox = True Then lstFiles.Clear
  
  'Reset counters
  cCode = 0
  cComments = 0
  cBlank = 0
  cTotal = 0
  cForms = 0
  cModules = 0
  cClasses = 0
  cResources = 0
  cReferences = 0
  cObjects = 0
  cProcedures = 0
  cFunctions = 0
  
  'Reset labels
  lblForms.Caption = 0
  lblModules.Caption = 0
  lblClasses.Caption = 0
  lblFileTotals.Caption = 0
  
  lblReferences.Caption = 0
  lblObjects.Caption = 0
  lblResources.Caption = 0
  lblObjectTotals.Caption = 0
  
  lblProcedures.Caption = 0
  lblFunctions.Caption = 0
  lblProcedureTotals.Caption = 0
  
  lblCode.Caption = 0
  lblComments.Caption = 0
  lblBlank.Caption = 0
  lblTotal.Caption = 0
  
  frameResults.Caption = "Results:"
  
End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

  txtFile.Text = Data.Files(1)

End Sub

Private Sub lstFiles_DblClick()

  Call AnalyzeFile(lstFiles.Text)

End Sub

Public Function AnalyzeFile(FileName As String)

Dim fFile As String, lCode As Long
Dim lComments As Long, lBlank As Long
Dim lProcedures As Long, lFunctions As Long
Dim x As Integer, lCount As Integer

  VB.Screen.MousePointer = 11
  
  fFile = FileName
  
  If fFile = "All Files" Then
    
    Call ResetCounts
    
    x = 1
    For x = 1 To lstFiles.ListCount - 1
      
      fFile = lstFiles.List(x)
      
      Call GetLineCount(fFile, lCode, lComments, lBlank, lProcedures, lFunctions)
      cCode = cCode + lCode
      cComments = cComments + lComments
      cBlank = cBlank + lBlank
      cProcedures = cProcedures + lProcedures
      cFunctions = cFunctions + lFunctions
      
      If LCase(Right(fFile, 3)) = "frm" Then
        cForms = cForms + 1
      ElseIf LCase(Right(fFile, 3)) = "bas" Then
        cModules = cModules + 1
      ElseIf LCase(Right(fFile, 3)) = "cls" Then
        cClasses = cClasses + 1
      End If
      
      lblForms.Caption = cForms
      lblModules.Caption = cModules
      lblClasses.Caption = cClasses
      lblFileTotals.Caption = (cForms + cClasses + cModules)
      
      lblProcedures.Caption = cProcedures
      lblFunctions.Caption = cFunctions
      lblProcedureTotals.Caption = (cProcedures + cFunctions)
      
      lblCode.Caption = cCode
      lblComments.Caption = cComments
      lblBlank.Caption = cBlank
      lblTotal.Caption = cCode + cComments + cBlank
      
      DoEvents
      
    Next x
    
  Else
  
    Call ResetCounts
  
    Call GetLineCount(fFile, lCode, lComments, lBlank, lProcedures, lFunctions)
    cCode = cCode + lCode
    cComments = cComments + lComments
    cBlank = cBlank + lBlank
    cProcedures = cProcedures + lProcedures
    cFunctions = cFunctions + lFunctions
    
    If Right(fFile, 3) = "frm" Then
      cForms = cForms + 1
    ElseIf Right(fFile, 3) = "bas" Then
      cModules = cModules + 1
    ElseIf Right(fFile, 3) = "cls" Then
      cClasses = cClasses + 1
    End If
    
    lblForms.Caption = cForms
    lblModules.Caption = cModules
    lblClasses.Caption = cClasses
    lblFileTotals.Caption = (cForms + cClasses + cModules)
    
    lblProcedures.Caption = cProcedures
    lblFunctions.Caption = cFunctions
    lblProcedureTotals.Caption = (cProcedures + cFunctions)
    
    lblCode.Caption = cCode
    lblComments.Caption = cComments
    lblBlank.Caption = cBlank
    lblTotal.Caption = cCode + cComments + cBlank
    
    DoEvents
    
  End If
  
  VB.Screen.MousePointer = 0
  
End Function

Private Sub mnuAbout_Click()

  frmAbout.Show

End Sub

Private Sub mnuExit_Click()

  Unload Me

End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

  txtFile.Text = Data.Files(1)
  
End Sub
