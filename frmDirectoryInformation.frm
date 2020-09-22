VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDirectoryInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Information"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   Icon            =   "frmDirectoryInformation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&Info"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2580
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CDlg1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "&File"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2580
      Width           =   700
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "C:\DirInfo.txt"
      Top             =   3000
      Width           =   5415
   End
   Begin VB.CheckBox chkSubDirectory 
      Caption         =   "Include &Subdirectories"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2580
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Hidden          =   -1  'True
      Left            =   3360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   300
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   3255
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2580
      Width           =   700
   End
   Begin VB.Label Label2 
      Caption         =   "File Name"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3045
      Width           =   735
   End
   Begin VB.Label lblPath 
      Alignment       =   2  'Center
      Caption         =   "lblPath"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   50
      Width           =   6375
   End
End
Attribute VB_Name = "frmDirectoryInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Written by Bob Iadicicco

'  Something I needed to turn in Source Code at work.
'  I found the basic idea and code at www.planet-source-code.com
'  submitted by Mick Collins and got carried away with it.
'  If you dbl-click on the label File Name you will get a dialog
'  box for the File Name Field.  It was all written in one form
'  so I could add it to any other projects I might want it in.

'  E-Mail me with any comments at LostKender@HotMail.com

Dim InfoLine() As String
Dim ILC As Long
Dim DirLine() As String
Dim DLC As Long
Dim SubToCheck() As String
Dim STC As Long

Private Sub cmdInfo_Click()
'Info Box
    Dim sInfo As String
    sInfo = "Something I needed to turn in Source Code at work."
    sInfo = sInfo + vbCrLf + vbCrLf
    sInfo = sInfo + "I found the basic idea and code at www.planet-source-code.com"
    sInfo = sInfo + " submitted by Mick Collins and got carried away with it."
    sInfo = sInfo + vbCrLf + vbCrLf
    sInfo = sInfo + "If you dbl-click on the label File Name you will get a dialog"
    sInfo = sInfo + " box for the File Name Field."
    sInfo = sInfo + vbCrLf + vbCrLf
    sInfo = sInfo + "It was all written in one form so I could add it to any other projects I might want it in."
    sInfo = sInfo + vbCrLf + vbCrLf
    sInfo = sInfo + " E-Mail me with any comments at LostKender@HotMail.com"
    MsgBox sInfo, vbInformation, "Lost Kender Products"
End Sub

Private Sub cmdPrint_Click()
'Print Button Command
    Screen.MousePointer = 11
    GetInfoLine
    PF "P"
    Unload Me
    Screen.MousePointer = 0
End Sub

Private Sub cmdFile_Click()
'File Button Command
    Screen.MousePointer = 11
    GetInfoLine
    PF "F"
    Unload Me
    Screen.MousePointer = 0
End Sub

Private Sub GetInfoLine()
'Get Directory Information
    Dim Counter As Long
    Dim SubCheck As String
    Dim STCCounter
    ILC = 0
    STC = 0
    STCCounter = 0
'Get chosen directory info
    GetDirInfo Dir1.Path
'If sub-directories is chosen, get that info too . . .
    If chkSubDirectory Then
        GetNestedInfo Dir1.Path
        If STC > 0 Then
Again:
            STCCounter = STCCounter + 1
            SubCheck = SubToCheck(STCCounter)
            GetNestedInfo SubCheck
            If STCCounter <> STC Then GoTo Again
        End If
    End If
End Sub

Private Sub GetNestedInfo(sPath As String)
'Check for Nested Directories and Save them
    Dim Counter As Long
    Dir1.Path = sPath
    For Counter = 0 To Dir1.ListCount - 1
        GetDirInfo Dir1.List(Counter)
        If Dir1.ListCount > 0 Then
            STC = STC + 1
            ReDim Preserve SubToCheck(STC)
            SubToCheck(STC) = Dir1.List(Counter)
        End If
    Next
End Sub

Private Sub GetDirInfo(sPath As String)
'Get files from Directory
    Dim Counter As Long
    Dim TitleFix As Long
'Title
    ReDim Preserve InfoLine(ILC + 2)
    InfoLine(ILC + 1) = "Directory of:  " & UCase(sPath)
    InfoLine(ILC + 2) = "Date      Time               Size  File"
    ILC = ILC + 2
    TitleFix = ILC
'Files
    File1.Path = sPath
    For Counter = 0 To File1.ListCount - 1
        If Right(sPath, 1) = "\" Then
            ILC = ILC + 1
            ReDim Preserve InfoLine(ILC)
            InfoLine(ILC) = LineInfo(sPath + File1.List(Counter)) + Space(2) + UCase(File1.List(Counter))
        Else
            ILC = ILC + 1
            ReDim Preserve InfoLine(ILC)
            InfoLine(ILC) = LineInfo(sPath + "\" + File1.List(Counter)) + Space(2) + UCase(File1.List(Counter))
        End If
    Next
'If no files were added, fix Title
    If TitleFix = ILC Then InfoLine(ILC) = "No files"
'Skip a line
    ILC = ILC + 1
    ReDim Preserve InfoLine(ILC)
    InfoLine(ILC) = ""
End Sub

Private Sub PF(PrFl As String)
'Send Information to Print or File
    Dim Counter As Long
    Dim fs, a

    If UCase(PrFl) = "F" Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(txtFileName.Text, True)
    End If
    
    For Counter = 0 To ILC
        If UCase(PrFl) = "F" Then
            a.WriteLine (InfoLine(Counter))
        Else
            Printer.Print InfoLine(Counter)
        End If
    Next

    If UCase(PrFl) = "F" Then
        a.Close
    Else
        Printer.EndDoc
    End If
End Sub

Private Function LineInfo(fName As String)
'Get more information about each file
    Dim nLength As Long
    Dim sSpaces As Long
    Dim NewEnt As String
    Dim DateFix As String
    Dim NewDate As String
'Add File Date
    DateFix = Str(FileDateTime(fName))
    If Mid(DateFix, 2, 1) = "/" Then
        'm1
        NewDate = "0" + Mid(DateFix, 1, 2)
        If Mid(DateFix, 4, 1) = "/" Then
            'm1 d1
            NewDate = NewDate + "0" + Mid(DateFix, 3, 4) + Space(2)
            If Len(DateFix) < 9 Then
                'm1 d1 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 9, 1) = ":" Then
                'm1 d1 h1
                NewDate = NewDate + "0" + Mid(DateFix, 8, 17)
            Else
                'm1 d1 h2
                NewDate = NewDate + Mid(DateFix, 8, 18)
            End If
        Else
            'm1 d2
            NewDate = NewDate + Mid(DateFix, 3, 5) + Space(2)
            If Len(DateFix) < 9 Then
                'm1 d2 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 10, 1) = ":" Then
                'm1 d2 h1
                NewDate = NewDate + "0" + Mid(DateFix, 9, 17)
            Else
                'm1 d2 h2
                NewDate = NewDate + Mid(DateFix, 9, 18)
            End If
        End If
    Else
        'm2
        NewDate = Mid(DateFix, 1, 3)
        If Mid(DateFix, 5, 1) = "/" Then
            'm2 d1
            NewDate = NewDate + "0" + Mid(DateFix, 4, 4) + Space(2)
            If Len(DateFix) < 9 Then
                'm2 d1 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 10, 1) = ":" Then
                'm2 d1 h1
                NewDate = NewDate + "0" + Mid(DateFix, 9, 17)
            Else
                'm2 d1 h2
                NewDate = NewDate + Mid(DateFix, 9, 18)
            End If
        Else
            'm2 d2
            NewDate = NewDate + Mid(DateFix, 4, 5) + Space(2)
            If Len(DateFix) < 9 Then
                'm2 d2 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 11, 1) = ":" Then
                'm2 d2 h1
                NewDate = NewDate + "0" + Mid(DateFix, 10, 17)
            Else
                'm2 d2 h2
                NewDate = NewDate + Mid(DateFix, 10, 18)
            End If
        End If
    End If
    LineInfo = NewDate + Space(2)
'Add File Length
    nLength = Len(Str(FileLen(fName)))
    sSpaces = 10 - nLength
    LineInfo = LineInfo + Space(sSpaces) + Str(FileLen(fName))
End Function

Private Sub Drive1_Change()
'Update Directory Box with Drive Information
    Dim fs, d, s, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(Left(Drive1.Drive, 2) + "\")
    If d.IsReady Then
        Dir1.Path = Left(Drive1.Drive, 2) + "\"
    Else
        MsgBox "Drive not Available", vbExclamation, "Directory Information"
        Drive1.Drive = Left(lblPath.Caption, 2)
    End If
End Sub

Private Sub Dir1_Change()
'Update File box and Path Caption with Directory Indormation
    lblPath.Caption = Dir1.Path
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
'Start the hole thing
    Dir1.Path = "C:\"
    lblPath.Caption = Dir1.Path
End Sub

Private Sub Label2_DblClick()
'Display Dialog Box for the path of the File to save the
' directory information too
    With CDlg1
        .DialogTitle = "DirInfo.txt Save Path"
        .CancelError = False
        .FileName = "DirInfo.txt"
        .Flags = cdlOFNHideReadOnly + cdlOFNShareAware
        .InitDir = "C:\"
        .Filter = "*.TXT|*.txt"
        .ShowSave
        txtFileName = .FileName
    End With
End Sub
